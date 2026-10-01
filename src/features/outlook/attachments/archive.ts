import { Inflate } from 'fflate'

const MAX_RAW_BYTES = 4 * 1024 * 1024
const MAX_ENTRY_BYTES = 8 * 1024 * 1024
const MAX_EXPANDED_BYTES = 16 * 1024 * 1024
const MAX_ENTRIES = 256
const INPUT_CHUNK_BYTES = 512

function invalid(): never {
  throw new Error('Attachment archive is invalid, unsupported, or exceeds safety limits')
}

const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
  let crc = value
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
  return crc >>> 0
})

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) crc = (crc >>> 8) ^ (crcTable[(crc ^ byte) & 0xff] ?? 0)
  return (crc ^ 0xffffffff) >>> 0
}

interface Entry {
  name: string
  flags: number
  method: number
  crc: number
  compressed: number
  expanded: number
  localOffset: number
  dataOffset: number
  endOffset: number
}

/**
 * A deliberately narrow ZIP reader for OOXML. It never writes to disk, follows
 * paths, or trusts an archive's advertised uncompressed lengths. Each Inflate
 * push receives at most 512 compressed bytes, bounding transient output as well
 * as retained output even when an entry lies about its expanded size.
 */
export function safeUnzip(bytes: Uint8Array): Map<string, Uint8Array> {
  try {
    return unzipChecked(bytes)
  } catch {
    // Never include archive names or a third-party parser's error in tool errors.
    return invalid()
  }
}

function unzipChecked(bytes: Uint8Array): Map<string, Uint8Array> {
  if (bytes.length < 22 || bytes.length > MAX_RAW_BYTES) invalid()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const u16 = (offset: number) => view.getUint16(offset, true)
  const u32 = (offset: number) => view.getUint32(offset, true)
  let end = -1
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset--) {
    if (u32(offset) === 0x06054b50 && offset + 22 + u16(offset + 20) === bytes.length) {
      end = offset
      break
    }
  }
  if (end < 0 || u16(end + 4) !== 0 || u16(end + 6) !== 0) invalid()
  const count = u16(end + 10)
  const centralSize = u32(end + 12)
  const centralOffset = u32(end + 16)
  if (
    count === 0 ||
    count > MAX_ENTRIES ||
    u16(end + 8) !== count ||
    centralOffset + centralSize !== end ||
    centralOffset >= end
  )
    invalid()

  const entries: Entry[] = []
  const names = new Set<string>()
  const decoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true })
  let offset = centralOffset
  let declaredTotal = 0
  for (let index = 0; index < count; index++) {
    if (offset + 46 > end || u32(offset) !== 0x02014b50) invalid()
    const flags = u16(offset + 8)
    const method = u16(offset + 10)
    const crc = u32(offset + 16)
    const compressed = u32(offset + 20)
    const expanded = u32(offset + 24)
    const nameLength = u16(offset + 28)
    const extraLength = u16(offset + 30)
    const commentLength = u16(offset + 32)
    const localOffset = u32(offset + 42)
    const next = offset + 46 + nameLength + extraLength + commentLength
    // ZIP64, multi-disk, encryption, patched data and unsupported codecs fail
    // closed. Only UTF-8, descriptor and DEFLATE compression-level bits remain.
    if (
      u16(offset + 6) > 20 ||
      (flags & ~0x080e) !== 0 ||
      (method !== 0 && method !== 8) ||
      u16(offset + 34) !== 0 ||
      next > end ||
      !nameLength ||
      nameLength > 1024 ||
      compressed > MAX_RAW_BYTES ||
      expanded > MAX_ENTRY_BYTES ||
      localOffset >= centralOffset ||
      (method === 0 && compressed !== expanded)
    )
      invalid()
    checkExtras(bytes, offset + 46 + nameLength, extraLength)
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLength)
    const name = decoder.decode(nameBytes)
    const segments = (name.endsWith('/') ? name.slice(0, -1) : name).split('/')
    if (
      // eslint-disable-next-line no-control-regex -- ZIP paths must reject control bytes.
      /[\\:\u0000-\u001f\u007f]/u.test(name) ||
      segments.some((part) => part === '' || part === '.' || part === '..') ||
      names.has(name.toLowerCase()) ||
      (name.endsWith('/') && (expanded !== 0 || compressed !== 0 || method !== 0))
    )
      invalid()
    names.add(name.toLowerCase())
    declaredTotal += expanded
    if (declaredTotal > MAX_EXPANDED_BYTES) invalid()

    if (localOffset + 30 > centralOffset || u32(localOffset) !== 0x04034b50) invalid()
    const localNameLength = u16(localOffset + 26)
    const localExtraLength = u16(localOffset + 28)
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength
    if (
      u16(localOffset + 4) > 20 ||
      u16(localOffset + 6) !== flags ||
      u16(localOffset + 8) !== method ||
      localNameLength !== nameLength ||
      dataOffset + compressed > centralOffset ||
      !nameBytes.every((byte, i) => byte === bytes[localOffset + 30 + i])
    )
      invalid()
    checkExtras(bytes, localOffset + 30 + localNameLength, localExtraLength)
    const descriptor = (flags & 8) !== 0
    for (const [fieldOffset, expected] of [
      [14, crc],
      [18, compressed],
      [22, expanded],
    ]) {
      if (fieldOffset === undefined || expected === undefined) invalid()
      const actual = u32(localOffset + fieldOffset)
      if (actual !== expected && !(descriptor && actual === 0)) invalid()
    }
    let endOffset = dataOffset + compressed
    if (descriptor) {
      const start = endOffset + (u32(endOffset) === 0x08074b50 ? 4 : 0)
      if (
        start + 12 > centralOffset ||
        u32(start) !== crc ||
        u32(start + 4) !== compressed ||
        u32(start + 8) !== expanded
      )
        invalid()
      endOffset = start + 12
    }
    entries.push({
      name,
      flags,
      method,
      crc,
      compressed,
      expanded,
      localOffset,
      dataOffset,
      endOffset,
    })
    offset = next
  }
  if (offset !== end) invalid()
  // Disallow overlapping payloads and unindexed local records/preambles.
  const ordered = [...entries].sort((a, b) => a.localOffset - b.localOffset)
  let localEnd = 0
  for (const entry of ordered) {
    if (entry.localOffset !== localEnd) invalid()
    localEnd = entry.endOffset
  }
  if (localEnd !== centralOffset) invalid()

  let actualTotal = 0
  const result = new Map<string, Uint8Array>()
  for (const entry of entries) {
    const compressed = bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressed)
    let expanded: Uint8Array
    if (entry.method === 0) {
      expanded = compressed.slice()
      actualTotal += expanded.length
    } else {
      const chunks: Uint8Array[] = []
      let size = 0
      const inflater = new Inflate((chunk) => {
        size += chunk.length
        actualTotal += chunk.length
        if (size > entry.expanded || size > MAX_ENTRY_BYTES || actualTotal > MAX_EXPANDED_BYTES)
          invalid()
        if (chunk.length) chunks.push(chunk)
      })
      if (compressed.length === 0) invalid()
      for (let start = 0; start < compressed.length; start += INPUT_CHUNK_BYTES) {
        const end = Math.min(start + INPUT_CHUNK_BYTES, compressed.length)
        inflater.push(compressed.subarray(start, end), end === compressed.length)
      }
      if (size !== entry.expanded) invalid()
      expanded = new Uint8Array(size)
      let start = 0
      for (const chunk of chunks) {
        expanded.set(chunk, start)
        start += chunk.length
      }
    }
    if (actualTotal > MAX_EXPANDED_BYTES || crc32(expanded) !== entry.crc) invalid()
    if (!entry.name.endsWith('/')) result.set(entry.name, expanded)
  }
  return result
}

function checkExtras(bytes: Uint8Array, start: number, length: number): void {
  const end = start + length
  if (end > bytes.length) invalid()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  while (start < end) {
    if (start + 4 > end) invalid()
    const tag = view.getUint16(start, true)
    const size = view.getUint16(start + 2, true)
    // ZIP64 and alternate unicode-path names introduce a second interpretation.
    if (tag === 0x0001 || tag === 0x7075 || start + 4 + size > end) invalid()
    start += 4 + size
  }
}
