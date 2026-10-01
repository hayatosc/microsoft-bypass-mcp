import { Zip, ZipDeflate, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { safeUnzip } from './archive.js'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array | undefined) => new TextDecoder().decode(bytes)

function records(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.length)
  const end = bytes.length - 22
  return { view, end, central: view.getUint32(end + 16, true) }
}

function stored() {
  return zipSync({ 'word/document.xml': encode('<document>Safe text</document>') }, { level: 0 })
}

describe('safeUnzip', () => {
  it('reads ordinary stored and deflated OOXML parts in memory', () => {
    for (const level of [0, 6] as const) {
      const bytes = zipSync(
        {
          '[Content_Types].xml': encode('<Types/>'),
          'word/document.xml': encode('<document>text</document>'),
        },
        { level },
      )
      const result = safeUnzip(bytes)
      expect(result.size).toBe(2)
      expect(decode(result.get('word/document.xml'))).toBe('<document>text</document>')
    }
  })

  it('accepts UTF-8 paths and ignores directory entries', () => {
    const bytes = zipSync(
      { 'word/': new Uint8Array(), 'word/日本語.xml': encode('hello') },
      { level: 0 },
    )
    expect([...safeUnzip(bytes).keys()]).toEqual(['word/日本語.xml'])
  })

  it.each([
    '../outside.xml',
    '/absolute.xml',
    'word/../outside.xml',
    'word//file.xml',
    'word\\file.xml',
    'C:/file.xml',
    'word/./file.xml',
  ])('rejects unsafe path %s', (name) => {
    expect(() => safeUnzip(zipSync({ [name]: encode('x') }))).toThrow('safety limits')
  })

  it('rejects duplicate and case-colliding paths', () => {
    expect(() =>
      safeUnzip(zipSync({ 'word/A.xml': encode('x'), 'word/a.xml': encode('y') })),
    ).toThrow('safety limits')
  })

  it('rejects encrypted, unsupported, ZIP64, multi-disk, and overlapping metadata', () => {
    const edits = [
      (view: DataView, central: number) => view.setUint16(central + 8, 1, true),
      (view: DataView, central: number) => view.setUint16(central + 10, 99, true),
      (view: DataView, central: number) => view.setUint32(central + 24, 0xffffffff, true),
      (view: DataView, central: number) => view.setUint16(central + 34, 1, true),
      (view: DataView, central: number) => view.setUint32(central + 42, 1, true),
      (view: DataView, central: number) => view.setUint16(central + 6, 45, true),
    ]
    for (const edit of edits) {
      const bytes = stored()
      const { view, central } = records(bytes)
      edit(view, central)
      expect(() => safeUnzip(bytes)).toThrow('safety limits')
    }
  })

  it('checks actual CRC instead of trusting central and local headers', () => {
    const bytes = stored()
    const { view } = records(bytes)
    const data = 30 + view.getUint16(26, true) + view.getUint16(28, true)
    bytes[data] = (bytes[data] ?? 0) ^ 1
    expect(() => safeUnzip(bytes)).toThrow('safety limits')
  })

  it('rejects dishonest expanded lengths during streaming, including ZIP bombs', () => {
    const bytes = zipSync({ 'word/document.xml': new Uint8Array(12 * 1024 * 1024) }, { level: 9 })
    const { view, central } = records(bytes)
    view.setUint32(22, 1024, true)
    view.setUint32(central + 24, 1024, true)
    expect(() => safeUnzip(bytes)).toThrow('safety limits')
  })

  it('rejects oversized raw files, expanded entries, aggregate bytes, and entry counts', () => {
    expect(() => safeUnzip(new Uint8Array(4 * 1024 * 1024 + 1))).toThrow('safety limits')
    expect(() => safeUnzip(zipSync({ 'large.xml': new Uint8Array(8 * 1024 * 1024 + 1) }))).toThrow(
      'safety limits',
    )
    expect(() =>
      safeUnzip(
        zipSync({
          'a.xml': new Uint8Array(6 * 1024 * 1024),
          'b.xml': new Uint8Array(6 * 1024 * 1024),
          'c.xml': new Uint8Array(6 * 1024 * 1024),
        }),
      ),
    ).toThrow('safety limits')
    const many = Object.fromEntries(
      Array.from({ length: 257 }, (_, index) => [`${index}.xml`, encode('x')]),
    )
    expect(() => safeUnzip(zipSync(many))).toThrow('safety limits')
  })

  it('rejects truncation, central/local disagreement and trailing data', () => {
    const bytes = stored()
    expect(() => safeUnzip(bytes.subarray(0, bytes.length - 1))).toThrow('safety limits')
    const extra = new Uint8Array(bytes.length + 1)
    extra.set(bytes)
    expect(() => safeUnzip(extra)).toThrow('safety limits')
    const { view } = records(bytes)
    view.setUint16(8, 8, true)
    expect(() => safeUnzip(bytes)).toThrow('safety limits')
  })

  it('verifies data descriptors in streaming ZIP archives', () => {
    const chunks: Uint8Array[] = []
    const zip = new Zip((error, chunk) => {
      if (error) throw error
      chunks.push(chunk)
    })
    const file = new ZipDeflate('word/document.xml')
    zip.add(file)
    file.push(encode('<document>streamed</document>'), true)
    zip.end()
    const bytes = new Uint8Array(chunks.reduce((size, chunk) => size + chunk.length, 0))
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.length
    }
    expect(decode(safeUnzip(bytes).get('word/document.xml'))).toBe('<document>streamed</document>')
    const { view, central } = records(bytes)
    const descriptor = central - 16
    view.setUint32(descriptor + 4, 0, true)
    expect(() => safeUnzip(bytes)).toThrow('safety limits')
  })
})
