import { Unzlib } from 'fflate'
import { getResolvedPDFJS } from 'unpdf'

const MAX_RAW_BYTES = 4 * 1024 * 1024
const MAX_STREAM_BYTES = 8 * 1024 * 1024
const MAX_EXPANDED_BYTES = 16 * 1024 * 1024
const MAX_OBJECTS = 10000
const MAX_TOKENS = 100000
const MAX_PAGES = 200
const MAX_READ_PAGES = 10
const MAX_CHARACTERS = 20000
const TIMEOUT_MS = 10000

const UNSUPPORTED = 'PDF is invalid, unsupported, encrypted, or exceeds safety limits'
function invalid(): never {
  throw new Error(UNSUPPORTED)
}

interface Token {
  kind: 'name' | 'number' | 'word' | 'string' | 'symbol'
  value: string
  start: number
}
type Value =
  | number
  | boolean
  | null
  | Value[]
  | { kind: 'name'; value: string }
  | { kind: 'ref'; id: number; generation: number }
  | { kind: 'string' }
  | { kind: 'dict'; entries: Map<string, Value> }
type Dictionary = Extract<Value, { kind: 'dict' }>

function isDictionary(value: Value | undefined): value is Dictionary {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    'kind' in value &&
    value.kind === 'dict'
  )
}
function named(value: Value | undefined, name: string): boolean {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    'kind' in value &&
    value.kind === 'name' &&
    value.value === name
  )
}
function isReference(value: Value | undefined): value is Extract<Value, { kind: 'ref' }> {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    'kind' in value &&
    value.kind === 'ref'
  )
}

function binaryString(bytes: Uint8Array): string {
  const parts: string[] = []
  for (let i = 0; i < bytes.length; i += 4096)
    parts.push(String.fromCharCode(...bytes.subarray(i, i + 4096)))
  return parts.join('')
}
const whitespace = (char: string) =>
  char === '\0' || char === '\t' || char === '\n' || char === '\f' || char === '\r' || char === ' '
const delimiter = (char: string) => !char || whitespace(char) || '()<>[]{}/%'.includes(char)

/** A bounded lexer for a deliberately conservative, classic-xref PDF subset. */
class Lexer {
  cursor = 0
  tokens = 0
  buffered: Token[] = []
  constructor(readonly text: string) {}

  peek(index = 0): Token | undefined {
    while (this.buffered.length <= index) {
      const token = this.scan()
      if (!token) return undefined
      this.buffered.push(token)
    }
    return this.buffered[index]
  }
  next(): Token {
    const token = this.peek()
    if (!token) invalid()
    this.buffered.shift()
    return token
  }
  expect(value: string): void {
    if (this.next().value !== value) invalid()
  }
  integer(): number {
    const token = this.next()
    const value = Number(token.value)
    if (token.kind !== 'number' || !Number.isSafeInteger(value) || value < 0) invalid()
    return value
  }

  private scan(): Token | undefined {
    const text = this.text
    while (this.cursor < text.length) {
      const char = text[this.cursor] ?? ''
      if (whitespace(char)) {
        this.cursor++
        continue
      }
      if (char === '%') {
        while (this.cursor < text.length && !'\r\n'.includes(text[this.cursor] ?? '')) this.cursor++
        continue
      }
      break
    }
    if (this.cursor === text.length) return undefined
    if (++this.tokens > MAX_TOKENS) invalid()
    const start = this.cursor
    const char = text[this.cursor++] ?? ''
    if (char === '(') {
      let depth = 1
      while (depth && this.cursor < text.length) {
        const next = text[this.cursor++]
        if (next === '\\') {
          this.cursor++
          continue
        }
        if (next === '(' && ++depth > 64) invalid()
        if (next === ')') depth--
      }
      if (depth) invalid()
      return { kind: 'string', value: '', start }
    }
    if (char === '<' && text[this.cursor] !== '<') {
      while (this.cursor < text.length && text[this.cursor] !== '>') {
        const next = text[this.cursor++] ?? ''
        if (!/[\da-fA-F]/u.test(next) && !whitespace(next)) invalid()
      }
      if (text[this.cursor++] !== '>') invalid()
      return { kind: 'string', value: '', start }
    }
    if (char === '<' || char === '>') {
      if (text[this.cursor++] !== char) invalid()
      return { kind: 'symbol', value: char + char, start }
    }
    if ('[]'.includes(char)) return { kind: 'symbol', value: char, start }
    if (char === '/') {
      let value = ''
      while (this.cursor < text.length && !delimiter(text[this.cursor] ?? '')) {
        const next = text[this.cursor++] ?? ''
        if (next === '#') {
          const hex = text.slice(this.cursor, this.cursor + 2)
          if (!/^[\da-fA-F]{2}$/u.test(hex)) invalid()
          value += String.fromCharCode(Number.parseInt(hex, 16))
          this.cursor += 2
        } else value += next
        if (value.length > 256) invalid()
      }
      return { kind: 'name', value, start }
    }
    if (delimiter(char)) invalid()
    while (this.cursor < text.length && !delimiter(text[this.cursor] ?? '')) this.cursor++
    const value = text.slice(start, this.cursor)
    if (value.length > 128) invalid()
    return { kind: /^[+-]?(?:\d+\.?\d*|\.\d+)$/u.test(value) ? 'number' : 'word', value, start }
  }
}

// Active content is unnecessary for text extraction. ObjStm/XRef and indirect
// stream encodings are excluded because this preflight must see every stream.
const forbiddenNames = new Set([
  'Encrypt',
  'ObjStm',
  'XRef',
  'XRefStm',
  'Prev',
  'FFilter',
  'FDecodeParms',
  'JavaScript',
  'JS',
  'Launch',
  'RichMedia',
  'XFA',
  'Form',
  'Type3',
])
function parseValue(lexer: Lexer, depth = 0): Value {
  if (depth > 32) invalid()
  const token = lexer.next()
  if (token.kind === 'name') {
    if (forbiddenNames.has(token.value)) invalid()
    return { kind: 'name', value: token.value }
  }
  if (token.kind === 'string') return { kind: 'string' }
  if (token.kind === 'number') {
    const value = Number(token.value)
    if (!Number.isFinite(value) || Math.abs(value) > 1e9) invalid()
    if (lexer.peek()?.kind === 'number' && lexer.peek(1)?.value === 'R') {
      const generation = lexer.integer()
      lexer.expect('R')
      if (!Number.isSafeInteger(value) || value < 1 || value > MAX_OBJECTS || generation !== 0)
        invalid()
      return { kind: 'ref', id: value, generation }
    }
    return value
  }
  if (token.value === 'null') return null
  if (token.value === 'true' || token.value === 'false') return token.value === 'true'
  if (token.value === '[') {
    const values: Value[] = []
    while (lexer.peek()?.value !== ']') {
      if (values.length >= MAX_OBJECTS) invalid()
      values.push(parseValue(lexer, depth + 1))
    }
    lexer.expect(']')
    return values
  }
  if (token.value === '<<') {
    const entries = new Map<string, Value>()
    while (lexer.peek()?.value !== '>>') {
      const key = lexer.next()
      if (key.kind !== 'name' || entries.has(key.value) || forbiddenNames.has(key.value)) invalid()
      const value = parseValue(lexer, depth + 1)
      // These structural tags must stay visible to the preflight checks.
      if ((key.value === 'Type' || key.value === 'Subtype') && isReference(value)) invalid()
      entries.set(key.value, value)
    }
    lexer.expect('>>')
    return { kind: 'dict', entries }
  }
  return invalid()
}

/** Bound actual decoded bytes before PDF.js can allocate a decoded stream. */
interface StreamBudget {
  total: number
  mappings: number
  mappingBytes: number
}

function checkStream(bytes: Uint8Array, dictionary: Dictionary, budget: StreamBudget): number {
  const filter = dictionary.entries.get('Filter')
  const parameters = dictionary.entries.get('DecodeParms')
  if (parameters !== undefined && parameters !== null) invalid()
  if (dictionary.entries.has('F')) invalid() // External stream files are never read.
  const flate =
    named(filter, 'FlateDecode') ||
    (Array.isArray(filter) && filter.length === 1 && named(filter[0], 'FlateDecode'))
  if (filter !== undefined && filter !== null && !flate) invalid()
  let size = 0
  let tail = ''
  const parts: string[] = []
  const inspect = (chunk: Uint8Array) => {
    size += chunk.length
    budget.total += chunk.length
    if (size > MAX_STREAM_BYTES || budget.total > MAX_EXPANDED_BYTES) invalid()
    // Inline images have independent codecs/predictors that evade stream bounds.
    // Reject their operator even in binary/font streams (safe false positives).
    const part = binaryString(chunk)
    parts.push(part)
    const text = tail + part
    // eslint-disable-next-line no-control-regex -- PDF lexical whitespace includes NUL.
    if (/(?:^|[\x00\t\n\f\r ()<>[\]{}/%])BI(?=$|[\x00\t\n\f\r ()<>[\]{}/%])/u.test(text)) invalid()
    tail = text.slice(-3)
  }
  if (!flate) {
    inspect(bytes)
  } else {
    const inflater = new Unzlib(inspect)
    if (bytes.length < 6) invalid()
    for (let start = 0; start < bytes.length; start += 512) {
      const end = Math.min(start + 512, bytes.length)
      inflater.push(bytes.subarray(start, end), end === bytes.length)
    }
  }
  checkCMap(parts.join(''), budget)
  return size
}

/** PDF.js permits huge CMap ranges, so bound mappings independently of bytes. */
function checkCMap(text: string, budget: StreamBudget): void {
  if (!/begin(?:bf|cid)(?:range|char)\b/u.test(text)) return
  const lexer = new Lexer(text)
  const hex = (token: Token, maxLength: number): string => {
    if (token.kind !== 'string' || text[token.start] !== '<') invalid()
    const end = text.indexOf('>', token.start)
    const value = text.slice(token.start + 1, end).replace(/\s/gu, '')
    if (!value || value.length > maxLength || !/^[a-fA-F0-9]+$/u.test(value)) invalid()
    return value
  }
  const code = () => Number.parseInt(hex(lexer.next(), 4), 16)
  const destination = (token: Token) => {
    if (token.kind === 'number') {
      const value = Number(token.value)
      if (!Number.isInteger(value) || value < 0 || value > 65535) invalid()
      return 2
    }
    return Math.ceil(hex(token, 256).length / 2)
  }
  const add = (count: number, bytes: number) => {
    budget.mappings += count
    budget.mappingBytes += bytes
    if (budget.mappings > 65536 || budget.mappingBytes > 1024 * 1024) invalid()
  }
  while (lexer.peek()) {
    const token = lexer.next()
    if (token.kind !== 'word' || !/^begin(?:bf|cid)(?:range|char)$/u.test(token.value)) continue
    const range = token.value.endsWith('range')
    const end = token.value.replace('begin', 'end')
    while (lexer.peek()?.value !== end) {
      const first = code()
      const last = range ? code() : first
      if (last < first) invalid()
      const count = last - first + 1
      const value = lexer.next()
      if (value.value === '[') {
        if (!range) invalid()
        let entries = 0
        let bytes = 0
        while (lexer.peek()?.value !== ']') {
          if (++entries > count) invalid()
          bytes += destination(lexer.next())
        }
        lexer.expect(']')
        add(count, bytes)
      } else add(count, count * destination(value))
    }
    lexer.expect(end)
  }
}

function checkFontWidths(objects: ReadonlyMap<number, { value: Value }>): void {
  let widthCount = 0
  const resolve = (value: Value | undefined, path = new Set<number>()): Value => {
    if (value === undefined || path.size > 32) invalid()
    if (!isReference(value)) return value
    if (path.has(value.id)) invalid()
    return resolve(objects.get(value.id)?.value, new Set([...path, value.id]))
  }
  const number = (value: Value | undefined): number => {
    const resolved = resolve(value)
    if (typeof resolved !== 'number' || !Number.isFinite(resolved)) invalid()
    return resolved
  }
  const code = (value: Value | undefined): number => {
    const cid = number(value)
    if (!Number.isInteger(cid) || cid < 0 || cid > 65535) invalid()
    return cid
  }
  const widths = (value: Value, vertical: boolean) => {
    const array = resolve(value)
    if (!Array.isArray(array)) invalid()
    const unit = vertical ? 3 : 1
    let cursor = 0
    while (cursor < array.length) {
      const start = code(array[cursor++])
      const end = resolve(array[cursor++])
      let count: number
      if (Array.isArray(end)) {
        if (end.length % unit !== 0) invalid()
        count = end.length / unit
        if (start + count > 65536) invalid()
        for (const measurement of end) number(measurement)
      } else {
        const last = code(end)
        if (last < start) invalid()
        count = last - start + 1
        for (let i = 0; i < unit; i++) number(array[cursor++])
      }
      widthCount += count
      if (widthCount > 65536) invalid()
    }
  }
  const visit = (value: Value) => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item)
      return
    }
    if (!isDictionary(value)) return
    for (const [key, item] of value.entries) {
      if (key === 'W' || key === 'W2') widths(item, key === 'W2')
      if (key === 'FirstChar' || key === 'LastChar') code(item)
      if (key === 'Differences') {
        const array = resolve(item)
        if (!Array.isArray(array)) invalid()
        let index = 0
        for (const entry of array) {
          const resolved = resolve(entry)
          if (typeof resolved === 'number') index = code(resolved)
          else {
            if (
              typeof resolved !== 'object' ||
              resolved === null ||
              Array.isArray(resolved) ||
              !('kind' in resolved) ||
              resolved.kind !== 'name' ||
              index++ > 65535
            )
              invalid()
          }
        }
      }
      visit(item)
    }
  }
  for (const object of objects.values()) visit(object.value)
}

function preflight(bytes: Uint8Array): void {
  if (!bytes.length || bytes.length > MAX_RAW_BYTES) invalid()
  const text = binaryString(bytes)
  // eslint-disable-next-line no-control-regex -- PDF lexical whitespace includes NUL.
  if (!/^%PDF-1\.[0-7](?:\r\n|\r|\n)/u.test(text) || !/%%EOF[\x00\t\n\f\r ]*$/u.test(text))
    invalid()
  const lexer = new Lexer(text)
  const objects = new Map<number, number>()
  const values = new Map<number, { value: Value; streamBytes?: number }>()
  const budget = { total: 0, mappings: 0, mappingBytes: 0 }
  while (lexer.peek()?.value !== 'xref') {
    const start = lexer.peek()?.start
    const id = lexer.integer()
    if (
      start === undefined ||
      id < 1 ||
      id > MAX_OBJECTS ||
      objects.has(id) ||
      lexer.integer() !== 0
    )
      invalid()
    lexer.expect('obj')
    const value = parseValue(lexer)
    let streamBytes: number | undefined
    if (lexer.peek()?.value === 'stream') {
      lexer.expect('stream')
      if (!isDictionary(value) || lexer.buffered.length) invalid()
      const length = value.entries.get('Length')
      if (
        typeof length !== 'number' ||
        !Number.isSafeInteger(length) ||
        length < 0 ||
        length > MAX_RAW_BYTES
      )
        invalid()
      let streamStart = lexer.cursor
      if (text.slice(streamStart, streamStart + 2) === '\r\n') streamStart += 2
      else if (text[streamStart] === '\n') streamStart++
      else invalid()
      const streamEnd = streamStart + length
      if (streamEnd > bytes.length) invalid()
      streamBytes = checkStream(bytes.subarray(streamStart, streamEnd), value, budget)
      lexer.cursor = streamEnd
      lexer.expect('endstream')
    }
    lexer.expect('endobj')
    objects.set(id, start)
    values.set(id, streamBytes === undefined ? { value } : { value, streamBytes })
    if (objects.size > MAX_OBJECTS) invalid()
  }
  checkFontWidths(values)
  // Count repeated page-content references too: a small file must not expand
  // into gigabytes of extraction work by referencing the same stream repeatedly.
  let contentReferences = 0
  const contentBytes = (value: Value, path: Set<number>, depth = 0): number => {
    if (depth > 32 || ++contentReferences > MAX_OBJECTS) invalid()
    if (isReference(value)) {
      const object = values.get(value.id)
      if (!object || path.has(value.id)) invalid()
      if (object.streamBytes !== undefined) return object.streamBytes
      return contentBytes(object.value, new Set([...path, value.id]), depth + 1)
    }
    if (!Array.isArray(value)) invalid()
    let total = 0
    for (const child of value) {
      total += contentBytes(child, path, depth + 1)
      if (total > MAX_EXPANDED_BYTES) invalid()
    }
    return total
  }
  let totalContent = 0
  let pageCount = 0
  for (const object of values.values()) {
    if (!isDictionary(object.value) || !named(object.value.entries.get('Type'), 'Page')) continue
    if (++pageCount > MAX_PAGES) invalid()
    const contents = object.value.entries.get('Contents')
    if (contents !== undefined && contents !== null)
      totalContent += contentBytes(contents, new Set())
    if (totalContent > MAX_EXPANDED_BYTES) invalid()
  }
  const xrefStart = lexer.next().start
  const indexed = new Set<number>()
  let xrefEntries = 0
  while (lexer.peek()?.value !== 'trailer') {
    const first = lexer.integer()
    const count = lexer.integer()
    xrefEntries += count
    if (!count || first + count > MAX_OBJECTS + 1 || xrefEntries > MAX_OBJECTS + 1) invalid()
    for (let index = first; index < first + count; index++) {
      const offset = lexer.integer()
      const generation = lexer.integer()
      const state = lexer.next().value
      if (state === 'n') {
        if (generation !== 0 || indexed.has(index) || objects.get(index) !== offset) invalid()
        indexed.add(index)
      } else if (state !== 'f') invalid()
    }
  }
  if (indexed.size !== objects.size) invalid()
  lexer.expect('trailer')
  const trailer = parseValue(lexer)
  if (!isDictionary(trailer)) invalid()
  const root = trailer.entries.get('Root')
  const size = trailer.entries.get('Size')
  if (
    !isReference(root) ||
    !objects.has(root.id) ||
    typeof size !== 'number' ||
    size > MAX_OBJECTS + 1 ||
    size < objects.size + 1
  )
    invalid()
  lexer.expect('startxref')
  if (lexer.integer() !== xrefStart || lexer.peek() !== undefined) invalid()
}

class NoExternalData {
  fetch(): Promise<never> {
    return Promise.reject(new Error('External PDF resources are disabled'))
  }
}

type PdfModule = Awaited<ReturnType<typeof getResolvedPDFJS>>
type PdfDocument = Awaited<ReturnType<PdfModule['getDocument']>['promise']>

async function withPdf<T>(
  bytes: Uint8Array,
  operation: (pdf: PdfDocument) => Promise<T>,
): Promise<T> {
  let loading: ReturnType<PdfModule['getDocument']> | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    preflight(bytes)
    const pdfjs = await getResolvedPDFJS()
    const parameters = {
      data: bytes.slice(),
      verbosity: 0,
      stopAtErrors: true,
      disableFontFace: true,
      // Defense in depth for PDF.js versions that support runtime codegen.
      isEvalSupported: false,
      useSystemFonts: true,
      useWorkerFetch: false,
      BinaryDataFactory: NoExternalData,
      useWasm: false,
      enableXfa: false,
      maxImageSize: 0,
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      disableAutoFetch: true,
      disableRange: true,
      disableStream: true,
      pdfBug: false,
    }
    loading = pdfjs.getDocument(parameters)
    const work = async () => {
      const pdf = await loading?.promise
      if (!pdf || !Number.isInteger(pdf.numPages) || pdf.numPages < 1 || pdf.numPages > MAX_PAGES)
        invalid()
      return operation(pdf)
    }
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error(UNSUPPORTED)), TIMEOUT_MS)
    })
    return await Promise.race([work(), timeout])
  } catch {
    return invalid()
  } finally {
    if (timer !== undefined) clearTimeout(timer)
    // Destroy the loading task even if loading, parsing or extraction failed.
    if (loading) await loading.destroy().catch(() => undefined)
  }
}

export async function inspectPdf(bytes: Uint8Array): Promise<{ pageCount: number }> {
  return withPdf(bytes, async (pdf) => ({ pageCount: pdf.numPages }))
}

export interface ReadPdfOptions {
  pageStart?: number
  pageEnd?: number
  maxCharacters?: number
}
export interface PdfPageText {
  page: number
  text: string
  truncated: boolean
}
export interface PdfText {
  pageCount: number
  pageStart: number
  pageEnd: number
  pages: PdfPageText[]
  truncated: boolean
}

export async function readPdf(bytes: Uint8Array, options: ReadPdfOptions = {}): Promise<PdfText> {
  const pageStart = options.pageStart ?? 1
  const maxCharacters = options.maxCharacters ?? MAX_CHARACTERS
  if (
    !Number.isInteger(pageStart) ||
    pageStart < 1 ||
    !Number.isInteger(maxCharacters) ||
    maxCharacters < 1 ||
    maxCharacters > MAX_CHARACTERS
  )
    invalid()
  return withPdf(bytes, async (pdf) => {
    const pageEnd = options.pageEnd ?? Math.min(pdf.numPages, pageStart + MAX_READ_PAGES - 1)
    if (
      !Number.isInteger(pageEnd) ||
      pageEnd < pageStart ||
      pageEnd > pdf.numPages ||
      pageEnd - pageStart + 1 > MAX_READ_PAGES
    )
      invalid()
    const pages: PdfPageText[] = []
    let remaining = maxCharacters
    for (let number = pageStart; number <= pageEnd; number++) {
      if (remaining === 0) {
        pages.push({ page: number, text: '', truncated: true })
        continue
      }
      const page = await pdf.getPage(number)
      let text = ''
      let truncated = false
      // The stream avoids materializing every text item for the whole page.
      const reader = page.streamTextContent().getReader()
      try {
        while (true) {
          const chunk: { done: boolean; value?: unknown } = await reader.read()
          if (chunk.done) break
          // PDF.js can race close against cancel. Drain the bounded page stream
          // after truncation, discarding chunks, then destroy the document.
          if (truncated) continue
          const value = chunk.value
          if (
            typeof value !== 'object' ||
            value === null ||
            !('items' in value) ||
            !Array.isArray(value.items)
          )
            invalid()
          for (const item of value.items) {
            if (
              typeof item !== 'object' ||
              item === null ||
              !('str' in item) ||
              typeof item.str !== 'string'
            )
              continue
            const piece = item.str + ('hasEOL' in item && item.hasEOL ? '\n' : '')
            if (piece.length > remaining) {
              text += piece.slice(0, remaining)
              remaining = 0
              truncated = true
              break
            }
            text += piece
            remaining -= piece.length
          }
        }
      } finally {
        reader.releaseLock()
        page.cleanup()
      }
      pages.push({ page: number, text, truncated })
    }
    return {
      pageCount: pdf.numPages,
      pageStart,
      pageEnd,
      pages,
      truncated: pages.some((page) => page.truncated),
    }
  })
}
