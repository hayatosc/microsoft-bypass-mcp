/** Request-local, namespace-aware OOXML readers. No formulas, relationships, or scripts execute. */
import { SaxesParser } from 'saxes'
import type { SaxesTagNS } from 'saxes'

export type OfficeEntries = ReadonlyMap<string, Uint8Array>
export type OfficeErrorCode = 'INVALID_OFFICE' | 'OFFICE_LIMIT' | 'INVALID_RANGE'

/** Messages deliberately exclude archive paths, XML, and attachment contents. */
export class OfficeParseError extends Error {
  constructor(readonly code: OfficeErrorCode) {
    super(
      code === 'OFFICE_LIMIT'
        ? 'The Office document exceeds the safe parsing limits.'
        : code === 'INVALID_RANGE'
          ? 'The requested document section or cell range is invalid.'
          : 'The Office document is malformed or uses unsupported structures.',
    )
    this.name = 'OfficeParseError'
  }
}

const WORD = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
])
const SHEET = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
])
const REL = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
])
const PACKAGE_REL = new Set(['http://schemas.openxmlformats.org/package/2006/relationships'])
const MAX_XML_BYTES = 8 * 1024 * 1024
const MAX_ELEMENTS = 500_000
const MAX_PARAGRAPHS = 10_000
const MAX_SECTIONS = 200
const MAX_CELLS = 50_000
const MAX_SHEETS = 50
const MAX_READ_CHARS = 20_000
const MAX_ROW = 1_048_576
const MAX_COLUMN = 16_384

interface ParseContext {
  entries: OfficeEntries
  elements: number
  parsedBytes: number
  cells: number
}
interface XmlHandlers {
  open?: (tag: SaxesTagNS, parents: readonly SaxesTagNS[]) => void
  close?: (tag: SaxesTagNS, parents: readonly SaxesTagNS[]) => void
  text?: (text: string, parents: readonly SaxesTagNS[]) => void
}

function invalid(): never {
  throw new OfficeParseError('INVALID_OFFICE')
}
function limited(): never {
  throw new OfficeParseError('OFFICE_LIMIT')
}
function badRange(): never {
  throw new OfficeParseError('INVALID_RANGE')
}
function isTag(tag: SaxesTagNS | undefined, local: string, namespaces: ReadonlySet<string>) {
  return tag !== undefined && tag.local === local && namespaces.has(tag.uri)
}
function attribute(tag: SaxesTagNS, local: string, namespaces?: ReadonlySet<string>) {
  return Object.values(tag.attributes).find(
    (value) => value.local === local && (namespaces ? namespaces.has(value.uri) : value.uri === ''),
  )?.value
}
function context(entries: OfficeEntries): ParseContext {
  let xmlBytes = 0
  for (const [path, bytes] of entries) {
    if (/\.(?:xml|rels)$/i.test(path)) xmlBytes += bytes.byteLength
    if (xmlBytes > MAX_XML_BYTES) limited()
  }
  return { entries, elements: 0, parsedBytes: 0, cells: 0 }
}
function decodeXml(bytes: Uint8Array) {
  let encoding = 'utf-8'
  if (
    (bytes[0] === 0xff && bytes[1] === 0xfe) ||
    (bytes[0] === 0x3c && bytes[1] === 0 && bytes[2] === 0x3f && bytes[3] === 0)
  )
    encoding = 'utf-16le'
  if (
    (bytes[0] === 0xfe && bytes[1] === 0xff) ||
    (bytes[0] === 0 && bytes[1] === 0x3c && bytes[2] === 0 && bytes[3] === 0x3f)
  )
    encoding = 'utf-16be'
  try {
    return new TextDecoder(encoding, { fatal: true, ignoreBOM: false }).decode(bytes)
  } catch {
    return invalid()
  }
}
function xml(
  ctx: ParseContext,
  path: string,
  root: string,
  namespaces: ReadonlySet<string>,
  handlers: XmlHandlers,
) {
  const bytes = ctx.entries.get(path)
  if (!bytes) invalid()
  ctx.parsedBytes += bytes.byteLength
  if (ctx.parsedBytes > MAX_XML_BYTES) limited()
  const source = decodeXml(bytes)
  // Reject declarations before parsing, including entity-expansion payloads. The SAX parser
  // also rejects undeclared entities and never loads external resources.
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source)) invalid()
  const parents: SaxesTagNS[] = []
  const parser = new SaxesParser({ xmlns: true })
  let attributeCount = 0
  parser.on('opentagstart', () => {
    attributeCount = 0
  })
  parser.on('attribute', () => {
    if (++attributeCount > 100) limited()
  })
  parser.on('error', () => invalid())
  parser.on('doctype', () => invalid())
  parser.on('xmldecl', (decl) => {
    if (decl.encoding && !/^utf-?(?:8|16(?:le|be)?)$/i.test(decl.encoding)) invalid()
  })
  parser.on('opentag', (tag) => {
    if (++ctx.elements > MAX_ELEMENTS || parents.length >= 64) limited()
    if (Object.keys(tag.attributes).length > 100) limited()
    if (parents.length === 0 && !isTag(tag, root, namespaces)) invalid()
    handlers.open?.(tag, parents)
    parents.push(tag)
  })
  parser.on('closetag', (tag) => {
    parents.pop()
    handlers.close?.(tag, parents)
  })
  const onText = (text: string) => handlers.text?.(text, parents)
  parser.on('text', onText)
  parser.on('cdata', onText)
  try {
    parser.write(source).close()
  } catch (error) {
    if (error instanceof OfficeParseError) throw error
    invalid()
  }
}

export interface DocxSection {
  id: string
  title: string
  headingLevel: number | null
  /** One-based inclusive paragraph positions. */
  paragraphStart: number
  paragraphEnd: number
  /** Zero-based, end-exclusive UTF-16 positions in the extracted body text. */
  start: number
  end: number
}
export interface DocxInspection {
  format: 'docx'
  paragraphCount: number
  textCharacters: number
  sections: DocxSection[]
  warnings: string[]
}
export interface DocxRead {
  format: 'docx'
  sectionId: string | null
  offset: number
  length: number
  sourceStart: number
  sourceEnd: number
  totalCharacters: number
  nextOffset: number | null
  text: string
  warnings: string[]
}
interface Paragraph {
  text: string
  headingLevel: number | null
}
interface WordStyle {
  basedOn: string | undefined
  heading: number | null
  outline: number | undefined
}
const DOCX_WARNINGS = [
  'Main-document body text only; headers, footers, notes, drawings, and text boxes are omitted. Table paragraphs are flattened in document order.',
  'Sections follow heading paragraphs. Paragraphs are separated by one newline; offsets count UTF-16 code units, not file bytes or rendered pages.',
]
function heading(value: string | undefined): number | null {
  const match = value?.match(/^heading\s*([1-9])$/i)
  return match?.[1] ? Number(match[1]) : null
}
function outline(value: string | undefined) {
  if (value === undefined || !/^[0-9]$/.test(value)) invalid()
  return Number(value)
}
function wordStyles(ctx: ParseContext) {
  const styles = new Map<string, WordStyle>()
  if (!ctx.entries.has('word/styles.xml')) return styles
  let current: WordStyle | undefined
  let id: string | undefined
  xml(ctx, 'word/styles.xml', 'styles', WORD, {
    open(tag, parents) {
      if (isTag(tag, 'style', WORD) && parents.length === 1) {
        if (styles.size >= 4096) limited()
        id = attribute(tag, 'styleId', WORD)
        current =
          attribute(tag, 'type', WORD) === 'paragraph'
            ? { basedOn: undefined, heading: heading(id), outline: undefined }
            : undefined
      } else if (current && parents.length === 2) {
        if (isTag(tag, 'name', WORD))
          current.heading = heading(attribute(tag, 'val', WORD)) ?? current.heading
        if (isTag(tag, 'basedOn', WORD)) current.basedOn = attribute(tag, 'val', WORD)
      } else if (current && isTag(tag, 'outlineLvl', WORD) && isTag(parents.at(-1), 'pPr', WORD)) {
        current.outline = outline(attribute(tag, 'val', WORD))
      }
    },
    close(tag, parents) {
      if (isTag(tag, 'style', WORD) && parents.length === 1) {
        if (current && id) {
          if (styles.has(id)) invalid()
          styles.set(id, current)
        }
        current = undefined
        id = undefined
      }
    },
  })
  return styles
}
function styleHeading(
  id: string | undefined,
  styles: ReadonlyMap<string, WordStyle>,
): number | null {
  const visited = new Set<string>()
  while (id !== undefined) {
    if (visited.has(id) || visited.size >= 32) invalid()
    visited.add(id)
    const style = styles.get(id)
    if (!style) return heading(id)
    if (style.outline !== undefined) return style.outline < 9 ? style.outline + 1 : null
    if (style.heading !== null) return style.heading
    id = style.basedOn
  }
  return null
}
function parseDocx(entries: OfficeEntries) {
  const ctx = context(entries)
  const styles = wordStyles(ctx)
  const paragraphs: Paragraph[] = []
  let bodySeen = false
  let bodyDepth = 0
  let paragraphDepth = 0
  let ignoredDepth = 0
  let chunks: string[] = []
  let paragraphStyle: string | undefined
  let paragraphOutline: number | undefined
  xml(ctx, 'word/document.xml', 'document', WORD, {
    open(tag, parents) {
      const depth = parents.length + 1
      if (isTag(tag, 'body', WORD) && parents.length === 1) {
        if (bodySeen) invalid()
        bodySeen = true
        bodyDepth = depth
      }
      if (!bodyDepth || ignoredDepth) return
      if (WORD.has(tag.uri) && ['del', 'drawing', 'pict', 'txbxContent'].includes(tag.local)) {
        ignoredDepth = depth
        return
      }
      if (isTag(tag, 'p', WORD)) {
        if (paragraphDepth) invalid()
        if (paragraphs.length >= MAX_PARAGRAPHS) limited()
        paragraphDepth = depth
        chunks = []
        paragraphStyle = undefined
        paragraphOutline = undefined
      } else if (paragraphDepth) {
        if (isTag(tag, 'pStyle', WORD) && isTag(parents.at(-1), 'pPr', WORD)) {
          paragraphStyle = attribute(tag, 'val', WORD)
        } else if (isTag(tag, 'outlineLvl', WORD) && isTag(parents.at(-1), 'pPr', WORD)) {
          paragraphOutline = outline(attribute(tag, 'val', WORD))
        } else if (isTag(tag, 'tab', WORD) && isTag(parents.at(-1), 'r', WORD)) chunks.push('\t')
        else if (
          (isTag(tag, 'br', WORD) || isTag(tag, 'cr', WORD)) &&
          isTag(parents.at(-1), 'r', WORD)
        )
          chunks.push('\n')
        else if (isTag(tag, 'noBreakHyphen', WORD) && isTag(parents.at(-1), 'r', WORD))
          chunks.push('\u2011')
        else if (isTag(tag, 'softHyphen', WORD) && isTag(parents.at(-1), 'r', WORD))
          chunks.push('\u00ad')
      }
    },
    text(text, parents) {
      if (paragraphDepth && !ignoredDepth && isTag(parents.at(-1), 't', WORD)) chunks.push(text)
    },
    close(tag, parents) {
      const depth = parents.length + 1
      if (ignoredDepth) {
        if (depth === ignoredDepth) ignoredDepth = 0
        return
      }
      if (paragraphDepth === depth && isTag(tag, 'p', WORD)) {
        paragraphs.push({
          text: chunks.join(''),
          headingLevel:
            paragraphOutline === undefined
              ? styleHeading(paragraphStyle, styles)
              : paragraphOutline < 9
                ? paragraphOutline + 1
                : null,
        })
        paragraphDepth = 0
      }
      if (bodyDepth === depth && isTag(tag, 'body', WORD)) bodyDepth = 0
    },
  })
  if (!bodySeen) invalid()
  const sections: DocxSection[] = []
  let position = 0
  for (const [index, paragraph] of paragraphs.entries()) {
    if (index === 0 || paragraph.headingLevel !== null) {
      if (sections.length >= MAX_SECTIONS) limited()
      const previous = sections.at(-1)
      if (previous) {
        previous.end = position
        previous.paragraphEnd = index
      }
      sections.push({
        id: `section-${sections.length + 1}`,
        title: paragraph.headingLevel === null ? 'Document' : paragraph.text.trim().slice(0, 200),
        headingLevel: paragraph.headingLevel,
        paragraphStart: index + 1,
        paragraphEnd: paragraphs.length,
        start: position,
        end: 0,
      })
    }
    position += paragraph.text.length + (index < paragraphs.length - 1 ? 1 : 0)
  }
  const last = sections.at(-1)
  if (last) last.end = position
  return { paragraphs, sections, text: paragraphs.map((paragraph) => paragraph.text).join('\n') }
}

export function inspectDocx(entries: OfficeEntries): DocxInspection {
  const parsed = parseDocx(entries)
  return {
    format: 'docx',
    paragraphCount: parsed.paragraphs.length,
    textCharacters: parsed.text.length,
    sections: parsed.sections,
    warnings: [...DOCX_WARNINGS],
  }
}

/** Offset is relative to the selected section, or to the whole document if omitted. */
export function readDocx(
  entries: OfficeEntries,
  selection: { sectionId?: string; offset?: number; length?: number },
): DocxRead {
  const offset = selection.offset ?? 0
  const length = selection.length ?? 10_000
  if (
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    !Number.isSafeInteger(length) ||
    length < 1 ||
    length > MAX_READ_CHARS
  )
    badRange()
  const parsed = parseDocx(entries)
  const section =
    selection.sectionId === undefined
      ? undefined
      : parsed.sections.find((item) => item.id === selection.sectionId)
  if (selection.sectionId !== undefined && !section) badRange()
  const start = section?.start ?? 0
  const end = section?.end ?? parsed.text.length
  const totalCharacters = end - start
  if (offset > totalCharacters) badRange()
  const sourceStart = start + offset
  const sourceEnd = Math.min(sourceStart + length, end)
  const text = parsed.text.slice(sourceStart, sourceEnd)
  return {
    format: 'docx',
    sectionId: selection.sectionId ?? null,
    offset,
    length: text.length,
    sourceStart,
    sourceEnd,
    totalCharacters,
    nextOffset: sourceEnd < end ? offset + text.length : null,
    text,
    warnings: [...DOCX_WARNINGS],
  }
}

export interface XlsxDimensions {
  range: string | null
  firstRow: number | null
  lastRow: number | null
  firstColumn: number | null
  lastColumn: number | null
  cellCount: number
}
export interface XlsxInspection {
  format: 'xlsx'
  sheets: { name: string; dimensions: XlsxDimensions }[]
  dateSystem: '1900' | '1904'
  warnings: string[]
}
export interface XlsxCell {
  address: string
  row: number
  column: number
  type: 'number' | 'string' | 'boolean' | 'error' | 'date' | 'blank'
  value: string | boolean | null
  rawValue: string | null
  present: boolean
  hasFormula: boolean
  cachedValueMissing: boolean
}
export interface XlsxRead {
  format: 'xlsx'
  sheet: string
  range: string
  dimensions: XlsxDimensions
  dateSystem: '1900' | '1904'
  cells: XlsxCell[]
  warnings: string[]
}
const XLSX_WARNINGS = [
  'Values are stored or cached values; formulas are never evaluated and their caches may be stale. Missing formula caches are marked explicitly.',
  'Numbers remain raw strings to preserve precision. Numeric dates remain serial numbers; styles and number formats are not applied. ISO date cells remain raw strings.',
  'Dimensions are computed from stored cells. Merged cells, drawings, external links, and embedded objects are not expanded or followed.',
]
interface Relationship {
  type: string
  target: string
  external: boolean
}
interface SheetPart {
  name: string
  path: string
}
function relatedPath(base: string, target: string) {
  // OPC package paths only. Reject URLs, encoded separators, traversal, and fragments.
  if (
    !target ||
    /[\\:%?#]/.test(target) ||
    target.split('').some((character) => character.charCodeAt(0) < 32) ||
    target.startsWith('//')
  )
    invalid()
  const parts = target.split('/')
  if (parts.some((part) => part === '.' || part === '..') || parts.slice(1).some((part) => !part))
    invalid()
  return target.startsWith('/') ? target.slice(1) : base + target
}
function workbook(ctx: ParseContext) {
  const relationships = new Map<string, Relationship>()
  xml(ctx, 'xl/_rels/workbook.xml.rels', 'Relationships', PACKAGE_REL, {
    open(tag, parents) {
      if (!isTag(tag, 'Relationship', PACKAGE_REL) || parents.length !== 1) return
      const id = attribute(tag, 'Id')
      const type = attribute(tag, 'Type')
      const target = attribute(tag, 'Target')
      const mode = attribute(tag, 'TargetMode')
      if (
        !id ||
        !type ||
        !target ||
        relationships.has(id) ||
        (mode && mode !== 'External' && mode !== 'Internal')
      )
        invalid()
      if (relationships.size >= 4096) limited()
      relationships.set(id, { type, target, external: mode === 'External' })
    },
  })
  const sheets: SheetPart[] = []
  const names = new Set<string>()
  const paths = new Set<string>()
  let dateSystem: '1900' | '1904' = '1900'
  xml(ctx, 'xl/workbook.xml', 'workbook', SHEET, {
    open(tag, parents) {
      if (isTag(tag, 'workbookPr', SHEET) && parents.length === 1) {
        const date1904 = attribute(tag, 'date1904')
        if (date1904 !== undefined && !['0', '1', 'true', 'false'].includes(date1904)) invalid()
        if (date1904 === '1' || date1904 === 'true') dateSystem = '1904'
      }
      if (!isTag(tag, 'sheet', SHEET) || !isTag(parents.at(-1), 'sheets', SHEET)) return
      if (sheets.length >= MAX_SHEETS) limited()
      const name = attribute(tag, 'name')
      const id = attribute(tag, 'id', REL)
      const relation = id === undefined ? undefined : relationships.get(id)
      if (
        !name ||
        name.length > 128 ||
        names.has(name.toLowerCase()) ||
        !relation ||
        relation.external ||
        ![...REL].some((uri) => relation.type === `${uri}/worksheet`)
      )
        invalid()
      const path = relatedPath('xl/', relation.target)
      if (!ctx.entries.has(path) || paths.has(path)) invalid()
      names.add(name.toLowerCase())
      paths.add(path)
      sheets.push({ name, path })
    },
  })
  if (sheets.length === 0) invalid()
  const strings = [...relationships.values()].filter((relation) =>
    [...REL].some((uri) => relation.type === `${uri}/sharedStrings`),
  )
  if (strings.length > 1) invalid()
  const stringsRelation = strings[0]
  if (stringsRelation?.external) invalid()
  const sharedStringPath = stringsRelation ? relatedPath('xl/', stringsRelation.target) : undefined
  return { sheets, dateSystem, sharedStringPath }
}
function excelText(text: string) {
  // One pass preserves escaped literal sequences such as _x005F_x0041_.
  return text.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  )
}
function sharedStrings(ctx: ParseContext, path: string | undefined) {
  const strings: string[] = []
  if (path === undefined) return strings
  let chunks: string[] | undefined
  xml(ctx, path, 'sst', SHEET, {
    open(tag, parents) {
      if (isTag(tag, 'si', SHEET) && parents.length === 1) {
        if (strings.length >= MAX_CELLS) limited()
        chunks = []
      }
    },
    text(text, parents) {
      if (
        chunks &&
        isTag(parents.at(-1), 't', SHEET) &&
        (isTag(parents.at(-2), 'si', SHEET) || isTag(parents.at(-2), 'r', SHEET)) &&
        !parents.some((tag) => isTag(tag, 'rPh', SHEET))
      )
        chunks.push(text)
    },
    close(tag, parents) {
      if (isTag(tag, 'si', SHEET) && parents.length === 1 && chunks) {
        strings.push(excelText(chunks.join('')))
        chunks = undefined
      }
    },
  })
  return strings
}
function columnName(column: number): string {
  let result = ''
  while (column > 0) {
    column--
    result = String.fromCharCode(65 + (column % 26)) + result
    column = Math.floor(column / 26)
  }
  return result
}
function coordinate(value: string) {
  const match = value.match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/)
  if (!match?.[1] || !match[2]) badRange()
  let column = 0
  for (const character of match[1].toUpperCase())
    column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2])
  if (column > MAX_COLUMN || row > MAX_ROW) badRange()
  return { column, row, address: `${columnName(column)}${row}` }
}
function cellCoordinate(value: string) {
  try {
    return coordinate(value)
  } catch {
    return invalid()
  }
}
function cellRange(value: string) {
  const parts = value.split(':')
  if (parts.length > 2 || !parts[0]) badRange()
  const first = coordinate(parts[0])
  const last = coordinate(parts[1] ?? parts[0])
  if (
    first.row > last.row ||
    first.column > last.column ||
    (last.row - first.row + 1) * (last.column - first.column + 1) > 500
  )
    badRange()
  return {
    first,
    last,
    range: first.address === last.address ? first.address : `${first.address}:${last.address}`,
  }
}
function emptyDimensions(): XlsxDimensions {
  return {
    range: null,
    firstRow: null,
    lastRow: null,
    firstColumn: null,
    lastColumn: null,
    cellCount: 0,
  }
}
interface StoredCell {
  address: string
  row: number
  column: number
  type: string
  value: string[]
  inline: string[]
  sawValue: boolean
  hasFormula: boolean
}
function readStoredCell(cell: StoredCell, strings: readonly string[]): XlsxCell {
  const raw = cell.value.join('')
  const cachedValueMissing =
    cell.hasFormula && (!cell.sawValue || (raw.trim() === '' && cell.type !== 'str'))
  const result: XlsxCell = {
    address: cell.address,
    row: cell.row,
    column: cell.column,
    type: 'blank',
    value: null,
    rawValue: cell.sawValue ? raw : null,
    present: true,
    hasFormula: cell.hasFormula,
    cachedValueMissing,
  }
  if (cachedValueMissing) return result
  if (cell.type === 'inlineStr') {
    if (cell.hasFormula) invalid()
    result.type = 'string'
    result.value = excelText(cell.inline.join(''))
  } else if (cell.type === 's') {
    if (!cell.sawValue || !/^(?:0|[1-9][0-9]*)$/.test(raw.trim())) invalid()
    const index = Number(raw.trim())
    if (!Number.isSafeInteger(index) || strings[index] === undefined) invalid()
    result.type = 'string'
    result.value = strings[index]
  } else if (cell.type === 'str') {
    result.type = 'string'
    result.value = cell.sawValue ? excelText(raw) : null
  } else if (cell.type === 'b' && cell.sawValue) {
    if (!['0', '1', 'true', 'false'].includes(raw.trim())) invalid()
    result.type = 'boolean'
    result.value = raw.trim() === '1' || raw.trim() === 'true'
  } else if (cell.type === 'e' || cell.type === 'd') {
    result.type = cell.type === 'e' ? 'error' : 'date'
    result.value = cell.sawValue ? raw : null
  } else if (cell.type === 'n' && cell.sawValue && raw.trim() !== '') {
    if (!/^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?$/.test(raw.trim())) invalid()
    result.type = 'number'
    result.value = raw
  }
  return result
}
function worksheet(ctx: ParseContext, path: string, consume?: (cell: StoredCell) => void) {
  const dimensions = emptyDimensions()
  let previousRow = 0
  let row: number | undefined
  let previousColumn = 0
  let cell: StoredCell | undefined
  let sheetDataSeen = false
  xml(ctx, path, 'worksheet', SHEET, {
    open(tag, parents) {
      if (isTag(tag, 'sheetData', SHEET) && parents.length === 1) {
        if (sheetDataSeen) invalid()
        sheetDataSeen = true
      }
      if (isTag(tag, 'row', SHEET) && isTag(parents.at(-1), 'sheetData', SHEET)) {
        const declared = attribute(tag, 'r')
        if (declared !== undefined && !/^[1-9][0-9]{0,6}$/.test(declared)) invalid()
        row = declared === undefined ? previousRow + 1 : Number(declared)
        if (row <= previousRow || row > MAX_ROW) invalid()
        previousRow = row
        previousColumn = 0
      } else if (isTag(tag, 'c', SHEET) && isTag(parents.at(-1), 'row', SHEET)) {
        if (row === undefined || cell) invalid()
        if (++ctx.cells > MAX_CELLS) limited()
        const ref = attribute(tag, 'r')
        const point =
          ref === undefined
            ? cellCoordinate(`${columnName(previousColumn + 1)}${row}`)
            : cellCoordinate(ref)
        if (point.row !== row || point.column <= previousColumn) invalid()
        previousColumn = point.column
        const type = attribute(tag, 't') ?? 'n'
        if (!['n', 's', 'str', 'inlineStr', 'b', 'e', 'd'].includes(type)) invalid()
        cell = { ...point, type, value: [], inline: [], sawValue: false, hasFormula: false }
        dimensions.cellCount++
        dimensions.firstRow = Math.min(dimensions.firstRow ?? point.row, point.row)
        dimensions.lastRow = Math.max(dimensions.lastRow ?? point.row, point.row)
        dimensions.firstColumn = Math.min(dimensions.firstColumn ?? point.column, point.column)
        dimensions.lastColumn = Math.max(dimensions.lastColumn ?? point.column, point.column)
      } else if (cell && isTag(parents.at(-1), 'c', SHEET)) {
        if (isTag(tag, 'f', SHEET)) {
          if (cell.hasFormula) invalid()
          cell.hasFormula = true
        }
        if (isTag(tag, 'v', SHEET)) {
          if (cell.sawValue) invalid()
          cell.sawValue = true
        }
      }
    },
    text(text, parents) {
      if (!cell) return
      if (isTag(parents.at(-1), 'v', SHEET) && isTag(parents.at(-2), 'c', SHEET))
        cell.value.push(text)
      if (
        isTag(parents.at(-1), 't', SHEET) &&
        parents.some((tag) => isTag(tag, 'is', SHEET)) &&
        !parents.some((tag) => isTag(tag, 'rPh', SHEET))
      )
        cell.inline.push(text)
    },
    close(tag, parents) {
      if (isTag(tag, 'c', SHEET) && isTag(parents.at(-1), 'row', SHEET) && cell) {
        consume?.(cell)
        cell = undefined
      }
      if (isTag(tag, 'row', SHEET) && isTag(parents.at(-1), 'sheetData', SHEET)) row = undefined
    },
  })
  if (!sheetDataSeen) invalid()
  if (
    dimensions.firstRow !== null &&
    dimensions.firstColumn !== null &&
    dimensions.lastRow !== null &&
    dimensions.lastColumn !== null
  ) {
    const first = `${columnName(dimensions.firstColumn)}${dimensions.firstRow}`
    const last = `${columnName(dimensions.lastColumn)}${dimensions.lastRow}`
    dimensions.range = first === last ? first : `${first}:${last}`
  }
  return dimensions
}

export function inspectXlsx(entries: OfficeEntries): XlsxInspection {
  const ctx = context(entries)
  const book = workbook(ctx)
  return {
    format: 'xlsx',
    sheets: book.sheets.map((sheet) => ({
      name: sheet.name,
      dimensions: worksheet(ctx, sheet.path),
    })),
    dateSystem: book.dateSystem,
    warnings: [...XLSX_WARNINGS],
  }
}

export function readXlsx(
  entries: OfficeEntries,
  selection: { sheet: string; range: string },
): XlsxRead {
  const range = cellRange(selection.range)
  const ctx = context(entries)
  const book = workbook(ctx)
  const sheet = book.sheets.find((item) => item.name === selection.sheet)
  if (!sheet) badRange()
  const strings = sharedStrings(ctx, book.sharedStringPath)
  const selected = new Map<string, XlsxCell>()
  const dimensions = worksheet(ctx, sheet.path, (cell) => {
    if (
      cell.row >= range.first.row &&
      cell.row <= range.last.row &&
      cell.column >= range.first.column &&
      cell.column <= range.last.column
    ) {
      selected.set(cell.address, readStoredCell(cell, strings))
    }
  })
  const cells: XlsxCell[] = []
  for (let row = range.first.row; row <= range.last.row; row++) {
    for (let column = range.first.column; column <= range.last.column; column++) {
      const address = `${columnName(column)}${row}`
      cells.push(
        selected.get(address) ?? {
          address,
          row,
          column,
          type: 'blank',
          value: null,
          rawValue: null,
          present: false,
          hasFormula: false,
          cachedValueMissing: false,
        },
      )
    }
  }
  return {
    format: 'xlsx',
    sheet: sheet.name,
    range: range.range,
    dimensions,
    dateSystem: book.dateSystem,
    cells,
    warnings: [...XLSX_WARNINGS],
  }
}
