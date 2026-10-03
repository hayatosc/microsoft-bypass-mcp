import { safeUnzip } from '../outlook/attachments/archive.js'
import { inspectDocx, inspectXlsx, readDocx, readXlsx } from '../outlook/attachments/office.js'
import { inspectPdf, readPdf } from '../outlook/attachments/pdf.js'
import { MAX_ATTACHMENT_BYTES, MAX_OUTPUT_CHARACTERS } from '../outlook/attachments/schema.js'
import type { ReadAttachmentInput } from '../outlook/attachments/schema.js'

export { MAX_ATTACHMENT_BYTES, MAX_OUTPUT_CHARACTERS }

export class DocumentSourceError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DocumentSourceError'
  }
}

export type DocumentFormat = 'pdf' | 'docx' | 'xlsx'
export type DocumentSelection = ReadAttachmentInput['selection']

export interface DecodedDocument<TSource extends { format: DocumentFormat }> {
  source: TSource
  bytes: Uint8Array
}

export function formatOf(name: string, contentType: string): DocumentFormat | null {
  const dot = name.lastIndexOf('.')
  const extension = dot === -1 ? '' : name.slice(dot + 1).toLowerCase()
  const media = contentType.split(';')[0]?.trim().toLowerCase() ?? ''
  const generic = media === '' || media === 'application/octet-stream'
  if (extension === 'pdf' && (generic || media === 'application/pdf')) return 'pdf'
  if (
    extension === 'docx' &&
    (generic || media === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
  )
    return 'docx'
  if (
    extension === 'xlsx' &&
    (generic || media === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  )
    return 'xlsx'
  return null
}

export function decodeBase64Document<TSource extends { format: DocumentFormat }>(
  source: TSource,
  contentBytes: string,
): DecodedDocument<TSource> {
  if (contentBytes.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(contentBytes))
    throw new DocumentSourceError('Invalid document encoding')
  const padding = contentBytes.endsWith('==') ? 2 : contentBytes.endsWith('=') ? 1 : 0
  const length = (contentBytes.length / 4) * 3 - padding
  if (length < 1 || length > MAX_ATTACHMENT_BYTES)
    throw new DocumentSourceError('Document exceeds the file size limit')
  let decoded: string
  try {
    decoded = atob(contentBytes)
  } catch {
    throw new DocumentSourceError('Invalid document encoding')
  }
  if (decoded.length !== length) throw new DocumentSourceError('Invalid document size')
  const bytes = new Uint8Array(decoded.length)
  for (let index = 0; index < decoded.length; index++) bytes[index] = decoded.charCodeAt(index)
  return { source, bytes }
}

export async function inspectDocument<TSource extends { format: DocumentFormat }>(
  document: DecodedDocument<TSource>,
) {
  if (document.source.format === 'pdf')
    return {
      source: document.source,
      untrustedContent: true as const,
      structure: { format: 'pdf' as const, ...(await inspectPdf(document.bytes)) },
    }
  const entries = safeUnzip(document.bytes)
  return {
    source: document.source,
    untrustedContent: true as const,
    structure: document.source.format === 'docx' ? inspectDocx(entries) : inspectXlsx(entries),
  }
}

export async function readDocument<TSource extends { format: DocumentFormat }>(
  document: DecodedDocument<TSource>,
  selection: DocumentSelection,
) {
  if (document.source.format !== selection.format)
    throw new DocumentSourceError('Requested format does not match this document')
  if (selection.format === 'pdf')
    return {
      source: document.source,
      untrustedContent: true as const,
      data: { format: 'pdf' as const, ...(await readPdf(document.bytes, selection)) },
    }
  const entries = safeUnzip(document.bytes)
  const data =
    selection.format === 'docx'
      ? readDocx(entries, {
          ...(selection.sectionId === undefined ? {} : { sectionId: selection.sectionId }),
          offset: selection.offset,
          length: selection.length,
        })
      : readXlsx(entries, selection)
  if (
    data.format === 'xlsx' &&
    data.cells.reduce((sum, cell) => {
      const valueLength = typeof cell.value === 'string' ? cell.value.length : 0
      // Preserve the existing budget for identical value/rawValue text.
      const rawLength = cell.rawValue === cell.value ? 0 : (cell.rawValue?.length ?? 0)
      return sum + valueLength + rawLength
    }, 0) > MAX_OUTPUT_CHARACTERS
  )
    throw new DocumentSourceError('Selected cells exceed the text limit; request a smaller range')
  return { source: document.source, untrustedContent: true as const, data }
}

export function boundedToolResult<T extends Record<string, unknown>>(result: T) {
  const text = JSON.stringify(result)
  if (new TextEncoder().encode(text).byteLength > 128 * 1024)
    throw new DocumentSourceError(
      'Result exceeds the output limit; request a smaller selection or list page',
    )
  return { content: [{ type: 'text' as const, text }], structuredContent: result }
}
