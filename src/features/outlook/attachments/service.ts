/** Request-local attachment retrieval, validation and parser dispatch. No storage/cache. */
import { z } from 'zod'

import type { PowerAutomateClient } from '../../../lib/power-automate.js'
import { safeUnzip } from './archive.js'
import { inspectDocx, inspectXlsx, readDocx, readXlsx } from './office.js'
import { inspectPdf, readPdf } from './pdf.js'
import { attachmentIdSchema, MAX_ATTACHMENT_BYTES, MAX_OUTPUT_CHARACTERS } from './schema.js'
import type { ReadAttachmentInput } from './schema.js'

export class AttachmentError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'AttachmentError'
  }
}
const graphMetadataSchema = z.object({
  '@odata.type': z.string().max(128),
  id: attachmentIdSchema,
  name: z.string().min(1).max(512),
  contentType: z.string().max(256),
  size: z.number().int().nonnegative(),
  isInline: z.boolean(),
})
type GraphMetadata = z.infer<typeof graphMetadataSchema>
const graphListSchema = z.object({
  value: z.array(graphMetadataSchema).max(50),
  '@odata.nextLink': z.string().max(8192).optional(),
})
const graphFileSchema = graphMetadataSchema.extend({
  '@odata.type': z.literal('#microsoft.graph.fileAttachment'),
  size: z.number().int().min(1).max(MAX_ATTACHMENT_BYTES),
  contentBytes: z
    .string()
    .min(4)
    .max(4 * Math.ceil(MAX_ATTACHMENT_BYTES / 3)),
})
function formatOf(metadata: GraphMetadata): 'pdf' | 'docx' | 'xlsx' | null {
  const extension = metadata.name.split('.').at(-1)?.toLowerCase()
  const contentType = metadata.contentType.split(';')[0]?.trim().toLowerCase()
  const generic = contentType === 'application/octet-stream' || contentType === ''
  if (extension === 'pdf' && (generic || contentType === 'application/pdf')) return 'pdf'
  if (
    extension === 'docx' &&
    (generic ||
      contentType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
  )
    return 'docx'
  if (
    extension === 'xlsx' &&
    (generic || contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  )
    return 'xlsx'
  return null
}
function typeOf(type: string): 'file' | 'item' | 'reference' | 'unknown' {
  switch (type) {
    case '#microsoft.graph.fileAttachment':
      return 'file'
    case '#microsoft.graph.itemAttachment':
      return 'item'
    case '#microsoft.graph.referenceAttachment':
      return 'reference'
    default:
      return 'unknown'
  }
}
export function normalizeAttachments(body: unknown, messageId: string, top: number, skip: number) {
  const parsed = graphListSchema.safeParse(body)
  if (!parsed.success || parsed.data.value.length > top)
    throw new AttachmentError('Invalid attachment list response')
  const attachments = parsed.data.value.map((metadata) => {
    const attachmentType = typeOf(metadata['@odata.type'])
    const supportedFormat = formatOf(metadata)
    return {
      attachmentId: metadata.id,
      name: metadata.name,
      contentType: metadata.contentType,
      size: metadata.size,
      isInline: metadata.isInline,
      attachmentType,
      supportedFormat,
      readable:
        attachmentType === 'file' &&
        supportedFormat !== null &&
        metadata.size > 0 &&
        metadata.size <= MAX_ATTACHMENT_BYTES,
    }
  })
  const hasMore = Boolean(parsed.data['@odata.nextLink'])
  if (hasMore && attachments.length === 0)
    throw new AttachmentError('Invalid attachment pagination response')
  // Never follow or return an upstream URL. A new fixed-path call supplies an offset.
  const next = skip + attachments.length
  return { messageId, attachments, hasMore, nextOffset: hasMore && next <= 10000 ? next : null }
}
async function retrieve(
  client: PowerAutomateClient,
  target: { messageId: string; attachmentId: string },
) {
  const raw = await client.call('get_attachment', target)
  const parsed = graphFileSchema.safeParse(raw)
  if (!parsed.success)
    throw new AttachmentError('Attachment must be a valid file attachment of at most 4 MiB')
  const data = parsed.data
  if (data.id !== target.attachmentId)
    throw new AttachmentError('Attachment response ID does not match')
  const format = formatOf(data)
  if (!format)
    throw new AttachmentError(
      'Only PDF, DOCX and XLSX file attachments with matching media types are supported',
    )
  const base64 = data.contentBytes
  // Padded standard base64 only. Estimate length before decoding, then verify actual length.
  if (base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)) {
    throw new AttachmentError('Invalid attachment encoding')
  }
  const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0
  const length = (base64.length / 4) * 3 - padding
  if (length > MAX_ATTACHMENT_BYTES || length < 1)
    throw new AttachmentError('Attachment exceeds the file size limit')
  let decoded: string
  try {
    decoded = atob(base64)
  } catch {
    throw new AttachmentError('Invalid attachment encoding')
  }
  if (decoded.length !== length) throw new AttachmentError('Invalid attachment size')
  // Graph size includes attachment metadata, so it need not equal contentBytes length.
  const bytes = new Uint8Array(decoded.length)
  for (let index = 0; index < decoded.length; index++) bytes[index] = decoded.charCodeAt(index)
  const source = {
    messageId: target.messageId,
    attachmentId: data.id,
    name: data.name,
    contentType: data.contentType,
    size: data.size,
    format,
  }
  return { source, bytes }
}
export async function inspectAttachment(
  client: PowerAutomateClient,
  target: { messageId: string; attachmentId: string },
) {
  const { source, bytes } = await retrieve(client, target)
  if (source.format === 'pdf')
    return {
      source,
      untrustedContent: true as const,
      structure: { format: 'pdf' as const, ...(await inspectPdf(bytes)) },
    }
  const entries = safeUnzip(bytes)
  const structure = source.format === 'docx' ? inspectDocx(entries) : inspectXlsx(entries)
  return { source, untrustedContent: true as const, structure }
}
export async function readAttachment(client: PowerAutomateClient, input: ReadAttachmentInput) {
  const { source, bytes } = await retrieve(client, {
    messageId: input.messageId,
    attachmentId: input.attachmentId,
  })
  const selection = input.selection
  if (source.format !== selection.format)
    throw new AttachmentError('Requested format does not match this attachment')
  if (selection.format === 'pdf')
    return {
      source,
      untrustedContent: true as const,
      data: { format: 'pdf' as const, ...(await readPdf(bytes, selection)) },
    }
  const entries = safeUnzip(bytes)
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
    data.cells.reduce(
      (sum, cell) => sum + (typeof cell.value === 'string' ? cell.value.length : 0),
      0,
    ) > MAX_OUTPUT_CHARACTERS
  ) {
    throw new AttachmentError('Selected cells exceed the text limit; request a smaller range')
  }
  return { source, untrustedContent: true as const, data }
}
/** Also bound JSON scaffolding, repeated raw values and escaped text. Never truncate invalid JSON. */
export function attachmentResult<T extends Record<string, unknown>>(result: T) {
  const text = JSON.stringify(result)
  if (new TextEncoder().encode(text).byteLength > 128 * 1024)
    throw new AttachmentError(
      'Result exceeds the output limit; request a smaller selection or list page',
    )
  return { content: [{ type: 'text' as const, text }], structuredContent: result }
}
