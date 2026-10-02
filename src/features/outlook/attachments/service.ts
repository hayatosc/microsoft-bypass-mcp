import { z } from 'zod'

import type { PowerAutomateClient } from '../../../lib/power-automate.js'
import {
  boundedToolResult,
  decodeBase64Document,
  formatOf,
  inspectDocument,
  readDocument,
} from '../../documents/source.js'
import { attachmentIdSchema, MAX_ATTACHMENT_BYTES } from './schema.js'
import type { ReadAttachmentInput } from './schema.js'

export class AttachmentError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'AttachmentError'
  }
}
const graphMetadataSchema = z
  .object({
    '@odata.type': z.string().max(128),
    id: attachmentIdSchema,
    name: z.string().min(1).max(512),
    contentType: z.string().max(256),
    size: z.number().int().nonnegative(),
    isInline: z.boolean(),
  })
  .passthrough()
type GraphMetadata = z.infer<typeof graphMetadataSchema>
const graphListSchema = z
  .object({
    value: z.array(graphMetadataSchema).max(50),
    '@odata.nextLink': z.string().max(8192).optional(),
  })
  .passthrough()
const graphFileSchema = graphMetadataSchema
  .extend({
    '@odata.type': z.literal('#microsoft.graph.fileAttachment'),
    size: z.number().int().min(1).max(MAX_ATTACHMENT_BYTES),
    contentBytes: z
      .string()
      .min(4)
      .max(4 * Math.ceil(MAX_ATTACHMENT_BYTES / 3)),
  })
  .passthrough()
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
    const supportedFormat = formatOf(metadata.name, metadata.contentType)
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
  const hasMore =
    parsed.data['@odata.nextLink'] !== undefined && parsed.data['@odata.nextLink'] !== ''
  if (hasMore && attachments.length === 0)
    throw new AttachmentError('Invalid attachment pagination response')
  const next = skip + attachments.length
  return { messageId, attachments, hasMore, nextOffset: hasMore && next <= 10000 ? next : null }
}
export function metadataFromGraph(data: GraphMetadata) {
  const attachmentType = typeOf(data['@odata.type'])
  const supportedFormat = formatOf(data.name, data.contentType)
  return {
    attachmentId: data.id,
    name: data.name,
    contentType: data.contentType,
    size: data.size,
    isInline: data.isInline,
    attachmentType,
    supportedFormat,
    readable:
      attachmentType === 'file' &&
      supportedFormat !== null &&
      data.size > 0 &&
      data.size <= MAX_ATTACHMENT_BYTES,
  }
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
  const format = formatOf(data.name, data.contentType)
  if (!format)
    throw new AttachmentError(
      'Only PDF, DOCX and XLSX file attachments with matching media types are supported',
    )
  return decodeBase64Document(
    {
      messageId: target.messageId,
      attachmentId: data.id,
      name: data.name,
      contentType: data.contentType,
      size: data.size,
      format,
    },
    data.contentBytes,
  )
}
export async function inspectAttachment(
  client: PowerAutomateClient,
  target: { messageId: string; attachmentId: string },
) {
  return await inspectDocument(await retrieve(client, target))
}
export async function readAttachment(client: PowerAutomateClient, input: ReadAttachmentInput) {
  return await readDocument(
    await retrieve(client, { messageId: input.messageId, attachmentId: input.attachmentId }),
    input.selection,
  )
}
export const attachmentResult = boundedToolResult
