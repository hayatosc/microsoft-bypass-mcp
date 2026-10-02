import { z } from 'zod'

import type { PowerAutomateClient } from '../../lib/power-automate.js'
import {
  boundedToolResult,
  decodeBase64Document,
  formatOf,
  inspectDocument,
  readDocument,
} from '../documents/source.js'
import type { DocumentSelection } from '../documents/source.js'
import { MAX_ATTACHMENT_BYTES } from '../outlook/attachments/schema.js'
import type { OneDriveReadInput } from './schema.js'

export class OneDriveError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'OneDriveError'
  }
}
const nativeMetadataSchema = z
  .object({
    Id: z.string().min(1).max(2048),
    Name: z.string().min(1).max(512),
    NameNoExt: z.string().optional(),
    DisplayName: z.string().optional(),
    Path: z.string().optional(),
    FileLocator: z.unknown().optional(),
    Size: z.number().int().nonnegative(),
    MediaType: z.string().max(256),
    IsFolder: z.boolean(),
    LastModified: z.string().nullable().optional(),
    ETag: z.string().nullable().optional(),
  })
  .passthrough()
const nativePageSchema = z
  .object({
    value: z.array(nativeMetadataSchema).max(100),
    nextLink: z.string().nullable().optional(),
    truncated: z.boolean().optional(),
  })
  .passthrough()
const nativeListSchema = z.union([z.array(nativeMetadataSchema).max(100), nativePageSchema])
const nativeContentSchema = z
  .object({
    metadata: nativeMetadataSchema,
    contentBytes: z
      .string()
      .min(4)
      .max(4 * Math.ceil(MAX_ATTACHMENT_BYTES / 3)),
  })
  .strict()
type NativeMetadata = z.infer<typeof nativeMetadataSchema>

function actualBase64Bytes(contentBytes: string): number {
  if (contentBytes.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(contentBytes))
    throw new OneDriveError('Invalid OneDrive content response')
  const padding = contentBytes.endsWith('==') ? 2 : contentBytes.endsWith('=') ? 1 : 0
  return (contentBytes.length / 4) * 3 - padding
}
function parseItems(body: unknown): {
  items: NativeMetadata[]
  nativeHasMore: boolean
  truncated: boolean
} {
  const parsed = nativeListSchema.safeParse(body)
  if (!parsed.success) throw new OneDriveError('Invalid OneDrive metadata response')
  if (Array.isArray(parsed.data))
    return { items: parsed.data, nativeHasMore: false, truncated: false }
  return {
    items: parsed.data.value,
    nativeHasMore:
      parsed.data.nextLink !== undefined &&
      parsed.data.nextLink !== null &&
      parsed.data.nextLink !== '',
    truncated: parsed.data.truncated === true,
  }
}
function normalizeOneDriveItem(item: NativeMetadata) {
  const supportedFormat = item.IsFolder ? null : formatOf(item.Name, item.MediaType)
  const tooLarge = item.Size > MAX_ATTACHMENT_BYTES
  const readable = !item.IsFolder && supportedFormat !== null && item.Size > 0 && !tooLarge
  return {
    fileId: item.Id,
    name: item.Name,
    contentType: item.MediaType,
    size: item.Size,
    isFolder: item.IsFolder,
    supportedFormat,
    readable,
    lastModifiedDateTime: item.LastModified ?? null,
    eTag: item.ETag ?? null,
    limitation: item.IsFolder
      ? 'Folders must be listed, not read as documents.'
      : tooLarge
        ? 'Files over 4 MiB cannot be inspected or read.'
        : supportedFormat === null
          ? 'Only PDF, DOCX and XLSX files are supported.'
          : item.Size === 0
            ? 'Empty files cannot be inspected or read.'
            : null,
  }
}
export function normalizeOneDriveList(body: unknown, limit: number) {
  const parsed = parseItems(body)
  if (parsed.items.length > limit) throw new OneDriveError('Invalid OneDrive page response')
  const files = parsed.items.map(normalizeOneDriveItem)
  const hasMore = parsed.nativeHasMore || parsed.truncated
  return {
    files,
    hasMore,
    nextCursor: null,
    incompleteReason: hasMore
      ? parsed.truncated
        ? 'The OneDrive connector stopped at its bounded page capacity; additional matching data may exist, but no URL or cursor is exposed or accepted.'
        : 'The OneDrive connector returned a bounded page with additional native data; no URL or cursor is exposed or accepted.'
      : null,
    truncated: parsed.truncated,
  }
}
export function normalizeOneDriveMetadata(body: unknown) {
  const parsed = nativeMetadataSchema.safeParse(body)
  if (!parsed.success) throw new OneDriveError('Invalid OneDrive metadata response')
  return normalizeOneDriveItem(parsed.data)
}
export function normalizeOneDriveMetadataFor(body: unknown, fileId: string) {
  const metadata = normalizeOneDriveMetadata(body)
  if (metadata.fileId !== fileId) throw new OneDriveError('OneDrive metadata ID does not match')
  return metadata
}
function sourceFromMetadata(metadata: ReturnType<typeof normalizeOneDriveMetadata>) {
  if (metadata.isFolder) throw new OneDriveError('Folders cannot be inspected or read as documents')
  if (!metadata.supportedFormat || !metadata.readable)
    throw new OneDriveError(
      'Only PDF, DOCX and XLSX owned OneDrive files up to 4 MiB are supported',
    )
  return {
    provider: 'onedrive' as const,
    fileId: metadata.fileId,
    name: metadata.name,
    contentType: metadata.contentType,
    size: metadata.size,
    lastModifiedDateTime: metadata.lastModifiedDateTime,
    eTag: metadata.eTag,
    format: metadata.supportedFormat,
  }
}
function assertSame(
  before: ReturnType<typeof normalizeOneDriveMetadata>,
  after: ReturnType<typeof normalizeOneDriveMetadata>,
) {
  if (after.fileId !== before.fileId) throw new OneDriveError('OneDrive content ID does not match')
  if (
    after.name !== before.name ||
    after.contentType !== before.contentType ||
    after.isFolder !== before.isFolder ||
    after.size !== before.size ||
    after.supportedFormat !== before.supportedFormat
  )
    throw new OneDriveError('OneDrive file metadata changed during retrieval')
  if (
    (before.lastModifiedDateTime !== null || after.lastModifiedDateTime !== null) &&
    after.lastModifiedDateTime !== before.lastModifiedDateTime
  )
    throw new OneDriveError('OneDrive file metadata changed during retrieval')
  if ((before.eTag !== null || after.eTag !== null) && after.eTag !== before.eTag)
    throw new OneDriveError('OneDrive file metadata changed during retrieval')
}
async function retrieve(client: PowerAutomateClient, fileId: string) {
  const before = normalizeOneDriveMetadataFor(
    await client.call('onedrive_get_metadata', { fileId }),
    fileId,
  )
  const source = sourceFromMetadata(before)
  const raw = await client.call('onedrive_get_content', { fileId })
  const parsed = nativeContentSchema.safeParse(raw)
  if (!parsed.success) throw new OneDriveError('Invalid OneDrive content response')
  if (actualBase64Bytes(parsed.data.contentBytes) !== before.size)
    throw new OneDriveError('OneDrive content size does not match metadata')
  const after = normalizeOneDriveMetadataFor(parsed.data.metadata, fileId)
  assertSame(before, after)
  return decodeBase64Document(source, parsed.data.contentBytes)
}
function toDocumentSelection(input: OneDriveReadInput): DocumentSelection {
  return input.selection
}
export async function searchOneDrive(client: PowerAutomateClient, query: string, limit: number) {
  return normalizeOneDriveList(
    await client.call('onedrive_search_files', { query, top: limit }),
    limit,
  )
}
export async function listOneDriveFolder(
  client: PowerAutomateClient,
  folderId: string | undefined,
  limit: number,
) {
  return {
    folderId: folderId ?? null,
    ...normalizeOneDriveList(
      await client.call('onedrive_list_folder', {
        ...(folderId === undefined ? {} : { folderId }),
        top: limit,
      }),
      limit,
    ),
  }
}
export async function inspectOneDriveFile(client: PowerAutomateClient, fileId: string) {
  return await inspectDocument(await retrieve(client, fileId))
}
export async function readOneDriveFile(client: PowerAutomateClient, input: OneDriveReadInput) {
  return await readDocument(await retrieve(client, input.fileId), toDocumentSelection(input))
}
export const oneDriveResult = boundedToolResult
