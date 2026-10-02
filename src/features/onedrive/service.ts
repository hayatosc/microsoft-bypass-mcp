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
import {
  ONE_DRIVE_FOLDER_WINDOW_SIZE,
  oneDriveFileIdSchema,
  oneDriveFolderCursorSchema,
  oneDriveListFolderInputSchema,
} from './schema.js'
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
const folderMetadataSchema = nativeMetadataSchema.extend({
  Id: oneDriveFileIdSchema,
  LastModified: z.string().max(2048).nullable().optional(),
  ETag: z.string().max(2048).nullable().optional(),
})
const nativeFolderListSchema = z.union([
  z.array(folderMetadataSchema).max(ONE_DRIVE_FOLDER_WINDOW_SIZE),
  nativePageSchema.extend({
    value: z.array(folderMetadataSchema).max(ONE_DRIVE_FOLDER_WINDOW_SIZE),
  }),
])
const folderCursorPayloadSchema = z
  .object({
    v: z.literal(1),
    operation: z.literal('onedrive_list_folder'),
    scope: z
      .string()
      .length(64)
      .regex(/^[a-f0-9]{64}$/),
    window: z
      .string()
      .length(64)
      .regex(/^[a-f0-9]{64}$/),
    offset: z
      .number()
      .int()
      .min(1)
      .max(ONE_DRIVE_FOLDER_WINDOW_SIZE - 1),
    limit: z.number().int().min(1).max(100),
  })
  .strict()
type FolderCursorPayload = z.infer<typeof folderCursorPayloadSchema>
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
function parseItems(
  body: unknown,
  schema = nativeListSchema,
): {
  items: NativeMetadata[]
  nativeHasMore: boolean
  truncated: boolean
} {
  const parsed = schema.safeParse(body)
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
async function sha256(text: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
function encodeFolderCursor(payload: FolderCursorPayload): string {
  // Fixed property order and ASCII-only payload make the encoding canonical.
  const { v, operation, scope, window, offset, limit } = payload
  return btoa(JSON.stringify({ v, operation, scope, window, offset, limit }))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')
}
function decodeFolderCursor(token: string, scope: string, limit: number): FolderCursorPayload {
  let payload: unknown
  try {
    const encoded = oneDriveFolderCursorSchema.parse(token).replace(/-/g, '+').replace(/_/g, '/')
    payload = JSON.parse(atob(encoded + '='.repeat((4 - (encoded.length % 4)) % 4)))
  } catch {
    throw new OneDriveError('Invalid OneDrive folder cursor')
  }
  const parsed = folderCursorPayloadSchema.safeParse(payload)
  if (
    !parsed.success ||
    parsed.data.scope !== scope ||
    parsed.data.limit !== limit ||
    encodeFolderCursor(parsed.data) !== token
  )
    throw new OneDriveError('Invalid OneDrive folder cursor for this request')
  return parsed.data
}
export async function listOneDriveFolder(
  client: PowerAutomateClient,
  folderId: string | undefined,
  limit: number,
  cursor?: string,
) {
  if (!oneDriveListFolderInputSchema.safeParse({ folderId, limit, cursor }).success)
    throw new OneDriveError('Invalid OneDrive folder request')
  const scope = await sha256(JSON.stringify({ folderId: folderId ?? null, limit }))
  // A cursor is an unkeyed read continuation, not authorization. Validate scope
  // before fetching; it never selects an upstream URL, folder or query itself.
  const continuation = cursor === undefined ? null : decodeFolderCursor(cursor, scope, limit)
  const parsed = parseItems(
    await client.call('onedrive_list_folder', {
      ...(folderId === undefined ? {} : { folderId }),
      top: ONE_DRIVE_FOLDER_WINDOW_SIZE,
    }),
    nativeFolderListSchema,
  )
  if (new Set(parsed.items.map((item) => item.Id)).size !== parsed.items.length)
    throw new OneDriveError('Invalid OneDrive folder response: duplicate item IDs')
  const windowFiles = parsed.items.map(normalizeOneDriveItem)
  // Hash the entire normalized, ordered window, not just the current MCP page.
  // Re-fetching and comparison avoid persistence and reject stale/reordered data.
  const window = await sha256(JSON.stringify(windowFiles))
  if (continuation !== null && continuation.window !== window)
    throw new OneDriveError(
      'OneDrive folder window changed; restart from the first page without a cursor',
    )
  const offset = continuation?.offset ?? 0
  if (continuation !== null && offset >= windowFiles.length)
    throw new OneDriveError('Invalid OneDrive folder cursor offset; restart from the first page')
  const files = windowFiles.slice(offset, offset + limit)
  const nextOffset = offset + files.length
  const localRemaining = nextOffset < windowFiles.length
  const capped = windowFiles.length >= ONE_DRIVE_FOLDER_WINDOW_SIZE
  const truncated = parsed.truncated || capped
  const upstreamIncomplete = truncated || parsed.nativeHasMore
  return {
    folderId: folderId ?? null,
    files,
    hasMore: localRemaining || upstreamIncomplete,
    nextCursor: localRemaining
      ? encodeFolderCursor({
          v: 1,
          operation: 'onedrive_list_folder',
          scope,
          window,
          offset: nextOffset,
          limit,
        })
      : null,
    incompleteReason: upstreamIncomplete
      ? 'The OneDrive folder window may be incomplete: only up to 1,000 items are available. Native folder pagination uses a minimum threshold, may overshoot on its final page, and is bounded before transport; root listing uses only its returned array. Cursors continue only within this window, not beyond it.'
      : null,
    truncated,
  }
}
export async function inspectOneDriveFile(client: PowerAutomateClient, fileId: string) {
  return await inspectDocument(await retrieve(client, fileId))
}
export async function readOneDriveFile(client: PowerAutomateClient, input: OneDriveReadInput) {
  return await readDocument(await retrieve(client, input.fileId), toDocumentSelection(input))
}
export const oneDriveResult = boundedToolResult
