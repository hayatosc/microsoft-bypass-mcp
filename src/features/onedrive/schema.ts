import { z } from 'zod'

import { MAX_ATTACHMENT_BYTES, MAX_OUTPUT_CHARACTERS } from '../documents/source.js'
import {
  inspectAttachmentOutputSchema,
  readAttachmentInputSchema,
  readAttachmentOutputSchema,
} from '../outlook/attachments/schema.js'
import { messageIdSchema } from '../outlook/schema.js'

export const oneDriveFileIdSchema = messageIdSchema
export const documentFormatSchema = z.enum(['pdf', 'docx', 'xlsx'])
export const oneDriveMetadataSchema = z
  .object({
    fileId: z.string(),
    name: z.string(),
    contentType: z.string(),
    size: z.number().int().nonnegative(),
    isFolder: z.boolean(),
    supportedFormat: documentFormatSchema.nullable(),
    readable: z.boolean(),
    lastModifiedDateTime: z.string().nullable(),
    eTag: z.string().nullable(),
    limitation: z.string().nullable(),
  })
  .strict()
export const oneDriveSearchInputSchema = z
  .object({
    query: z.string().trim().min(1).max(512),
    limit: z.number().int().min(1).max(100).default(10),
  })
  .strict()
export const oneDriveSearchOutputSchema = z
  .object({
    files: z.array(oneDriveMetadataSchema).max(100),
    hasMore: z.boolean(),
    nextCursor: z.null(),
    incompleteReason: z.string().nullable(),
    truncated: z.boolean(),
  })
  .strict()
export const oneDriveListFolderInputSchema = z
  .object({
    folderId: oneDriveFileIdSchema.optional(),
    limit: z.number().int().min(1).max(100).default(50),
  })
  .strict()
export const oneDriveListFolderOutputSchema = z
  .object({
    folderId: z.string().nullable(),
    files: z.array(oneDriveMetadataSchema).max(100),
    hasMore: z.boolean(),
    nextCursor: z.null(),
    incompleteReason: z.string().nullable(),
    truncated: z.boolean(),
  })
  .strict()
export const oneDriveMetadataInputSchema = z.object({ fileId: oneDriveFileIdSchema }).strict()
export const oneDriveMetadataOutputSchema = oneDriveMetadataSchema
export const oneDriveSourceSchema = z
  .object({
    provider: z.literal('onedrive'),
    fileId: z.string(),
    name: z.string(),
    contentType: z.string(),
    size: z.number().int().min(1).max(MAX_ATTACHMENT_BYTES),
    lastModifiedDateTime: z.string().nullable(),
    eTag: z.string().nullable(),
    format: documentFormatSchema,
  })
  .strict()
export const oneDriveInspectOutputSchema = inspectAttachmentOutputSchema
  .extend({ source: oneDriveSourceSchema })
  .strict()
export const oneDriveInspectInputSchema = oneDriveMetadataInputSchema
const baseSelectionSchema = readAttachmentInputSchema.shape.selection
export const oneDriveReadInputSchema = z
  .object({ fileId: oneDriveFileIdSchema, selection: baseSelectionSchema })
  .strict()
export const oneDriveReadOutputSchema = readAttachmentOutputSchema
  .extend({ source: oneDriveSourceSchema })
  .strict()
export type OneDriveReadInput = z.infer<typeof oneDriveReadInputSchema>
export type OneDriveSource = z.infer<typeof oneDriveSourceSchema>
export { MAX_ATTACHMENT_BYTES, MAX_OUTPUT_CHARACTERS }
