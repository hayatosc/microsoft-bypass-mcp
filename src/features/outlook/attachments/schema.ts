/** Bounded, format-specific attachment tool contracts. All returned content is untrusted. */
import { z } from 'zod'

export const MAX_ATTACHMENT_BYTES = 4 * 1024 * 1024
export const MAX_OUTPUT_CHARACTERS = 20000
export const attachmentIdSchema = z
  .string()
  .min(1)
  .max(2048)
  // Security boundary: control characters are deliberately excluded from Graph IDs.
  // oxlint-disable-next-line no-control-regex
  .regex(/^(?!\.{1,2}$)[^\s\u0000-\u001f\u007f]+$(?![\s\S])/)
export const attachmentTargetSchema = z.object({
  messageId: attachmentIdSchema,
  attachmentId: attachmentIdSchema,
})
export const listAttachmentsInputSchema = z.object({
  messageId: attachmentIdSchema,
  limit: z.number().int().min(1).max(50).default(20),
  offset: z.number().int().min(0).max(10000).default(0),
})
export const readAttachmentInputSchema = attachmentTargetSchema.extend({
  selection: z.discriminatedUnion('format', [
    z
      .object({
        format: z.literal('pdf'),
        pageStart: z.number().int().min(1).max(200),
        pageEnd: z.number().int().min(1).max(200),
        maxCharacters: z.number().int().min(1).max(MAX_OUTPUT_CHARACTERS).default(10000),
      })
      .refine(
        (v) => v.pageEnd >= v.pageStart && v.pageEnd - v.pageStart < 10,
        'Select an ascending range of at most 10 pages',
      ),
    z.object({
      format: z.literal('docx'),
      sectionId: z.string().min(1).max(64).optional(),
      offset: z.number().int().min(0).max(8000000).default(0),
      length: z.number().int().min(1).max(MAX_OUTPUT_CHARACTERS).default(10000),
    }),
    z.object({
      format: z.literal('xlsx'),
      sheet: z.string().min(1).max(128),
      range: z
        .string()
        .max(32)
        .regex(/^[A-Za-z]{1,3}[1-9][0-9]{0,6}(:[A-Za-z]{1,3}[1-9][0-9]{0,6})?$/),
    }),
  ]),
})
export type ReadAttachmentInput = z.infer<typeof readAttachmentInputSchema>
export const attachmentFormatSchema = z.enum(['pdf', 'docx', 'xlsx'])
export const attachmentMetadataSchema = z.object({
  attachmentId: attachmentIdSchema,
  name: z.string().max(512),
  contentType: z.string().max(256),
  size: z.number().int().nonnegative(),
  isInline: z.boolean(),
  attachmentType: z.enum(['file', 'item', 'reference', 'unknown']),
  supportedFormat: attachmentFormatSchema.nullable(),
  readable: z.boolean(),
})
export const listAttachmentsOutputSchema = z.object({
  messageId: attachmentIdSchema,
  attachments: z.array(attachmentMetadataSchema).max(50),
  hasMore: z.boolean(),
  nextOffset: z.number().int().nullable(),
})
export const sourceSchema = z.object({
  messageId: attachmentIdSchema,
  attachmentId: attachmentIdSchema,
  name: z.string().max(512),
  contentType: z.string().max(256),
  size: z.number().int().max(MAX_ATTACHMENT_BYTES),
  format: attachmentFormatSchema,
})
const warnings = z.array(z.string())
const dimensionsSchema = z.object({
  range: z.string().nullable(),
  firstRow: z.number().nullable(),
  lastRow: z.number().nullable(),
  firstColumn: z.number().nullable(),
  lastColumn: z.number().nullable(),
  cellCount: z.number(),
})
export const inspectAttachmentOutputSchema = z.object({
  source: sourceSchema,
  untrustedContent: z.literal(true),
  structure: z.discriminatedUnion('format', [
    z.object({ format: z.literal('pdf'), pageCount: z.number().int().min(1).max(200) }),
    z.object({
      format: z.literal('docx'),
      paragraphCount: z.number(),
      textCharacters: z.number(),
      sections: z.array(
        z.object({
          id: z.string(),
          title: z.string(),
          headingLevel: z.number().nullable(),
          paragraphStart: z.number(),
          paragraphEnd: z.number(),
          start: z.number(),
          end: z.number(),
        }),
      ),
      warnings,
    }),
    z.object({
      format: z.literal('xlsx'),
      sheets: z.array(z.object({ name: z.string(), dimensions: dimensionsSchema })),
      dateSystem: z.enum(['1900', '1904']),
      warnings,
    }),
  ]),
})
export const readAttachmentOutputSchema = z.object({
  source: sourceSchema,
  untrustedContent: z.literal(true),
  data: z.discriminatedUnion('format', [
    z.object({
      format: z.literal('pdf'),
      pageCount: z.number(),
      pageStart: z.number(),
      pageEnd: z.number(),
      pages: z.array(z.object({ page: z.number(), text: z.string(), truncated: z.boolean() })),
      truncated: z.boolean(),
    }),
    z.object({
      format: z.literal('docx'),
      sectionId: z.string().nullable(),
      offset: z.number(),
      length: z.number(),
      sourceStart: z.number(),
      sourceEnd: z.number(),
      totalCharacters: z.number(),
      nextOffset: z.number().nullable(),
      text: z.string(),
      warnings,
    }),
    z.object({
      format: z.literal('xlsx'),
      sheet: z.string(),
      range: z.string(),
      dimensions: dimensionsSchema,
      dateSystem: z.enum(['1900', '1904']),
      cells: z.array(
        z.object({
          address: z.string(),
          row: z.number(),
          column: z.number(),
          type: z.enum(['number', 'string', 'boolean', 'error', 'date', 'blank']),
          value: z.union([z.string(), z.boolean(), z.null()]),
          rawValue: z.string().nullable(),
          present: z.boolean(),
          hasFormula: z.boolean(),
          cachedValueMissing: z.boolean(),
        }),
      ),
      warnings,
    }),
  ]),
})
