import { z } from 'zod'

// Security boundary: control characters, including NEL U+0085, are deliberately excluded from IDs.
// oxlint-disable-next-line no-control-regex
const safeIdPattern = /^(?!\.{1,2}$)[^\s\u0000-\u001f\u007f\u0085]+$(?![\s\S])/

export const recipientSchema = z.object({ name: z.string(), address: z.string() }).strict()
export type Recipient = z.infer<typeof recipientSchema>

export const messageIdSchema = z.string().min(1).max(2048).regex(safeIdPattern)
export const mailboxScopeSchema = z.enum(['inbox', 'sent', 'all'])

const isoStringSchema = z.iso.datetime()

/** Compare validated UTC timestamps without losing Graph's sub-millisecond precision. */
export function compareUtcTimestamps(left: string, right: string): number {
  const milliseconds = Date.parse(left) - Date.parse(right)
  if (milliseconds !== 0) return milliseconds
  const leftFraction = /\.(\d+)Z$/.exec(left)?.[1] ?? ''
  const rightFraction = /\.(\d+)Z$/.exec(right)?.[1] ?? ''
  const precision = Math.max(leftFraction.length, rightFraction.length)
  const a = leftFraction.padEnd(precision, '0')
  const b = rightFraction.padEnd(precision, '0')
  return a < b ? -1 : a > b ? 1 : 0
}
export const listFiltersSchema = z
  .object({
    isRead: z.boolean().optional(),
    hasAttachments: z.boolean().optional(),
    receivedAfter: isoStringSchema.min(20).max(28).optional(),
    receivedBefore: isoStringSchema.min(20).max(28).optional(),
  })
  .strict()
  .refine(
    (value) =>
      value.receivedAfter === undefined ||
      value.receivedBefore === undefined ||
      compareUtcTimestamps(value.receivedAfter, value.receivedBefore) < 0,
    'receivedAfter must be before receivedBefore',
  )

const messageBaseSchema = z
  .object({
    id: z.string(),
    subject: z.string(),
    from: recipientSchema,
    receivedDateTime: isoStringSchema,
    sentDateTime: isoStringSchema.optional(),
    parentFolderId: z.string().optional(),
    conversationId: z.string().optional(),
    hasAttachments: z.boolean(),
    importance: z.enum(['low', 'normal', 'high']),
    isRead: z.boolean(),
  })
  .strict()

export const messageSummarySchema = messageBaseSchema.extend({ bodyPreview: z.string() }).strict()
export type MessageSummary = z.infer<typeof messageSummarySchema>

export const messageDetailSchema = messageBaseSchema
  .extend({
    bodyPreview: z.string().optional(),
    to: z.array(recipientSchema),
    cc: z.array(recipientSchema),
    body: z.object({ contentType: z.enum(['text', 'html']), content: z.string() }).strict(),
  })
  .strict()
export type MessageDetail = z.infer<typeof messageDetailSchema>

export const untrustedMessageDetailSchema = messageDetailSchema
  .extend({ untrustedContent: z.literal(true) })
  .strict()

export const cursorSchema = z
  .string()
  .min(1)
  .max(4096)
  .regex(/^[A-Za-z0-9_-]+$/)

export const listMessagesInputSchema = z
  .object({
    limit: z.number().int().min(1).max(50).default(5),
    mailbox: mailboxScopeSchema.default('inbox'),
    folderId: messageIdSchema.optional(),
    cursor: cursorSchema.optional(),
    filters: listFiltersSchema.optional(),
  })
  .strict()
export const listMessagesOutputSchema = z
  .object({
    messages: z.array(messageSummarySchema).max(50),
    hasMore: z.boolean(),
    nextCursor: z.string().nullable(),
    incompleteReason: z.string().nullable().optional(),
  })
  .strict()

export const searchMessagesInputSchema = z
  .object({
    query: z.string().trim().min(1).max(512),
    limit: z.number().int().min(1).max(50).default(10),
    mailbox: mailboxScopeSchema.default('inbox'),
    folderId: messageIdSchema.optional(),
  })
  .strict()
export const searchMessagesOutputSchema = z
  .object({
    messages: z.array(messageSummarySchema).max(50),
    hasMore: z.boolean(),
    nextCursor: z.null(),
    incompleteReason: z.string().nullable(),
  })
  .strict()

export const getMessageInputSchema = z.object({ messageId: messageIdSchema }).strict()
export const getMessageOutputSchema = untrustedMessageDetailSchema

export const listMailFoldersInputSchema = z
  .object({ limit: z.number().int().min(1).max(50).default(25) })
  .strict()
export const mailFolderSchema = z
  .object({
    id: messageIdSchema,
    displayName: z.string(),
    parentFolderId: messageIdSchema.nullable(),
    childFolderCount: z.number().int().nonnegative().nullable(),
    totalItemCount: z.number().int().nonnegative().nullable(),
    unreadItemCount: z.number().int().nonnegative().nullable(),
  })
  .strict()
export const listMailFoldersOutputSchema = z
  .object({
    folders: z.array(mailFolderSchema).max(50),
    hasMore: z.boolean(),
    incompleteReason: z.string().nullable(),
  })
  .strict()

export const getConversationInputSchema = z
  .object({
    conversationId: messageIdSchema,
    limit: z.number().int().min(1).max(50).default(20),
    cursor: cursorSchema.optional(),
  })
  .strict()
export const getConversationOutputSchema = z
  .object({
    conversationId: z.string(),
    messages: z.array(untrustedMessageDetailSchema).max(50),
    hasMore: z.boolean(),
    nextCursor: z.string().nullable(),
    incompleteReason: z.string().nullable(),
  })
  .strict()
