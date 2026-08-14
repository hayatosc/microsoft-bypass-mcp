/**
 * Zod schemas for normalized message shapes and MCP tool inputs/outputs.
 */
import { z } from 'zod'

/** A normalized email recipient (display name + address). */
export const recipientSchema = z.object({
  name: z.string(),
  address: z.string(),
})
/** Normalized email recipient. */
export type Recipient = z.infer<typeof recipientSchema>

const messageBaseSchema = z.object({
  id: z.string(),
  subject: z.string(),
  from: recipientSchema,
  receivedDateTime: z.iso.datetime(),
  hasAttachments: z.boolean(),
  importance: z.enum(['low', 'normal', 'high']),
  isRead: z.boolean(),
})

/** Summary of a mailbox message (metadata only, no full body). */
export const messageSummarySchema = messageBaseSchema.extend({
  bodyPreview: z.string(),
})
/** Normalized message summary. */
export type MessageSummary = z.infer<typeof messageSummarySchema>

/** A full mailbox message, including recipients and the body. */
export const messageDetailSchema = messageBaseSchema.extend({
  to: z.array(recipientSchema),
  cc: z.array(recipientSchema),
  body: z.object({
    contentType: z.enum(['text', 'html']),
    content: z.string(),
  }),
})
/** Normalized message detail. */
export type MessageDetail = z.infer<typeof messageDetailSchema>

/** Input schema for `outlook_list_messages`. */
export const listMessagesInputSchema = z.object({
  limit: z.number().int().min(1).max(50).default(5),
})
/** Output schema for `outlook_list_messages`. */
export const listMessagesOutputSchema = z.object({
  messages: z.array(messageSummarySchema),
  hasMore: z.boolean(),
})

/** Input schema for `outlook_search_messages`. */
export const searchMessagesInputSchema = z.object({
  query: z.string().min(1),
  limit: z.number().int().min(1).max(50).default(10),
})
/** Output schema for `outlook_search_messages`. */
export const searchMessagesOutputSchema = z.object({
  messages: z.array(messageSummarySchema),
  hasMore: z.boolean(),
})

/** Input schema for `outlook_get_message`. */
export const getMessageInputSchema = z.object({
  messageId: z.string().min(1),
})
/** Output schema for `outlook_get_message`. */
export const getMessageOutputSchema = messageDetailSchema
