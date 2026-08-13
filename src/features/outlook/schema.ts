/**
 * Zod schemas for normalized message shapes and MCP tool inputs/outputs.
 */
import { z } from 'zod'

export const recipientSchema = z.object({
  name: z.string(),
  address: z.string(),
})
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

export const messageSummarySchema = messageBaseSchema.extend({
  bodyPreview: z.string(),
})
export type MessageSummary = z.infer<typeof messageSummarySchema>

export const messageDetailSchema = messageBaseSchema.extend({
  to: z.array(recipientSchema),
  cc: z.array(recipientSchema),
  body: z.object({
    contentType: z.enum(['text', 'html']),
    content: z.string(),
  }),
})
export type MessageDetail = z.infer<typeof messageDetailSchema>

export const listMessagesInputSchema = z.object({
  limit: z.number().int().min(1).max(50).default(5),
})
export const listMessagesOutputSchema = z.object({
  messages: z.array(messageSummarySchema),
  hasMore: z.boolean(),
})

export const searchMessagesInputSchema = z.object({
  query: z.string().min(1),
  limit: z.number().int().min(1).max(50).default(10),
})
export const searchMessagesOutputSchema = z.object({
  messages: z.array(messageSummarySchema),
  hasMore: z.boolean(),
})

export const getMessageInputSchema = z.object({
  messageId: z.string().min(1),
})
export const getMessageOutputSchema = messageDetailSchema
