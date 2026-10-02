/**
 * Normalization of unknown Graph message payloads (proxied by Power Automate)
 * into the tool output shapes. Malformed payloads raise PowerAutomateError
 * with a message that never contains the payload itself.
 */
import { z } from 'zod'

import { PowerAutomateError } from '../../lib/power-automate.js'
import type { MessageDetail, MessageSummary, Recipient } from './schema.js'

const graphRecipientSchema = z.object({
  emailAddress: z.object({
    // Graph returns null when a sender/recipient has no display name.
    name: z
      .string()
      .nullable()
      .transform((n) => n ?? ''),
    address: z.string(),
  }),
})

// Graph returns `from` as null for some system-generated messages.
const nullableGraphRecipientSchema = graphRecipientSchema
  .nullable()
  .transform((recipient) => recipient ?? { emailAddress: { name: '', address: '' } })

const graphMessageSummarySchema = z.object({
  id: z.string(),
  subject: z.string(),
  from: nullableGraphRecipientSchema,
  receivedDateTime: z.iso.datetime(),
  hasAttachments: z.boolean(),
  importance: z.enum(['low', 'normal', 'high']),
  isRead: z.boolean(),
  bodyPreview: z
    .string()
    .nullable()
    .transform((n) => n ?? ''),
})

const graphMessageDetailSchema = z.object({
  id: z.string(),
  subject: z.string(),
  from: nullableGraphRecipientSchema,
  toRecipients: z.array(graphRecipientSchema),
  ccRecipients: z.array(graphRecipientSchema),
  receivedDateTime: z.iso.datetime(),
  hasAttachments: z.boolean(),
  importance: z.enum(['low', 'normal', 'high']),
  isRead: z.boolean(),
  body: z.object({
    contentType: z.enum(['text', 'html']),
    content: z.string(),
  }),
})

const graphMessageListSchema = z.object({
  value: z.array(graphMessageSummarySchema),
  '@odata.nextLink': z.string().optional(),
})

type GraphRecipient = z.infer<typeof graphRecipientSchema>
type GraphMessageSummary = z.infer<typeof graphMessageSummarySchema>
type GraphMessageDetail = z.infer<typeof graphMessageDetailSchema>

function toRecipient(recipient: GraphRecipient): Recipient {
  return {
    name: recipient.emailAddress.name,
    address: recipient.emailAddress.address,
  }
}

function toSummary(message: GraphMessageSummary): MessageSummary {
  return {
    id: message.id,
    subject: message.subject,
    from: toRecipient(message.from),
    receivedDateTime: message.receivedDateTime,
    hasAttachments: message.hasAttachments,
    importance: message.importance,
    isRead: message.isRead,
    bodyPreview: message.bodyPreview,
  }
}

function toDetail(message: GraphMessageDetail): MessageDetail {
  return {
    id: message.id,
    subject: message.subject,
    from: toRecipient(message.from),
    to: message.toRecipients.map(toRecipient),
    cc: message.ccRecipients.map(toRecipient),
    receivedDateTime: message.receivedDateTime,
    hasAttachments: message.hasAttachments,
    importance: message.importance,
    isRead: message.isRead,
    body: { contentType: message.body.contentType, content: message.body.content },
  }
}

/** Normalized result of a Graph list response ({ value: [...] }). */
export interface MessageList {
  messages: MessageSummary[]
  hasMore: boolean
}

/** Normalizes a Graph list response into message summaries plus a hasMore flag. */
export function normalizeMessageList(body: unknown): MessageList {
  const parsed = graphMessageListSchema.safeParse(body)
  if (!parsed.success) {
    throw new PowerAutomateError('malformed response: expected a list of messages')
  }
  return {
    messages: parsed.data.value.map(toSummary),
    hasMore: parsed.data['@odata.nextLink'] !== undefined,
  }
}

/** Normalizes a single Graph message into a message detail. */
export function normalizeMessage(body: unknown): MessageDetail {
  const parsed = graphMessageDetailSchema.safeParse(body)
  if (!parsed.success) {
    throw new PowerAutomateError('malformed response: expected a message')
  }
  return toDetail(parsed.data)
}
