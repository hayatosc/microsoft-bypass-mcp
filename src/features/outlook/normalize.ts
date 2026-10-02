import { z } from 'zod'

import { PowerAutomateError } from '../../lib/power-automate.js'
import type { MessageDetail, MessageSummary, Recipient } from './schema.js'
import { compareUtcTimestamps, messageIdSchema } from './schema.js'

const graphRecipientSchema = z
  .object({
    emailAddress: z
      .object({
        name: z
          .string()
          .nullable()
          .transform((name) => name ?? ''),
        address: z.string(),
      })
      .passthrough(),
  })
  .passthrough()
const nullableGraphRecipientSchema = graphRecipientSchema
  .nullable()
  .transform((recipient) => recipient ?? { emailAddress: { name: '', address: '' } })
const graphMessageSummarySchema = z
  .object({
    id: messageIdSchema,
    subject: z
      .string()
      .nullable()
      .transform((name) => name ?? ''),
    from: nullableGraphRecipientSchema,
    receivedDateTime: z.iso.datetime(),
    sentDateTime: z.iso.datetime().optional(),
    parentFolderId: messageIdSchema.optional(),
    conversationId: messageIdSchema.optional(),
    hasAttachments: z.boolean(),
    importance: z.enum(['low', 'normal', 'high']),
    isRead: z.boolean(),
    bodyPreview: z
      .string()
      .nullable()
      .optional()
      .transform((value) => value ?? ''),
  })
  .passthrough()
const graphMessageDetailSchema = graphMessageSummarySchema
  .extend({
    toRecipients: z.array(graphRecipientSchema).default([]),
    ccRecipients: z.array(graphRecipientSchema).default([]),
    body: z.object({ contentType: z.enum(['text', 'html']), content: z.string() }).passthrough(),
  })
  .passthrough()
const graphMessageListSchema = z
  .object({
    value: z.array(graphMessageSummarySchema).max(50),
    '@odata.nextLink': z.string().max(8192).optional(),
  })
  .passthrough()
const graphMessageDetailListSchema = z
  .object({
    value: z.array(graphMessageDetailSchema).max(50),
    '@odata.nextLink': z.string().max(8192).optional(),
  })
  .passthrough()
const graphFolderSchema = z
  .object({
    id: messageIdSchema,
    displayName: z.string(),
    parentFolderId: messageIdSchema.nullable().optional(),
    childFolderCount: z.number().int().nonnegative().nullable().optional(),
    totalItemCount: z.number().int().nonnegative().nullable().optional(),
    unreadItemCount: z.number().int().nonnegative().nullable().optional(),
  })
  .passthrough()
const graphFolderListSchema = z
  .object({
    value: z.array(graphFolderSchema).max(50),
    '@odata.nextLink': z.string().max(8192).optional(),
  })
  .passthrough()

type GraphRecipient = z.infer<typeof graphRecipientSchema>
type GraphMessageSummary = z.infer<typeof graphMessageSummarySchema>
type GraphMessageDetail = z.infer<typeof graphMessageDetailSchema>

function toRecipient(recipient: GraphRecipient): Recipient {
  return { name: recipient.emailAddress.name, address: recipient.emailAddress.address }
}
function summary(message: GraphMessageSummary): MessageSummary {
  return {
    id: message.id,
    subject: message.subject,
    from: toRecipient(message.from),
    receivedDateTime: message.receivedDateTime,
    ...(message.sentDateTime === undefined ? {} : { sentDateTime: message.sentDateTime }),
    ...(message.parentFolderId === undefined ? {} : { parentFolderId: message.parentFolderId }),
    ...(message.conversationId === undefined ? {} : { conversationId: message.conversationId }),
    hasAttachments: message.hasAttachments,
    importance: message.importance,
    isRead: message.isRead,
    bodyPreview: message.bodyPreview,
  }
}
function detail(message: GraphMessageDetail): MessageDetail {
  return {
    ...summary(message),
    to: message.toRecipients.map(toRecipient),
    cc: message.ccRecipients.map(toRecipient),
    body: { contentType: message.body.contentType, content: message.body.content },
  }
}
export interface PageState {
  nextLink: string | null
  count: number
}
export interface MessageList {
  messages: MessageSummary[]
  hasMore: boolean
  page: PageState
}
export interface ConversationList {
  messages: MessageDetail[]
  hasMore: boolean
  page: PageState
}

function nonEmptyNextLink(link: string | undefined): string | null {
  return link === undefined || link === '' ? null : link
}
export function normalizeMessageList(body: unknown, requestedLimit = 50): MessageList {
  const parsed = graphMessageListSchema.safeParse(body)
  if (!parsed.success || parsed.data.value.length > requestedLimit)
    throw new PowerAutomateError('malformed response: expected a list of messages')
  const nextLink = nonEmptyNextLink(parsed.data['@odata.nextLink'])
  return {
    messages: parsed.data.value.map(summary),
    hasMore: nextLink !== null,
    page: { nextLink, count: parsed.data.value.length },
  }
}
export function normalizeMessage(body: unknown, requestedId?: string): MessageDetail {
  const parsed = graphMessageDetailSchema.safeParse(body)
  if (!parsed.success) throw new PowerAutomateError('malformed response: expected a message')
  const message = detail(parsed.data)
  if (requestedId !== undefined && message.id !== requestedId)
    throw new PowerAutomateError('malformed response: message ID does not match')
  return message
}
export function normalizeConversation(
  body: unknown,
  conversationId: string,
  requestedLimit = 50,
): ConversationList {
  const parsed = graphMessageDetailListSchema.safeParse(body)
  if (!parsed.success || parsed.data.value.length > requestedLimit)
    throw new PowerAutomateError('malformed response: expected conversation messages')
  const messages = parsed.data.value.map(detail)
  if (messages.some((message) => message.conversationId !== conversationId))
    throw new PowerAutomateError('malformed response: conversation message ID does not match')
  const seen = new Set<string>()
  for (const message of messages) {
    if (seen.has(message.id))
      throw new PowerAutomateError('malformed response: duplicate conversation message')
    seen.add(message.id)
  }
  messages.sort(
    (a, b) =>
      compareUtcTimestamps(
        b.sentDateTime ?? b.receivedDateTime,
        a.sentDateTime ?? a.receivedDateTime,
      ) || a.id.localeCompare(b.id),
  )
  const nextLink = nonEmptyNextLink(parsed.data['@odata.nextLink'])
  return {
    messages,
    hasMore: nextLink !== null,
    page: { nextLink, count: parsed.data.value.length },
  }
}
export function normalizeMailFolders(body: unknown, requestedLimit = 50) {
  const parsed = graphFolderListSchema.safeParse(body)
  if (!parsed.success || parsed.data.value.length > requestedLimit)
    throw new PowerAutomateError('malformed response: expected a list of mail folders')
  return {
    folders: parsed.data.value.map((folder) => ({
      id: folder.id,
      displayName: folder.displayName,
      parentFolderId: folder.parentFolderId ?? null,
      childFolderCount: folder.childFolderCount ?? null,
      totalItemCount: folder.totalItemCount ?? null,
      unreadItemCount: folder.unreadItemCount ?? null,
    })),
    hasMore: nonEmptyNextLink(parsed.data['@odata.nextLink']) !== null,
  }
}
