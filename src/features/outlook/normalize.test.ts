import { describe, expect, it } from 'vitest'

import { PowerAutomateError } from '../../lib/power-automate.js'
import { normalizeConversation, normalizeMessage, normalizeMessageList } from './normalize.js'

const graphSummaryItem = {
  id: 'msg-1',
  subject: 'Weekly report',
  from: { emailAddress: { name: 'Jane Doe', address: 'jane@example.com' } },
  receivedDateTime: '2025-01-01T09:00:00Z',
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  bodyPreview: 'Hello world...',
} as const
const summary = {
  id: 'msg-1',
  subject: 'Weekly report',
  from: { name: 'Jane Doe', address: 'jane@example.com' },
  receivedDateTime: '2025-01-01T09:00:00Z',
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  bodyPreview: 'Hello world...',
}
const graphDetailItem = {
  ...graphSummaryItem,
  toRecipients: [{ emailAddress: { name: 'John Smith', address: 'john@example.com' } }],
  ccRecipients: [{ emailAddress: { name: 'Admin', address: 'admin@example.com' } }],
  body: { contentType: 'text', content: 'Hello world' },
} as const
const detail = {
  ...summary,
  to: [{ name: 'John Smith', address: 'john@example.com' }],
  cc: [{ name: 'Admin', address: 'admin@example.com' }],
  body: { contentType: 'text', content: 'Hello world' },
}

describe('normalizeMessageList', () => {
  it('maps a Graph list response to summaries with hasMore=false', () => {
    expect(normalizeMessageList({ value: [graphSummaryItem] })).toEqual({
      messages: [summary],
      hasMore: false,
      page: { nextLink: null, count: 1 },
    })
  })
  it('sets hasMore when @odata.nextLink is present without returning it as a tool field', () => {
    expect(
      normalizeMessageList({
        value: [graphSummaryItem],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=1',
      }),
    ).toEqual({
      messages: [summary],
      hasMore: true,
      page: { nextLink: 'https://graph.microsoft.com/v1.0/me/messages?$skip=1', count: 1 },
    })
  })
  it('maps a null sender to an empty recipient', () => {
    const withNullFrom = { ...graphSummaryItem, from: null }
    expect(normalizeMessageList({ value: [withNullFrom] }).messages).toEqual([
      { ...summary, from: { name: '', address: '' } },
    ])
  })
  it('throws a clean error on a malformed list response', () => {
    expect(() => normalizeMessageList({ value: [{ id: 1 }] })).toThrow(PowerAutomateError)
  })
})

describe('normalizeMessage', () => {
  it('maps a single Graph message to MessageDetail and preserves legacy missing metadata', () => {
    expect(normalizeMessage(graphDetailItem)).toEqual(detail)
  })
  it('does not require bodyPreview on Graph message detail', () => {
    const { bodyPreview: _unused, ...withoutPreview } = graphDetailItem
    expect(normalizeMessage(withoutPreview)).toEqual({ ...detail, bodyPreview: '' })
  })
  it('throws a clean error on a malformed message', () => {
    expect(() => normalizeMessage({ id: 'msg-1' })).toThrow(PowerAutomateError)
  })
})

describe('normalizeConversation', () => {
  it('requires exact returned conversation IDs and duplicate-free messages', () => {
    const message = { ...graphDetailItem, conversationId: 'conv-1' }
    expect(normalizeConversation({ value: [message] }, 'conv-1').messages).toHaveLength(1)
    expect(() =>
      normalizeConversation({ value: [{ ...message, conversationId: 'other' }] }, 'conv-1'),
    ).toThrow('conversation')
    expect(() => normalizeConversation({ value: [message, message] }, 'conv-1')).toThrow(
      'duplicate',
    )
  })
})
