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

it('retains valid present from and supported null name behavior across shared normalizers', () => {
  for (const name of ['Jane Doe', null]) {
    const message = {
      ...graphDetailItem,
      conversationId: 'conv-1',
      from: { emailAddress: { name, address: 'jane@example.com' } },
    }
    const from = { name: name ?? '', address: 'jane@example.com' }
    expect(normalizeMessage(message).from).toEqual(from)
    expect(normalizeMessageList({ value: [message] }).messages[0]?.from).toEqual(from)
    expect(normalizeConversation({ value: [message] }, 'conv-1').messages[0]?.from).toEqual(from)
  }
})

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
  it('maps omitted and null from identically without poisoning a mixed list/search page', () => {
    const { from: _unused, ...withoutFrom } = graphSummaryItem
    const draft = { ...withoutFrom, id: 'draft-1', sender: graphSummaryItem.from }
    const page = { value: [graphSummaryItem, draft] }
    const normalized = normalizeMessageList(page, 2)
    expect(normalized).toEqual(
      normalizeMessageList({ value: [graphSummaryItem, { ...draft, from: null }] }, 2),
    )
    expect(normalized).toEqual({
      messages: [summary, { ...summary, id: 'draft-1', from: { name: '', address: '' } }],
      hasMore: false,
      page: { nextLink: null, count: 2 },
    })
    expect(() => normalizeMessageList(page, 1)).toThrow('MESSAGE_LIST_LIMIT')
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
  it('maps omitted and null from identically without inferring from sender or recipients', () => {
    const { from: _unused, ...withoutFrom } = graphDetailItem
    const draft = { ...withoutFrom, sender: graphDetailItem.from }
    const normalized = normalizeMessage(draft, graphDetailItem.id)
    expect(normalized).toEqual(normalizeMessage({ ...draft, from: null }, graphDetailItem.id))
    expect(normalized).toEqual({ ...detail, from: { name: '', address: '' } })
    expect(() => normalizeMessage(draft, 'different-id')).toThrow('message ID does not match')
  })
  it('throws a clean error on a malformed message', () => {
    expect(() => normalizeMessage({ id: 'msg-1' })).toThrow(PowerAutomateError)
  })
})

describe('normalizeConversation', () => {
  it('maps omitted and null from identically in a mixed page while retaining all guards', () => {
    const message = { ...graphDetailItem, conversationId: 'conv-1' }
    const { from: _unused, ...withoutFrom } = message
    const draft = {
      ...withoutFrom,
      id: 'draft-1',
      receivedDateTime: '2025-01-02T09:00:00Z',
      sender: message.from,
    }
    const page = { value: [message, draft] }
    const normalized = normalizeConversation(page, 'conv-1', 2)
    expect(normalized).toEqual(
      normalizeConversation({ value: [message, { ...draft, from: null }] }, 'conv-1', 2),
    )
    expect(normalized).toEqual({
      messages: [
        {
          ...detail,
          id: draft.id,
          conversationId: 'conv-1',
          receivedDateTime: draft.receivedDateTime,
          from: { name: '', address: '' },
        },
        { ...detail, conversationId: 'conv-1' },
      ],
      hasMore: false,
      page: { nextLink: null, count: 2 },
    })
    expect(() => normalizeConversation(page, 'conv-1', 1)).toThrow('MESSAGE_LIST_LIMIT')
    expect(() => normalizeConversation(page, 'other')).toThrow('conversation message ID')
    expect(() => normalizeConversation({ value: [draft, draft] }, 'conv-1')).toThrow('duplicate')
  })
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
