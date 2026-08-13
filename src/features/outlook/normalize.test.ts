import { describe, expect, it } from 'vitest'

import { PowerAutomateError } from '../../lib/power-automate.js'
import { normalizeMessage, normalizeMessageList } from './normalize.js'

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
  id: 'msg-1',
  subject: 'Weekly report',
  from: { emailAddress: { name: 'Jane Doe', address: 'jane@example.com' } },
  toRecipients: [{ emailAddress: { name: 'John Smith', address: 'john@example.com' } }],
  ccRecipients: [{ emailAddress: { name: 'Admin', address: 'admin@example.com' } }],
  receivedDateTime: '2025-01-01T09:00:00Z',
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  body: { contentType: 'text', content: 'Hello world' },
} as const

const detail = {
  id: 'msg-1',
  subject: 'Weekly report',
  from: { name: 'Jane Doe', address: 'jane@example.com' },
  to: [{ name: 'John Smith', address: 'john@example.com' }],
  cc: [{ name: 'Admin', address: 'admin@example.com' }],
  receivedDateTime: '2025-01-01T09:00:00Z',
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  body: { contentType: 'text', content: 'Hello world' },
}

describe('normalizeMessageList', () => {
  it('maps a Graph list response to summaries with hasMore=false', () => {
    expect(normalizeMessageList({ value: [graphSummaryItem] })).toEqual({
      messages: [summary],
      hasMore: false,
    })
  })

  it('sets hasMore when @odata.nextLink is present', () => {
    expect(
      normalizeMessageList({ value: [graphSummaryItem], '@odata.nextLink': 'https://graph/' }),
    ).toEqual({ messages: [summary], hasMore: true })
  })

  it('throws a clean error on a malformed list response', () => {
    expect(() => normalizeMessageList({ value: [{ id: 1 }] })).toThrow(PowerAutomateError)
  })
})

describe('normalizeMessage', () => {
  it('maps a single Graph message to MessageDetail', () => {
    expect(normalizeMessage(graphDetailItem)).toEqual(detail)
  })

  it('throws a clean error on a malformed message', () => {
    expect(() => normalizeMessage({ id: 'msg-1' })).toThrow(PowerAutomateError)
  })
})
