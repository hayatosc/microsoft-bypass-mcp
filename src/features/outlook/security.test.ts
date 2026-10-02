import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { listFiltersSchema } from './schema.js'
import { createOutlookMcpServer } from './server.js'

const SELECT =
  'id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview'
const CANARY = 'SYNTHETIC_PRIVATE_CANARY'
const message = {
  id: 'message-1',
  subject: 'Synthetic subject',
  from: { emailAddress: { name: 'Synthetic sender', address: 'sender@example.test' } },
  receivedDateTime: '2026-01-01T12:00:00Z',
  sentDateTime: '2026-01-01T12:00:00Z',
  parentFolderId: 'inbox-id',
  conversationId: 'conversation-1',
  hasAttachments: false,
  importance: 'normal',
  isRead: false,
  bodyPreview: 'Synthetic preview',
  toRecipients: [],
  ccRecipients: [],
  body: { contentType: 'text', content: 'Synthetic body' },
}
const requestSchema = z.object({
  operation: z.string(),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})
type Request = z.infer<typeof requestSchema>
const clients: Client[] = []
afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()))
  vi.restoreAllMocks()
})

function nextLink(
  options: {
    path?: string
    top?: number
    skip?: string
    filter?: string
    conversation?: boolean
  } = {},
) {
  const url = new URL(
    options.path ??
      (options.conversation ? '/v1.0/me/messages' : '/v1.0/me/mailFolders/inbox/messages'),
    'https://graph.microsoft.com',
  )
  url.searchParams.set('$top', String(options.top ?? 3))
  url.searchParams.set('$skip', options.skip ?? '73')
  url.searchParams.set(
    '$select',
    options.conversation ? `${SELECT},toRecipients,ccRecipients,body` : SELECT,
  )
  if (options.filter !== undefined) url.searchParams.set('$filter', options.filter)
  else if (!options.conversation) url.searchParams.set('$orderby', 'receivedDateTime desc')
  return url.toString()
}
function changeLink(change: (url: URL) => void) {
  const url = new URL(nextLink())
  change(url)
  return url.toString()
}
async function flow(respond: (request: Request, index: number) => unknown) {
  const records: Request[] = []
  const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
    const request = requestSchema.parse(
      JSON.parse(typeof init?.body === 'string' ? init.body : '{}'),
    )
    records.push(request)
    const data = respond(request, records.length - 1)
    return Response.json({
      ok: true,
      requestId: request.requestId,
      operation: request.operation,
      data,
    })
  })
  const client = new PowerAutomateClient({
    baseUrl: `https://example.test/${CANARY}`,
    gatewayKey: CANARY,
    fetchFn,
  })
  const mcp = new Client({ name: 'security-tests', version: '1' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), createOutlookMcpServer(client).connect(b)])
  clients.push(mcp)
  return { mcp, records, fetchFn }
}
function call(mcp: Client, name: string, args: Record<string, unknown>) {
  return mcp.request({ method: 'tools/call', params: { name, arguments: args } })
}
function cursorFrom(result: { [key: string]: unknown }): string {
  return z.object({ nextCursor: z.string() }).parse(result.structuredContent).nextCursor
}
function decodeCursor(cursor: string): Record<string, unknown> {
  return z
    .record(z.string(), z.unknown())
    .parse(JSON.parse(atob(cursor.replace(/-/g, '+').replace(/_/g, '/'))))
}
function encodeCursor(payload: Record<string, unknown>) {
  return btoa(JSON.stringify(payload)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

describe('Outlook read-tool security boundaries', () => {
  it('continues at the validated Graph skip rather than the returned message count', async () => {
    const { mcp, records, fetchFn } = await flow((_request, index) => ({
      value: [message],
      ...(index === 0 ? { '@odata.nextLink': nextLink() } : {}),
    }))
    const first = await call(mcp, 'outlook_list_messages', { limit: 3 })
    const cursor = cursorFrom(first)
    expect(decodeCursor(cursor)).toMatchObject({ operation: 'list_messages', limit: 3, skip: 73 })
    const second = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(records[1]?.args).toEqual({ top: 3, skip: 73, mailbox: 'inbox' })
    expect(second.structuredContent).toMatchObject({ hasMore: false, nextCursor: null })
    expect(fetchFn.mock.calls.map(([url]) => url)).toEqual([
      `https://example.test/${CANARY}`,
      `https://example.test/${CANARY}`,
    ])
    expect(JSON.stringify(first)).not.toContain('graph.microsoft.com')
    expect(JSON.stringify(decodeCursor(cursor))).not.toContain('Synthetic')
  })

  it.each([
    ['mailbox', 'outlook_list_messages', { limit: 3, mailbox: 'sent' }],
    ['folder', 'outlook_list_messages', { limit: 3, folderId: 'different-folder' }],
    ['filters', 'outlook_list_messages', { limit: 3, filters: { isRead: false } }],
    ['limit', 'outlook_list_messages', { limit: 4 }],
    ['operation', 'outlook_get_conversation', { limit: 3, conversationId: 'conversation-1' }],
  ] as const)(
    'binds a continuation to its original %s before calling the flow',
    async (_label, name, args) => {
      const { mcp, fetchFn } = await flow(() => ({
        value: [message],
        '@odata.nextLink': nextLink(),
      }))
      const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
      const result = await call(mcp, name, { ...args, cursor })
      expect(result.isError).toBe(true)
      expect(result.content).toEqual([{ type: 'text', text: 'Invalid continuation cursor' }])
      expect(fetchFn).toHaveBeenCalledOnce()
    },
  )

  it.each([
    ['additional field', { url: `https://example.test/${CANARY}` }],
    ['wrong version', { v: 2 }],
    ['zero skip', { skip: 0 }],
    ['negative skip', { skip: -1 }],
    ['fractional skip', { skip: 1.5 }],
    ['skip beyond cap', { skip: 10001 }],
    ['non-integer limit', { limit: 3.5 }],
  ] as const)('rejects cursor payload with %s before calling the flow', async (_label, patch) => {
    const { mcp, fetchFn } = await flow(() => ({ value: [message], '@odata.nextLink': nextLink() }))
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const invalid = encodeCursor({ ...decodeCursor(cursor), ...patch })
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor: invalid })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('Invalid continuation cursor')
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(fetchFn).toHaveBeenCalledOnce()
  })

  it('rejects a non-canonical cursor representation even when its values match', async () => {
    const { mcp, fetchFn } = await flow(() => ({ value: [message], '@odata.nextLink': nextLink() }))
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const payload = decodeCursor(cursor)
    const reordered = encodeCursor(Object.fromEntries(Object.entries(payload).reverse()))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor: reordered })
    expect(result.isError).toBe(true)
    expect(fetchFn).toHaveBeenCalledOnce()
  })

  it.each([
    [
      'external origin',
      changeLink((url) => {
        url.hostname = 'evil.example.test'
      }),
    ],
    [
      'HTTP scheme',
      changeLink((url) => {
        url.protocol = 'http:'
      }),
    ],
    [
      'URL username',
      changeLink((url) => {
        url.username = CANARY
      }),
    ],
    [
      'URL password',
      changeLink((url) => {
        url.password = CANARY
      }),
    ],
    [
      'non-default port',
      changeLink((url) => {
        url.port = '444'
      }),
    ],
    [
      'fragment',
      changeLink((url) => {
        url.hash = CANARY
      }),
    ],
    ['wrong route', nextLink({ path: '/v1.0/me/messages' })],
    ['wrong top', nextLink({ top: 4 })],
    [
      'unknown query',
      changeLink((url) => {
        url.searchParams.set('$search', CANARY)
      }),
    ],
    [
      'duplicate skip',
      changeLink((url) => {
        url.searchParams.append('$skip', '74')
      }),
    ],
    [
      'missing select',
      changeLink((url) => {
        url.searchParams.delete('$select')
      }),
    ],
    [
      'changed select',
      changeLink((url) => {
        url.searchParams.set('$select', 'id,body')
      }),
    ],
    [
      'changed ordering',
      changeLink((url) => {
        url.searchParams.set('$orderby', 'sentDateTime asc')
      }),
    ],
    [
      'unexpected filter',
      changeLink((url) => {
        url.searchParams.set('$filter', 'isRead eq true')
      }),
    ],
    ['non-progressing skip', nextLink({ skip: '0' })],
    ['skip beyond cap', nextLink({ skip: '10001' })],
    ['fractional skip', nextLink({ skip: '2.5' })],
    ['skip token', nextLink({ skip: CANARY })],
    ['relative URL', '/v1.0/me/messages?$skip=73'],
  ])('rejects a nextLink with %s without following it', async (_label, link) => {
    const { mcp, fetchFn } = await flow(() => ({ value: [message], '@odata.nextLink': link }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3 })
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([{ type: 'text', text: 'Invalid pagination response' }])
    expect(result.structuredContent).toBeUndefined()
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
  })

  it('rejects a nextLink that replays the current cursor skip', async () => {
    const { mcp, records } = await flow(() => ({ value: [message], '@odata.nextLink': nextLink() }))
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(records[1]?.args.skip).toBe(73)
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('Invalid pagination response')
  })

  it('accepts the maximum continuation skip and an empty final page', async () => {
    const { mcp, records } = await flow((_request, index) =>
      index === 0
        ? { value: [message], '@odata.nextLink': nextLink({ skip: '10000' }) }
        : { value: [], '@odata.nextLink': '' },
    )
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(records[1]?.args.skip).toBe(10000)
    expect(result.structuredContent).toMatchObject({
      messages: [],
      hasMore: false,
      nextCursor: null,
    })
  })

  it('rejects an empty page claiming continuation', async () => {
    const { mcp } = await flow(() => ({ value: [], '@odata.nextLink': nextLink() }))
    expect((await call(mcp, 'outlook_list_messages', { limit: 3 })).isError).toBe(true)
  })

  it('validates filtered folder pagination with the exact escaped route and no ordering', async () => {
    const folderId = "資料/A'B+="
    const filters = { isRead: false, hasAttachments: true }
    const link = nextLink({
      path: `/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages`,
      filter: 'isRead eq false and hasAttachments eq true',
    })
    const { mcp, records } = await flow((_request, index) => ({
      value: [message],
      ...(index === 0 ? { '@odata.nextLink': link } : {}),
    }))
    const args = { limit: 3, mailbox: 'all', folderId, filters }
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', args))
    expect((await call(mcp, 'outlook_list_messages', { ...args, cursor })).isError).not.toBe(true)
    expect(records[1]?.args).toEqual({ top: 3, skip: 73, mailbox: 'all', folderId, filters })
  })

  it('round-trips a Unicode conversation ID with OData quotes without embedding it in the cursor', async () => {
    const conversationId = "会話'😀/+=é"
    const link = nextLink({
      conversation: true,
      filter: `conversationId eq '${conversationId.replace(/'/g, "''")}'`,
    })
    const { mcp, records } = await flow((_request, index) => ({
      value: [{ ...message, conversationId }],
      ...(index === 0 ? { '@odata.nextLink': link } : {}),
    }))
    const args = { conversationId, limit: 3 }
    const first = await call(mcp, 'outlook_get_conversation', args)
    const cursor = cursorFrom(first)
    expect(JSON.stringify(decodeCursor(cursor))).not.toContain(conversationId)
    const second = await call(mcp, 'outlook_get_conversation', { ...args, cursor })
    expect(second.structuredContent).toMatchObject({
      conversationId,
      messages: [{ conversationId, untrustedContent: true }],
      nextCursor: null,
    })
    expect(records[1]?.args).toEqual({ conversationId, top: 3, skip: 73 })
  })

  it.each(['conversation-10', 'CONVERSATION-1', undefined])(
    'rejects a conversation response with non-exact ID %s',
    async (conversationId) => {
      const { mcp } = await flow(() => ({ value: [{ ...message, conversationId }] }))
      const result = await call(mcp, 'outlook_get_conversation', {
        conversationId: 'conversation-1',
      })
      expect(result.isError).toBe(true)
      expect(JSON.stringify(result)).toContain('conversation message ID does not match')
      expect(result.structuredContent).toBeUndefined()
    },
  )

  it('rejects duplicate message IDs in a conversation page', async () => {
    const { mcp } = await flow(() => ({
      value: [message, { ...message, subject: 'Different copy' }],
    }))
    const result = await call(mcp, 'outlook_get_conversation', { conversationId: 'conversation-1' })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('duplicate conversation message')
  })

  it('sorts fractional timestamps numerically with a stable ID tie-breaker', async () => {
    const { mcp } = await flow(() => ({
      value: [
        { ...message, id: 'zero', sentDateTime: '2026-01-01T12:00:00Z' },
        { ...message, id: 'tie-b', sentDateTime: '2026-01-01T12:00:00.9Z' },
        { ...message, id: 'middle', sentDateTime: '2026-01-01T12:00:00.100Z' },
        { ...message, id: 'tie-a', sentDateTime: '2026-01-01T12:00:00.900Z' },
        {
          ...message,
          id: 'fallback',
          sentDateTime: undefined,
          receivedDateTime: '2026-01-01T12:00:01Z',
        },
      ],
    }))
    const result = await call(mcp, 'outlook_get_conversation', { conversationId: 'conversation-1' })
    const parsed = z
      .object({ messages: z.array(z.object({ id: z.string() })) })
      .parse(result.structuredContent)
    expect(parsed.messages.map((item) => item.id)).toEqual([
      'fallback',
      'tie-a',
      'tie-b',
      'middle',
      'zero',
    ])
  })

  it.each([
    ['ascending interval', '2026-01-01T12:00:00Z', '2026-01-01T12:00:00.100Z', true],
    ['descending interval', '2026-01-01T12:00:00.100Z', '2026-01-01T12:00:00Z', false],
    ['equivalent instants', '2026-01-01T12:00:00.000Z', '2026-01-01T12:00:00Z', false],
  ] as const)(
    'compares filter bounds by instant for an %s',
    (_label, receivedAfter, receivedBefore, accepted) => {
      expect(listFiltersSchema.safeParse({ receivedAfter, receivedBefore }).success).toBe(accepted)
    },
  )

  it('preserves accepted sub-millisecond precision in filter bounds', () => {
    expect(
      listFiltersSchema.safeParse({
        receivedAfter: '2026-01-01T12:00:00.1000001Z',
        receivedBefore: '2026-01-01T12:00:00.1000002Z',
      }).success,
    ).toBe(true)
  })

  it('preserves accepted sub-millisecond precision when sorting conversation messages', async () => {
    const { mcp } = await flow(() => ({
      value: [
        { ...message, id: 'earlier-a', sentDateTime: '2026-01-01T12:00:00.1000001Z' },
        { ...message, id: 'later-z', sentDateTime: '2026-01-01T12:00:00.1000002Z' },
      ],
    }))
    const result = await call(mcp, 'outlook_get_conversation', { conversationId: 'conversation-1' })
    const parsed = z
      .object({ messages: z.array(z.object({ id: z.string() })) })
      .parse(result.structuredContent)
    expect(parsed.messages.map((item) => item.id)).toEqual(['later-z', 'earlier-a'])
  })

  it('enforces the requested page length before returning messages', async () => {
    const { mcp } = await flow(() => ({ value: [message, { ...message, id: 'message-2' }] }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 1 })
    expect(result.isError).toBe(true)
    expect(result.structuredContent).toBeUndefined()
  })

  it('bounds serialized UTF-8 output and does not return oversized message bodies', async () => {
    const { mcp } = await flow(() => ({
      ...message,
      body: { contentType: 'text', content: '日'.repeat(45000) + CANARY },
    }))
    const result = await call(mcp, 'outlook_get_message', { messageId: 'message-1' })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('output limit')
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(result.structuredContent).toBeUndefined()
  })

  it('sanitizes upstream failures and logs only the documented diagnostic fields', async () => {
    const logs = vi.spyOn(console, 'log').mockImplementation(() => {})
    const { mcp } = await flow(() => {
      throw new Error(`${CANARY} https://graph.microsoft.com/private`)
    })
    const result = await call(mcp, 'outlook_get_conversation', { conversationId: CANARY })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
    expect(logs).toHaveBeenCalledOnce()
    const log = z
      .record(z.string(), z.unknown())
      .parse(JSON.parse(z.string().parse(logs.mock.calls[0]?.[0])))
    expect(Object.keys(log).sort()).toEqual([
      'durationMs',
      'operation',
      'requestId',
      'status',
      'success',
      'type',
    ])
    expect(log).toMatchObject({ operation: 'get_conversation', success: false })
    expect(JSON.stringify(logs.mock.calls)).not.toContain(CANARY)
  })
})
