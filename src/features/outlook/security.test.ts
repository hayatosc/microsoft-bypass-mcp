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
const messageSummary = {
  id: message.id,
  subject: message.subject,
  from: { name: 'Synthetic sender', address: 'sender@example.test' },
  receivedDateTime: message.receivedDateTime,
  sentDateTime: message.sentDateTime,
  parentFolderId: message.parentFolderId,
  conversationId: message.conversationId,
  hasAttachments: message.hasAttachments,
  importance: message.importance,
  isRead: message.isRead,
  bodyPreview: message.bodyPreview,
}
const messageDetail = {
  ...messageSummary,
  to: [],
  cc: [],
  body: message.body,
  untrustedContent: true,
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
function paginationReason(reason: string) {
  return `Continuation unavailable (${reason}). Returned messages are a bounded page; no usable next cursor.`
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
      'PAGINATION_ORIGIN',
    ],
    [
      'HTTP scheme',
      changeLink((url) => {
        url.protocol = 'http:'
      }),
      'PAGINATION_ORIGIN',
    ],
    [
      'URL username',
      changeLink((url) => {
        url.username = CANARY
      }),
      'PAGINATION_ORIGIN',
    ],
    [
      'URL password',
      changeLink((url) => {
        url.password = CANARY
      }),
      'PAGINATION_ORIGIN',
    ],
    [
      'non-default port',
      changeLink((url) => {
        url.port = '444'
      }),
      'PAGINATION_ORIGIN',
    ],
    [
      'fragment',
      changeLink((url) => {
        url.hash = CANARY
      }),
      'PAGINATION_ORIGIN',
    ],
    ['wrong route', nextLink({ path: '/v1.0/me/messages' }), 'PAGINATION_PATH_ME_MESSAGES'],
    ['wrong top', nextLink({ top: 4 }), 'PAGINATION_TOP'],
    [
      'unknown query',
      changeLink((url) => {
        url.searchParams.set('$search', CANARY)
      }),
      'PAGINATION_QUERY_KEYS',
    ],
    [
      'duplicate skip',
      changeLink((url) => {
        url.searchParams.append('$skip', '74')
      }),
      'PAGINATION_QUERY_KEYS',
    ],
    [
      'missing select',
      changeLink((url) => {
        url.searchParams.delete('$select')
      }),
      'PAGINATION_SELECT',
    ],
    [
      'changed select',
      changeLink((url) => {
        url.searchParams.set('$select', 'id,body')
      }),
      'PAGINATION_SELECT',
    ],
    [
      'changed ordering',
      changeLink((url) => {
        url.searchParams.set('$orderby', 'sentDateTime asc')
      }),
      'PAGINATION_ORDER',
    ],
    [
      'unexpected filter',
      changeLink((url) => {
        url.searchParams.set('$filter', 'isRead eq true')
      }),
      'PAGINATION_FILTER',
    ],
    ['non-progressing skip', nextLink({ skip: '0' }), 'PAGINATION_SKIP'],
    ['skip beyond cap', nextLink({ skip: '10001' }), 'PAGINATION_SKIP'],
    ['fractional skip', nextLink({ skip: '2.5' }), 'PAGINATION_SKIP'],
    ['skip token', nextLink({ skip: CANARY }), 'PAGINATION_SKIP'],
    ['relative URL', '/v1.0/me/messages?$skip=73', 'PAGINATION_URL'],
    ['malformed URL', `https://[${CANARY}`, 'PAGINATION_URL'],
    [
      'extra path segment',
      nextLink({ path: '/v1.0/me/mailFolders/inbox/messages/extra' }),
      'PAGINATION_PATH_OTHER',
    ],
    [
      'different API version',
      nextLink({ path: '/beta/me/mailFolders/inbox/messages' }),
      'PAGINATION_PATH_OTHER',
    ],
    [
      'case-changed query key',
      changeLink((url) => {
        url.searchParams.set('$SKIP', '73')
      }),
      'PAGINATION_QUERY_KEYS',
    ],
    [
      'case-changed select value',
      changeLink((url) => {
        url.searchParams.set('$select', SELECT.toLowerCase())
      }),
      'PAGINATION_SELECT',
    ],
    [
      'case-changed ordering value',
      changeLink((url) => {
        url.searchParams.set('$orderby', 'receiveddatetime desc')
      }),
      'PAGINATION_ORDER',
    ],
    [
      'missing top',
      changeLink((url) => {
        url.searchParams.delete('$top')
      }),
      'PAGINATION_TOP',
    ],
    [
      'missing skip',
      changeLink((url) => {
        url.searchParams.delete('$skip')
      }),
      'PAGINATION_SKIP',
    ],
    ['negative skip', nextLink({ skip: '-1' }), 'PAGINATION_SKIP'],
    ['unsafe integer skip', nextLink({ skip: '9007199254740993' }), 'PAGINATION_SKIP'],
  ])('preserves the valid current page for a nextLink with %s', async (_label, link, reason) => {
    const logs = vi.spyOn(console, 'log').mockImplementation(() => {})
    const { mcp, records, fetchFn } = await flow(() => ({
      value: [{ ...message, upstreamOnly: CANARY }],
      '@odata.nextLink': link,
    }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3 })
    expect(result.isError).not.toBe(true)
    expect(result.structuredContent).toEqual({
      messages: [messageSummary],
      hasMore: true,
      nextCursor: null,
      incompleteReason: paginationReason(reason),
    })
    expect(result.content).toEqual([
      { type: 'text', text: JSON.stringify(result.structuredContent) },
    ])
    expect(records[0]?.args).toEqual({ top: 3, skip: 0, mailbox: 'inbox' })
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(fetchFn.mock.calls[0]?.[0]).toBe(`https://example.test/${CANARY}`)
    expect(JSON.stringify(result)).not.toContain(link)
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
    expect(JSON.stringify(result)).not.toContain('evil.example.test')
    expect(JSON.stringify(logs.mock.calls)).not.toContain(CANARY)
    expect(JSON.stringify(logs.mock.calls)).not.toContain(link)
  })

  it('preserves the next page when its nextLink replays the current cursor skip', async () => {
    const { mcp, records, fetchFn } = await flow((_request, index) => ({
      value: [{ ...message, id: `message-${index + 1}` }],
      '@odata.nextLink': nextLink(),
    }))
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(records[1]?.args.skip).toBe(73)
    expect(result.isError).not.toBe(true)
    expect(result.structuredContent).toEqual({
      messages: [{ ...messageSummary, id: 'message-2' }],
      hasMore: true,
      nextCursor: null,
      incompleteReason: paginationReason('PAGINATION_SKIP'),
    })
    expect(fetchFn).toHaveBeenCalledTimes(2)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
  })

  it('accepts the maximum continuation skip and an empty final page', async () => {
    const { mcp, records, fetchFn } = await flow((_request, index) =>
      index === 0
        ? { value: [message], '@odata.nextLink': nextLink({ skip: '10000' }) }
        : { value: [], '@odata.nextLink': '' },
    )
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    expect(decodeCursor(cursor)).toMatchObject({ skip: 10000 })
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(result.isError).not.toBe(true)
    expect(records[1]?.args.skip).toBe(10000)
    expect(result.structuredContent).toEqual({
      messages: [],
      hasMore: false,
      nextCursor: null,
      incompleteReason: null,
    })
    expect(fetchFn).toHaveBeenCalledTimes(2)
  })

  it('preserves the page at the skip cap without issuing a cursor beyond the cap', async () => {
    const { mcp, records, fetchFn } = await flow((_request, index) => ({
      value: [message],
      '@odata.nextLink': nextLink({ skip: index === 0 ? '10000' : '10003' }),
    }))
    const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', { limit: 3 }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, cursor })
    expect(result.isError).not.toBe(true)
    expect(records[1]?.args.skip).toBe(10000)
    expect(result.structuredContent).toEqual({
      messages: [messageSummary],
      hasMore: true,
      nextCursor: null,
      incompleteReason: paginationReason('PAGINATION_SKIP'),
    })
    expect(fetchFn).toHaveBeenCalledTimes(2)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
  })

  it.each([
    ['outlook_list_messages', { limit: 3 }, nextLink()],
    [
      'outlook_get_conversation',
      { conversationId: 'conversation-1', limit: 3 },
      nextLink({ conversation: true, filter: "conversationId eq 'conversation-1'" }),
    ],
  ] as const)(
    'returns bounded empty-page metadata for %s claiming continuation',
    async (name, args, link) => {
      const { mcp, fetchFn } = await flow(() => ({ value: [], '@odata.nextLink': link }))
      const result = await call(mcp, name, args)
      expect(result.isError).not.toBe(true)
      expect(result.structuredContent).toEqual({
        ...(name === 'outlook_get_conversation' ? { conversationId: 'conversation-1' } : {}),
        messages: [],
        hasMore: true,
        nextCursor: null,
        incompleteReason: paginationReason('PAGINATION_EMPTY_PAGE'),
      })
      expect(fetchFn).toHaveBeenCalledOnce()
      expect(JSON.stringify(result)).not.toContain(link)
    },
  )

  it.each([
    ['lowercase mailfolders', '/v1.0/me/mailfolders/inbox/messages', 'inbox'],
    ['mixed-case known inbox segments', '/V1.0/ME/MailFolders/Inbox/Messages', 'inbox'],
    ['mixed-case known sent segments', '/V1.0/ME/MAILFOLDERS/SentItems/MESSAGES', 'sent'],
    ['mixed-case mailbox-wide segments', '/V1.0/ME/MESSAGES', 'all'],
  ] as const)(
    'accepts %s without changing the continuation scope',
    async (_label, path, mailbox) => {
      const { mcp, records, fetchFn } = await flow((_request, index) => ({
        value: [message],
        ...(index === 0 ? { '@odata.nextLink': nextLink({ path }) } : {}),
      }))
      const args = { limit: 3, mailbox }
      const first = await call(mcp, 'outlook_list_messages', args)
      expect(first.isError).not.toBe(true)
      expect(first.structuredContent).toMatchObject({
        messages: [messageSummary],
        hasMore: true,
        incompleteReason: null,
      })
      const cursor = cursorFrom(first)
      expect(decodeCursor(cursor)).toMatchObject({ operation: 'list_messages', skip: 73, limit: 3 })
      const second = await call(mcp, 'outlook_list_messages', { ...args, cursor })
      expect(second.isError).not.toBe(true)
      expect(records[1]?.args).toEqual({ top: 3, skip: 73, mailbox })
      expect(fetchFn).toHaveBeenCalledTimes(2)
    },
  )

  it.each([
    ['FolderCaseSensitive', 'foldercasesensitive'],
    ['INBOX', 'inbox'],
    ['SentItems', 'sentitems'],
  ])('never casefolds an explicit folder ID %s', async (folderId, changedFolderId) => {
    const link = nextLink({ path: `/v1.0/me/mailfolders/${changedFolderId}/messages` })
    const { mcp, records, fetchFn } = await flow(() => ({
      value: [message],
      '@odata.nextLink': link,
    }))
    const result = await call(mcp, 'outlook_list_messages', { limit: 3, folderId })
    expect(result.isError).not.toBe(true)
    expect(result.structuredContent).toEqual({
      messages: [messageSummary],
      hasMore: true,
      nextCursor: null,
      incompleteReason: paginationReason('PAGINATION_PATH_ME_FOLDER_SEGMENT'),
    })
    expect(records[0]?.args).toEqual({ top: 3, skip: 0, mailbox: 'inbox', folderId })
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(JSON.stringify(result)).not.toContain(link)
  })

  it.each([
    ["/v1.0/me/mailfolders('SYNTHETIC_PRIVATE_CANARY')/messages", 'ME_FOLDER_ODATA'],
    ['/v1.0/me/mailfolders(%27SYNTHETIC_PRIVATE_CANARY%27)/messages', 'ME_FOLDER_ODATA'],
    ['/v1.0/me/mailfolders/SYNTHETIC_PRIVATE_CANARY/messages', 'ME_FOLDER_SEGMENT'],
    ['/v1.0/users/SYNTHETIC_PRIVATE_CANARY/messages', 'USER_SEGMENT_MESSAGES'],
    [
      '/v1.0/users/SYNTHETIC_PRIVATE_CANARY/mailfolders/inbox/messages',
      'USER_SEGMENT_FOLDER_SEGMENT',
    ],
    [
      "/v1.0/users/SYNTHETIC_PRIVATE_CANARY/mailfolders('inbox')/messages",
      'USER_SEGMENT_FOLDER_ODATA',
    ],
    ["/v1.0/users('SYNTHETIC_PRIVATE_CANARY')/messages", 'USER_ODATA_MESSAGES'],
    [
      "/v1.0/users('SYNTHETIC_PRIVATE_CANARY')/mailfolders/inbox/messages",
      'USER_ODATA_FOLDER_SEGMENT',
    ],
    [
      "/v1.0/users('SYNTHETIC_PRIVATE_CANARY')/mailfolders('inbox')/messages",
      'USER_ODATA_FOLDER_ODATA',
    ],
    ["/v1.0/me/mailfolders('SYNTHETIC_PRIVATE_CANARY')/messages/delta", 'OTHER'],
    ['/beta/users/SYNTHETIC_PRIVATE_CANARY/messages', 'OTHER'],
    ['/v1.0/users/SYNTHETIC_PRIVATE_CANARY/contacts', 'OTHER'],
  ])(
    'classifies rejected shape %s without exposing or accepting its identity',
    async (path, shape) => {
      const link = nextLink({ path })
      const { mcp, fetchFn } = await flow(() => ({ value: [message], '@odata.nextLink': link }))
      const result = await call(mcp, 'outlook_list_messages', { limit: 3 })
      expect(result.isError).not.toBe(true)
      expect(result.structuredContent).toEqual({
        messages: [messageSummary],
        hasMore: true,
        nextCursor: null,
        incompleteReason: paginationReason(`PAGINATION_PATH_${shape}`),
      })
      expect(fetchFn).toHaveBeenCalledOnce()
      expect(JSON.stringify(result)).not.toContain(CANARY)
      expect(JSON.stringify(result)).not.toContain(link)
    },
  )

  it.each(['canonical', 'mixed-case'] as const)(
    'validates filtered folder pagination with the %s escaped route and no ordering',
    async (routeCase) => {
      const folderId = "資料/A'B+="
      const filters = { isRead: false, hasAttachments: true }
      const link = nextLink({
        path:
          routeCase === 'canonical'
            ? `/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages`
            : `/V1.0/ME/mailfolders/${encodeURIComponent(folderId)}/Messages`,
        filter: 'isRead eq false and hasAttachments eq true',
      })
      const { mcp, records, fetchFn } = await flow((_request, index) => ({
        value: [message],
        ...(index === 0 ? { '@odata.nextLink': link } : {}),
      }))
      const args = { limit: 3, mailbox: 'all', folderId, filters }
      const cursor = cursorFrom(await call(mcp, 'outlook_list_messages', args))
      expect((await call(mcp, 'outlook_list_messages', { ...args, cursor })).isError).not.toBe(true)
      expect(records[1]?.args).toEqual({ top: 3, skip: 73, mailbox: 'all', folderId, filters })
      expect(fetchFn).toHaveBeenCalledTimes(2)
    },
  )

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

  it.each([
    ['invalid origin', `https://evil.example.test/${CANARY}`, 'PAGINATION_ORIGIN'],
    [
      'wrong route',
      nextLink({
        conversation: true,
        path: '/v1.0/me/mailFolders/inbox/messages',
        filter: "conversationId eq 'conversation-1'",
      }),
      'PAGINATION_PATH_ME_FOLDER_SEGMENT',
    ],
    [
      'case-changed filter value',
      nextLink({ conversation: true, filter: "conversationId eq 'CONVERSATION-1'" }),
      'PAGINATION_FILTER',
    ],
    [
      'changed filter property',
      nextLink({ conversation: true, filter: "conversationid eq 'conversation-1'" }),
      'PAGINATION_FILTER',
    ],
    [
      'skip beyond cap',
      nextLink({ conversation: true, filter: "conversationId eq 'conversation-1'", skip: '10001' }),
      'PAGINATION_SKIP',
    ],
  ])('preserves normalized conversation messages for %s', async (_label, link, reason) => {
    const { mcp, records, fetchFn } = await flow(() => ({
      value: [
        { ...message, id: 'older', sentDateTime: '2026-01-01T11:00:00Z', upstreamOnly: CANARY },
        { ...message, id: 'newer', upstreamOnly: CANARY },
      ],
      '@odata.nextLink': link,
    }))
    const result = await call(mcp, 'outlook_get_conversation', {
      conversationId: 'conversation-1',
      limit: 3,
    })
    expect(result.isError).not.toBe(true)
    expect(result.structuredContent).toEqual({
      conversationId: 'conversation-1',
      messages: [
        { ...messageDetail, id: 'newer' },
        { ...messageDetail, id: 'older', sentDateTime: '2026-01-01T11:00:00Z' },
      ],
      hasMore: true,
      nextCursor: null,
      incompleteReason: paginationReason(reason),
    })
    expect(result.content).toEqual([
      { type: 'text', text: JSON.stringify(result.structuredContent) },
    ])
    expect(records[0]?.args).toEqual({ conversationId: 'conversation-1', top: 3, skip: 0 })
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(JSON.stringify(result)).not.toContain(link)
    expect(JSON.stringify(result)).not.toContain(CANARY)
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
    expect(JSON.stringify(result)).not.toContain('evil.example.test')
  })

  it('accepts mixed-case conversation endpoint segments without changing the filter value', async () => {
    const conversationId = 'Conversation-CaseSensitive'
    const link = nextLink({
      conversation: true,
      path: '/V1.0/ME/Messages',
      filter: `conversationId eq '${conversationId}'`,
    })
    const { mcp, records, fetchFn } = await flow((_request, index) => ({
      value: [{ ...message, conversationId }],
      ...(index === 0 ? { '@odata.nextLink': link } : {}),
    }))
    const args = { conversationId, limit: 3 }
    const first = await call(mcp, 'outlook_get_conversation', args)
    const cursor = cursorFrom(first)
    expect(decodeCursor(cursor)).toMatchObject({
      operation: 'get_conversation',
      skip: 73,
      limit: 3,
    })
    expect(first.structuredContent).toMatchObject({
      conversationId,
      hasMore: true,
      incompleteReason:
        'Conversation result is a bounded page sorted only within that page, not the complete globally ordered thread.',
    })
    const second = await call(mcp, 'outlook_get_conversation', { ...args, cursor })
    expect(second.isError).not.toBe(true)
    expect(records[1]?.args).toEqual({ conversationId, top: 3, skip: 73 })
    expect(fetchFn).toHaveBeenCalledTimes(2)
  })

  it.each([
    [
      'malformed message',
      'outlook_list_messages',
      { limit: 3 },
      [{ ...message, receivedDateTime: CANARY }],
    ],
    [
      'oversized page',
      'outlook_list_messages',
      { limit: 1 },
      [message, { ...message, id: 'message-2' }],
    ],
    [
      'mismatched conversation',
      'outlook_get_conversation',
      { conversationId: 'conversation-1' },
      [{ ...message, conversationId: 'different-conversation' }],
    ],
    [
      'duplicate conversation IDs',
      'outlook_get_conversation',
      { conversationId: 'conversation-1' },
      [message, message],
    ],
    [
      'malformed conversation body',
      'outlook_get_conversation',
      { conversationId: 'conversation-1' },
      [{ ...message, body: { contentType: 'unsupported', content: CANARY } }],
    ],
  ] as const)(
    'still rejects %s when the nextLink is also hostile',
    async (_label, name, args, value) => {
      const { mcp, fetchFn } = await flow(() => ({
        value,
        '@odata.nextLink': `https://evil.example.test/${CANARY}`,
      }))
      const result = await call(mcp, name, args)
      expect(result.isError).toBe(true)
      expect(result.structuredContent).toBeUndefined()
      expect(fetchFn).toHaveBeenCalledOnce()
      expect(JSON.stringify(result)).not.toContain(CANARY)
      expect(JSON.stringify(result)).not.toContain('evil.example.test')
    },
  )

  it.each([
    [
      'outlook_list_messages',
      { limit: 3 },
      { ...message, bodyPreview: '日'.repeat(45000) + CANARY },
    ],
    [
      'outlook_get_conversation',
      { conversationId: 'conversation-1', limit: 3 },
      { ...message, body: { contentType: 'text', content: '日'.repeat(45000) + CANARY } },
    ],
  ] as const)(
    'still enforces serialized output bounds for %s with rejected continuation',
    async (name, args, oversizedMessage) => {
      const { mcp, fetchFn } = await flow(() => ({
        value: [oversizedMessage],
        '@odata.nextLink': `https://evil.example.test/${CANARY}`,
      }))
      const result = await call(mcp, name, args)
      expect(result.isError).toBe(true)
      expect(result.structuredContent).toBeUndefined()
      expect(JSON.stringify(result)).toContain('output limit')
      expect(JSON.stringify(result)).not.toContain(CANARY)
      expect(JSON.stringify(result)).not.toContain('evil.example.test')
      expect(fetchFn).toHaveBeenCalledOnce()
    },
  )

  it.each([
    ['outlook_list_messages', { limit: 3 }],
    ['outlook_get_conversation', { conversationId: 'conversation-1', limit: 3 }],
  ] as const)(
    'still rejects invalid caller cursors for %s without an upstream call',
    async (name, args) => {
      const { mcp, fetchFn } = await flow(() => ({
        value: [message],
        '@odata.nextLink': `https://evil.example.test/${CANARY}`,
      }))
      const result = await call(mcp, name, { ...args, cursor: encodeCursor({ url: CANARY }) })
      expect(result.isError).toBe(true)
      expect(result.content).toEqual([{ type: 'text', text: 'Invalid continuation cursor' }])
      expect(result.structuredContent).toBeUndefined()
      expect(fetchFn).not.toHaveBeenCalled()
      expect(JSON.stringify(result)).not.toContain(CANARY)
    },
  )

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
