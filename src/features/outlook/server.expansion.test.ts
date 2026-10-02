import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { createOutlookMcpServer, TOOL_NAMES } from './server.js'

const graphSummaryItem = {
  id: 'msg-1',
  subject: 'Weekly report',
  from: { emailAddress: { name: 'Jane Doe', address: 'jane@example.com' } },
  receivedDateTime: '2025-01-01T09:00:00Z',
  sentDateTime: '2025-01-01T08:59:00Z',
  parentFolderId: 'inbox-id',
  conversationId: 'conv-1',
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  bodyPreview: 'Hello world...',
} as const
const graphDetailItem = {
  ...graphSummaryItem,
  toRecipients: [{ emailAddress: { name: 'John Smith', address: 'john@example.com' } }],
  ccRecipients: [],
  body: { contentType: 'text', content: 'Hello world' },
} as const
const sentRequestSchema = z.object({
  operation: z.string(),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})
function mockFlow(respond: (operation: string) => unknown) {
  const records: unknown[] = []
  const fetchFn: typeof fetch = async (_input, init) => {
    const body = typeof init?.body === 'string' ? JSON.parse(init.body) : undefined
    records.push(body)
    const { operation, requestId } = sentRequestSchema.parse(body)
    return Response.json({ ok: true, requestId, operation, data: respond(operation) })
  }
  return { records, fetchFn }
}
async function connectClient(client: PowerAutomateClient): Promise<Client> {
  const mcpClient = new Client({ name: 'test', version: '0.0.0' })
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  await Promise.all([
    mcpClient.connect(clientTransport),
    createOutlookMcpServer(client).connect(serverTransport),
  ])
  return mcpClient
}

describe('createOutlookMcpServer', () => {
  it('registers the fixed Outlook and OneDrive read tools', async () => {
    const { fetchFn } = mockFlow(() => ({ value: [] }))
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test', gatewayKey: 'test', fetchFn }),
    )
    const { tools } = await mcpClient.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual([...TOOL_NAMES].sort())
  })
  it('preserves default inbox list behavior and adds cursor metadata', async () => {
    const { records, fetchFn } = mockFlow(() => ({ value: [graphSummaryItem] }))
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_list_messages',
      arguments: { limit: 7 },
    })
    expect(sentRequestSchema.parse(records[0]).operation).toBe('list_messages')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ top: 7, skip: 0, mailbox: 'inbox' })
    expect(result.structuredContent).toMatchObject({
      messages: [
        {
          id: 'msg-1',
          conversationId: 'conv-1',
          parentFolderId: 'inbox-id',
          sentDateTime: '2025-01-01T08:59:00Z',
        },
      ],
      hasMore: false,
      nextCursor: null,
    })
  })
  it('does not invent search continuation', async () => {
    const { records, fetchFn } = mockFlow(() => ({
      value: [graphSummaryItem],
      '@odata.nextLink': 'https://graph.invalid/next',
    }))
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_search_messages',
      arguments: { query: 'PMDA', mailbox: 'all' },
    })
    expect(sentRequestSchema.parse(records[0]).args).toEqual({
      query: 'PMDA',
      top: 10,
      mailbox: 'all',
    })
    expect(result.structuredContent).toMatchObject({ hasMore: true, nextCursor: null })
    expect(JSON.stringify(result)).not.toContain('graph.invalid')
  })
  it('reads a mailbox-wide conversation page', async () => {
    const { records, fetchFn } = mockFlow((operation) =>
      operation === 'get_conversation' ? { value: [graphDetailItem] } : graphDetailItem,
    )
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_get_conversation',
      arguments: { conversationId: 'conv-1', limit: 5 },
    })
    expect(sentRequestSchema.parse(records[0]).operation).toBe('get_conversation')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({
      conversationId: 'conv-1',
      top: 5,
      skip: 0,
    })
    expect(result.structuredContent).toMatchObject({
      conversationId: 'conv-1',
      messages: [{ id: 'msg-1', body: { content: 'Hello world' } }],
      hasMore: false,
    })
  })
  it('surfaces Power Automate errors without leaking the URL', async () => {
    const fetchFn: typeof fetch = async () => new Response('nope', { status: 500 })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test',
        fetchFn,
      }),
    )
    const result = await mcpClient.request({
      method: 'tools/call',
      params: { name: 'outlook_get_message', arguments: { messageId: 'msg-1' } },
    })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('get_message failed with HTTP status 500')
    expect(JSON.stringify(result)).not.toContain('example.test')
  })
})
