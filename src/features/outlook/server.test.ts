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
  hasAttachments: true,
  importance: 'high',
  isRead: false,
  bodyPreview: 'Hello world...',
} as const
const graphDetailItem = {
  ...graphSummaryItem,
  toRecipients: [{ emailAddress: { name: 'John Smith', address: 'john@example.com' } }],
  ccRecipients: [{ emailAddress: { name: 'Admin', address: 'admin@example.com' } }],
  body: { contentType: 'text', content: 'Hello world' },
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
const sentRequestSchema = z.object({
  operation: z.string(),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})
function mockFlow(data: unknown): { records: unknown[]; fetchFn: typeof fetch } {
  const records: unknown[] = []
  const fetchFn: typeof fetch = async (_input, init) => {
    const body = typeof init?.body === 'string' ? JSON.parse(init.body) : undefined
    records.push(body)
    const { operation, requestId } = sentRequestSchema.parse(body)
    return Response.json({ ok: true, requestId, operation, data })
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
  it('registers the fixed tools with expected annotations', async () => {
    const { fetchFn } = mockFlow({ value: [] })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const { tools } = await mcpClient.listTools()
    expect(tools.map((tool) => tool.name).sort()).toEqual([...TOOL_NAMES].sort())
    const draftNames = new Set([
      'outlook_create_draft',
      'outlook_create_reply_draft',
      'outlook_add_draft_attachment',
    ])
    for (const tool of tools) {
      if (draftNames.has(tool.name)) {
        expect(tool.annotations).toMatchObject({
          readOnlyHint: false,
          destructiveHint: false,
          idempotentHint: false,
          openWorldHint: true,
        })
      } else {
        expect(tool.annotations?.readOnlyHint).toBe(true)
      }
    }
  })
  it('outlook_list_messages translates limit to top and returns summaries', async () => {
    const { records, fetchFn } = mockFlow({ value: [graphSummaryItem] })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_list_messages',
      arguments: { limit: 7 },
    })
    expect(sentRequestSchema.parse(records[0]).operation).toBe('list_messages')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ top: 7, skip: 0, mailbox: 'inbox' })
    expect(result.structuredContent).toEqual({
      messages: [summary],
      hasMore: false,
      nextCursor: null,
      incompleteReason: null,
    })
  })
  it('outlook_search_messages sends the query and a default top', async () => {
    const { records, fetchFn } = mockFlow({ value: [graphSummaryItem] })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_search_messages',
      arguments: { query: 'PMDA' },
    })
    expect(sentRequestSchema.parse(records[0]).operation).toBe('search_messages')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({
      query: 'PMDA',
      top: 10,
      mailbox: 'inbox',
    })
    expect(result.structuredContent).toEqual({
      messages: [summary],
      hasMore: false,
      nextCursor: null,
      incompleteReason: null,
    })
  })
  it('outlook_get_message sends the messageId and returns the full detail', async () => {
    const { records, fetchFn } = mockFlow(graphDetailItem)
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const result = await mcpClient.callTool({
      name: 'outlook_get_message',
      arguments: { messageId: 'msg-1' },
    })
    expect(sentRequestSchema.parse(records[0]).operation).toBe('get_message')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ messageId: 'msg-1' })
    expect(result.structuredContent).toEqual({
      ...summary,
      to: [{ name: 'John Smith', address: 'john@example.com' }],
      cc: [{ name: 'Admin', address: 'admin@example.com' }],
      body: { contentType: 'text', content: 'Hello world' },
      untrustedContent: true,
    })
  })
  it('surfaces Power Automate errors without leaking the URL', async () => {
    const fetchFn: typeof fetch = async () => new Response('nope', { status: 500 })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const result = await mcpClient.request({
      method: 'tools/call',
      params: { name: 'outlook_get_message', arguments: { messageId: 'msg-1' } },
    })
    expect(result.isError).toBe(true)
    const text = result.content.map((block) => (block.type === 'text' ? block.text : '')).join('')
    expect(text).toContain('get_message failed with HTTP status 500')
    expect(text).not.toContain('example.test')
  })
})
