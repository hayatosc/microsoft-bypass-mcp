import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { CallToolResultSchema } from '@modelcontextprotocol/sdk/types.js'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { createOutlookMcpServer } from './server.js'

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
  operation: z.enum(['list_messages', 'search_messages', 'get_message']),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})

function mockFlow(respond: () => Response): { records: unknown[]; fetchFn: typeof fetch } {
  const records: unknown[] = []
  const fetchFn: typeof fetch = async (input, init) => {
    records.push(typeof init?.body === 'string' ? JSON.parse(init.body) : undefined)
    return respond()
  }
  return { records, fetchFn }
}

async function connectClient(client: PowerAutomateClient): Promise<Client> {
  const mcpClient = new Client({ name: 'test', version: '0.0.0' })
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const server = createOutlookMcpServer(client)
  await Promise.all([mcpClient.connect(clientTransport), server.connect(serverTransport)])
  return mcpClient
}

describe('createOutlookMcpServer', () => {
  it('registers exactly the three tools', async () => {
    const { fetchFn } = mockFlow(
      () => new Response(JSON.stringify({ ok: true, data: { value: [] } }), { status: 200 }),
    )
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test', fetchFn }),
    )

    const { tools } = await mcpClient.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual([
      'outlook_get_message',
      'outlook_list_messages',
      'outlook_search_messages',
    ])
  })

  it('outlook_list_messages translates limit to top and returns summaries', async () => {
    const { records, fetchFn } = mockFlow(
      () =>
        new Response(JSON.stringify({ ok: true, data: { value: [graphSummaryItem] } }), {
          status: 200,
        }),
    )
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_list_messages',
      arguments: { limit: 7 },
    })

    expect(sentRequestSchema.parse(records[0]).operation).toBe('list_messages')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ top: 7 })
    expect(result.structuredContent).toEqual({ messages: [summary] })
  })

  it('outlook_search_messages sends the query and a default top', async () => {
    const { records, fetchFn } = mockFlow(
      () =>
        new Response(JSON.stringify({ ok: true, data: { value: [graphSummaryItem] } }), {
          status: 200,
        }),
    )
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_search_messages',
      arguments: { query: 'PMDA' },
    })

    expect(sentRequestSchema.parse(records[0]).operation).toBe('search_messages')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ query: 'PMDA', top: 10 })
    expect(result.structuredContent).toEqual({ messages: [summary] })
  })

  it('outlook_get_message sends the messageId and returns the full detail', async () => {
    const { records, fetchFn } = mockFlow(
      () => new Response(JSON.stringify({ ok: true, data: graphDetailItem }), { status: 200 }),
    )
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_get_message',
      arguments: { messageId: 'msg-1' },
    })

    expect(sentRequestSchema.parse(records[0]).operation).toBe('get_message')
    expect(sentRequestSchema.parse(records[0]).args).toEqual({ messageId: 'msg-1' })
    expect(result.structuredContent).toEqual({
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
    })
  })

  it('surfaces Power Automate errors without leaking the URL', async () => {
    const { fetchFn } = mockFlow(() => new Response('nope', { status: 500 }))
    const mcpClient = await connectClient(
      new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn }),
    )

    const result = await mcpClient.request(
      {
        method: 'tools/call',
        params: { name: 'outlook_get_message', arguments: { messageId: 'msg-1' } },
      },
      CallToolResultSchema,
    )

    expect(result.isError).toBe(true)
    const text = result.content.map((block) => (block.type === 'text' ? block.text : '')).join('')
    expect(text).toContain('get_message failed with HTTP status 500')
    expect(text).not.toContain('example.test')
  })
})
