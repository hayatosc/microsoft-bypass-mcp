import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../../lib/power-automate.js'
import { createOutlookMcpServer } from '../server.js'

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

describe('draft MCP tools', () => {
  it('registers draft tools as non-read-only non-idempotent draft operations', async () => {
    const { fetchFn } = mockFlow({ id: 'draft-1', isDraft: true })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    const { tools } = await mcpClient.listTools()
    const draftTools = tools.filter(
      (tool) =>
        tool.name.startsWith('outlook_create_') || tool.name === 'outlook_add_draft_attachment',
    )
    expect(draftTools.map((tool) => tool.name).sort()).toEqual([
      'outlook_add_draft_attachment',
      'outlook_create_draft',
      'outlook_create_reply_draft',
    ])
    for (const tool of draftTools) {
      expect(tool.annotations).toMatchObject({
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      })
      expect(tool.description).toContain('approval')
      expect(tool.description).toContain('never sends')
      expect(tool.description).toContain('blindly retry')
    }
  })

  it('creates a draft with normalized typed params and sanitized output', async () => {
    const { records, fetchFn } = mockFlow({ id: 'draft-1', isDraft: true, body: 'secret' })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_create_draft',
      arguments: {
        to: ['to@example.edu'],
        cc: ['cc@example.edu'],
        subject: 'Subject',
        body: 'Sensitive body',
      },
    })

    expect(sentRequestSchema.parse(records[0])).toMatchObject({
      operation: 'create_draft',
      args: {
        to: ['to@example.edu'],
        cc: ['cc@example.edu'],
        subject: 'Subject',
        body: 'Sensitive body',
      },
    })
    expect(result.structuredContent).toEqual({ draftId: 'draft-1', isDraft: true })
    expect(JSON.stringify(result)).not.toContain('Sensitive body')
    expect(JSON.stringify(result)).not.toContain('secret')
  })

  it('creates a reply draft with only messageId and plain body', async () => {
    const { records, fetchFn } = mockFlow({ id: 'reply-draft-1', isDraft: true })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_create_reply_draft',
      arguments: { messageId: 'message-1', body: 'Reply body' },
    })

    expect(sentRequestSchema.parse(records[0])).toMatchObject({
      operation: 'create_reply_draft',
      args: { messageId: 'message-1', body: 'Reply body' },
    })
    expect(result.structuredContent).toEqual({ draftId: 'reply-draft-1', isDraft: true })
    expect(JSON.stringify(result)).not.toContain('Reply body')
  })

  it('adds a draft attachment and never returns raw bytes', async () => {
    const { records, fetchFn } = mockFlow({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
      contentBytes: 'aGVsbG8=',
      graphUrl: 'https://graph.microsoft.com/secret',
    })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    const result = await mcpClient.callTool({
      name: 'outlook_add_draft_attachment',
      arguments: {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      },
    })

    expect(sentRequestSchema.parse(records[0])).toMatchObject({
      operation: 'add_draft_attachment',
      args: {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      },
    })
    expect(result.structuredContent).toEqual({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
    })
    expect(JSON.stringify(result)).not.toContain('aGVsbG8=')
    expect(JSON.stringify(result)).not.toContain('graph.microsoft.com')
  })

  it('rejects arbitrary passthrough fields before calling the flow', async () => {
    const { records, fetchFn } = mockFlow({ id: 'draft-1', isDraft: true })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    await expect(
      mcpClient.callTool({
        name: 'outlook_create_draft',
        arguments: {
          to: ['to@example.edu'],
          subject: 'Subject',
          body: 'Body',
          graphUrl: 'https://graph.microsoft.com/v1.0/me/sendMail',
        },
      }),
    ).resolves.toMatchObject({ isError: true })
    expect(records).toHaveLength(0)
  })

  it('returns an ambiguous write warning on draft operation failures without leaking body', async () => {
    const fetchFn: typeof fetch = async () => new Response('failed with body text', { status: 500 })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )

    const result = await mcpClient.request({
      method: 'tools/call',
      params: {
        name: 'outlook_create_draft',
        arguments: { to: ['to@example.edu'], subject: 'Subject', body: 'Sensitive body' },
      },
    })
    expect(result.isError).toBe(true)
    const text = result.content.map((block) => (block.type === 'text' ? block.text : '')).join('')
    expect(text).toContain('write outcome is ambiguous')
    expect(text).toContain('Inspect Drafts before retrying')
    expect(text).not.toContain('Sensitive body')
    expect(text).not.toContain('example.test')
  })
})
