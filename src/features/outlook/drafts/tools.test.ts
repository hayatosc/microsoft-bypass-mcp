import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../../lib/power-automate.js'
import { syntheticGraphSizeAttachment } from '../attachments/synthetic-fixtures.js'
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
    const attachmentTool = draftTools.find((tool) => tool.name === 'outlook_add_draft_attachment')
    expect(attachmentTool?.description).toContain('nonempty attachment of at most 2 MiB')
    expect(attachmentTool?.description).toContain('raw-file size, not Graph')
    const output = z
      .object({ properties: z.object({ size: z.object({ description: z.string() }) }) })
      .parse(attachmentTool?.outputSchema)
    expect(output.properties.size.description).toBe(
      'Verified raw-file size in bytes, not Microsoft Graph attachment metadata size.',
    )
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

  it('returns raw size 889 for the overhead DOCX and keeps bytes out of results and logs', async () => {
    const file = syntheticGraphSizeAttachment()
    expect(atob(file.contentBytes)).toHaveLength(889)
    expect(file.size).toBe(1223)
    // The generated flow verifies Graph's returned bytes and projects raw size.
    const { records, fetchFn } = mockFlow({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: file.name,
      size: 889,
      contentBytes: file.contentBytes,
      graphUrl: 'https://graph.microsoft.com/synthetic-redaction-canary',
    })
    const mcpClient = await connectClient(
      new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      }),
    )
    const logs = vi.spyOn(console, 'log').mockImplementation(() => {})
    try {
      const result = await mcpClient.callTool({
        name: 'outlook_add_draft_attachment',
        arguments: {
          draftId: 'draft-1',
          name: file.name,
          contentType: file.contentType,
          contentBytes: file.contentBytes,
        },
      })
      expect(result.isError).not.toBe(true)
      expect(result.structuredContent).toEqual({
        draftId: 'draft-1',
        attachmentId: 'attachment-1',
        name: file.name,
        size: 889,
      })
      expect(records).toHaveLength(1)
      expect(logs).toHaveBeenCalledOnce()
      const telemetry = z.string().parse(logs.mock.calls[0]?.[0])
      expect(Object.keys(JSON.parse(telemetry)).sort()).toEqual([
        'durationMs',
        'operation',
        'requestId',
        'status',
        'success',
        'type',
      ])
      for (const serialized of [JSON.stringify(result), telemetry]) {
        expect(serialized).not.toContain(file.contentBytes)
        expect(serialized).not.toContain('contentBytes')
        expect(serialized).not.toContain('synthetic-redaction-canary')
        expect(serialized).not.toContain('example.test')
      }
      expect(telemetry).not.toContain(file.name)
      expect(telemetry).not.toContain('draft-1')
      expect(telemetry).not.toContain('attachment-1')
    } finally {
      logs.mockRestore()
    }
  })

  it('rejects raw max+1 before any flow request', async () => {
    const { records, fetchFn } = mockFlow({})
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
        contentBytes: btoa('x'.repeat(2097153)),
      },
    })
    expect(result.isError).toBe(true)
    expect(records).toHaveLength(0)
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
