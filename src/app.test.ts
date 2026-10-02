import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import app from './app.js'
import { syntheticAttachment } from './features/outlook/attachments/synthetic-fixtures.js'
import { getAccessConfig, getPowerAutomateGatewayKey, getPowerAutomateUrl } from './lib/env.js'

const env = {
  POWER_AUTOMATE_URL: 'https://example.test/flow',
  POWER_AUTOMATE_GATEWAY_KEY: 'test-gateway-key',
}

describe('mcp endpoint', () => {
  it('handles an initialize request with no app-level auth', async () => {
    const res = await app.request(
      '/mcp',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json, text/event-stream',
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: {
            protocolVersion: '2025-06-18',
            capabilities: {},
            clientInfo: { name: 'test', version: '0.0.0' },
          },
        }),
      },
      env,
    )
    expect(res.status).toBe(200)
  })
})

describe('attachment tools through Hono', () => {
  afterEach(() => vi.restoreAllMocks())

  it.each(['pdf', 'docx', 'xlsx'] as const)(
    'lists, inspects, and reads a synthetic %s through /mcp',
    async (format) => {
      const file = syntheticAttachment(format)
      const fetchMock = vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
        expect(input).toBe(env.POWER_AUTOMATE_URL)
        expect(new Headers(init?.headers).get('X-MCP-Gateway-Key')).toBe(
          env.POWER_AUTOMATE_GATEWAY_KEY,
        )
        const request = z
          .object({
            operation: z.enum(['list_attachments', 'get_attachment']),
            requestId: z.uuid(),
          })
          .parse(JSON.parse(z.string().parse(init?.body)))
        const data = request.operation === 'list_attachments' ? { value: [file] } : file
        return Response.json({ ok: true, ...request, data })
      })
      async function rpc(method: string, params: unknown) {
        const response = await app.request(
          '/mcp',
          {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Accept: 'application/json, text/event-stream',
            },
            body: JSON.stringify({ jsonrpc: '2.0', id: 1, method, params }),
          },
          env,
        )
        expect(response.status).toBe(200)
        const text = await response.text()
        const payload = response.headers.get('content-type')?.includes('text/event-stream')
          ? text
              .split('\n')
              .find((line) => line.startsWith('data:'))
              ?.slice(5)
          : text
        const envelope = z
          .object({
            result: z
              .object({
                isError: z.boolean().optional(),
                tools: z.array(z.object({ name: z.string() })).optional(),
                structuredContent: z.record(z.string(), z.unknown()).optional(),
              })
              .passthrough(),
          })
          .parse(JSON.parse(payload ?? 'null'))
        expect(envelope.result.isError).not.toBe(true)
        expect(JSON.stringify(envelope.result)).not.toContain('contentBytes')
        return envelope.result
      }
      const tools = await rpc('tools/list', {})
      expect(tools.tools).toHaveLength(6)
      const listing = await rpc('tools/call', {
        name: 'outlook_list_attachments',
        arguments: { messageId: 'msg-1' },
      })
      expect(listing.structuredContent).toMatchObject({
        attachments: [{ supportedFormat: format }],
        hasMore: false,
      })
      const target = { messageId: 'msg-1', attachmentId: 'att-1' }
      const inspection = await rpc('tools/call', {
        name: 'outlook_inspect_attachment',
        arguments: target,
      })
      expect(inspection.structuredContent).toMatchObject({ structure: { format } })
      const selection =
        format === 'pdf'
          ? { format, pageStart: 1, pageEnd: 1 }
          : format === 'docx'
            ? { format, offset: 0, length: 100 }
            : { format, sheet: 'Budget', range: 'A1:B2' }
      const read = await rpc('tools/call', {
        name: 'outlook_read_attachment',
        arguments: { ...target, selection },
      })
      expect(read.structuredContent).toMatchObject({ data: { format }, untrustedContent: true })
      expect(JSON.stringify(read)).toContain('Synthetic')
      expect(fetchMock).toHaveBeenCalledTimes(3)
    },
  )
})

describe('public info', () => {
  it('GET / exposes the server name and tools', async () => {
    const res = await app.request('/')
    expect(res.status).toBe(200)
    const text = await res.text()
    expect(text).toContain('outlook_list_messages')
    expect(text).toContain('outlook_search_messages')
    expect(text).toContain('outlook_get_message')
  })
})

describe('env accessors', () => {
  it('fail fast on missing required values', () => {
    expect(() => getPowerAutomateUrl({})).toThrow('POWER_AUTOMATE_URL')
    expect(() => getPowerAutomateGatewayKey({})).toThrow('POWER_AUTOMATE_GATEWAY_KEY')
  })

  it('returns null when Access is not configured', () => {
    expect(getAccessConfig({})).toBeNull()
  })

  it('returns the config when both Access values are set', () => {
    expect(
      getAccessConfig({
        TEAM_DOMAIN: 'https://team.cloudflareaccess.com',
        POLICY_AUD: 'aud',
      }),
    ).toEqual({ domain: 'https://team.cloudflareaccess.com', aud: 'aud' })
  })

  it('fails fast on a partially configured Access pair', () => {
    expect(() => getAccessConfig({ TEAM_DOMAIN: 'https://team.cloudflareaccess.com' })).toThrow(
      'TEAM_DOMAIN',
    )
  })
})
