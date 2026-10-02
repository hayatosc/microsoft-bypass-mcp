import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import {
  PowerAutomateAmbiguousWriteError,
  PowerAutomateClient,
  PowerAutomateError,
} from './power-automate.js'

afterEach(() => {
  vi.unstubAllGlobals()
})

const requestSchema = z.object({
  operation: z.enum([
    'list_messages',
    'search_messages',
    'get_message',
    'create_draft',
    'create_reply_draft',
    'add_draft_attachment',
  ]),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})

type RequestRecord = { url: string; headers: Record<string, string>; body: unknown }

function mockFetch(respond: (record: RequestRecord) => Response): {
  records: RequestRecord[]
  fetchFn: typeof fetch
} {
  const records: RequestRecord[] = []
  const fetchFn: typeof fetch = async (input, init) => {
    const record = {
      url: typeof input === 'string' ? input : input instanceof URL ? input.href : input.url,
      headers: Object.fromEntries(new Headers(init?.headers).entries()),
      body: typeof init?.body === 'string' ? JSON.parse(init.body) : undefined,
    }
    records.push(record)
    return respond(record)
  }
  return { records, fetchFn }
}

function successResponse(record: RequestRecord, data: unknown): Response {
  const { operation, requestId } = requestSchema.parse(record.body)
  return new Response(JSON.stringify({ ok: true, requestId, operation, data }), { status: 200 })
}

describe('PowerAutomateClient', () => {
  it('sends the operation and args and includes a generated requestId', async () => {
    const cases = [
      { operation: 'list_messages' as const, args: { top: 5 } },
      { operation: 'search_messages' as const, args: { query: 'PMDA', top: 10 } },
      { operation: 'get_message' as const, args: { messageId: 'msg-1' } },
      {
        operation: 'create_draft' as const,
        args: { to: ['to@example.edu'], subject: 'Subject', body: 'Body' },
      },
      { operation: 'create_reply_draft' as const, args: { messageId: 'msg-1', body: 'Body' } },
      {
        operation: 'add_draft_attachment' as const,
        args: {
          draftId: 'draft-1',
          name: 'note.txt',
          contentType: 'text/plain',
          contentBytes: 'aGVsbG8=',
        },
      },
    ]
    const { records, fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    for (const { operation, args } of cases) {
      await client.call(operation, args)
    }

    expect(records).toHaveLength(cases.length)
    expect(records[0]?.url).toBe('https://example.test/flow')
    expect(records[0]?.headers['x-mcp-gateway-key']).toBe('test-gateway-key')
    const sent = cases.map((_, i) => requestSchema.parse(records[i]?.body))
    expect(sent.map((s) => s.operation)).toEqual([
      'list_messages',
      'search_messages',
      'get_message',
      'create_draft',
      'create_reply_draft',
      'add_draft_attachment',
    ])
    expect(sent.map((s) => s.args)).toEqual([
      { top: 5 },
      { query: 'PMDA', top: 10 },
      { messageId: 'msg-1' },
      { to: ['to@example.edu'], subject: 'Subject', body: 'Body' },
      { messageId: 'msg-1', body: 'Body' },
      {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      },
    ])
  })

  it('generates a fresh requestId per call', async () => {
    const { records, fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await client.call('list_messages', { top: 5 })
    await client.call('list_messages', { top: 5 })

    const requestIds = records.map((r) => requestSchema.parse(r.body).requestId)
    expect(new Set(requestIds).size).toBe(2)
  })

  it('parses the response JSON and unwraps the envelope', async () => {
    const data = { value: [{ id: 'msg-1' }] }
    const { fetchFn } = mockFetch((record) => successResponse(record, data))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await expect(client.call('get_message', { messageId: 'msg-1' })).resolves.toEqual(data)
  })

  it('throws a typed error when the 2xx envelope is invalid', async () => {
    const { fetchFn } = mockFetch(
      () => new Response(JSON.stringify({ ok: false }), { status: 200 }),
    )
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages returned an unexpected response',
    )
  })

  it('throws when the response echoes a different requestId', async () => {
    const { fetchFn } = mockFetch((record) => {
      const { operation } = requestSchema.parse(record.body)
      return new Response(
        JSON.stringify({
          ok: true,
          requestId: crypto.randomUUID(),
          operation,
          data: { value: [] },
        }),
        { status: 200 },
      )
    })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages response requestId does not match',
    )
  })

  it('throws when the response echoes a different operation', async () => {
    const { fetchFn } = mockFetch((record) => {
      const { requestId } = requestSchema.parse(record.body)
      return new Response(
        JSON.stringify({ ok: true, requestId, operation: 'get_message', data: { value: [] } }),
        { status: 200 },
      )
    })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages response operation does not match',
    )
  })

  it('throws a typed error on non-2xx responses', async () => {
    const { fetchFn } = mockFetch(() => new Response('nope', { status: 500 }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages failed with HTTP status 500',
    )
  })

  it('marks draft write HTTP failures as ambiguous without leaking status body', async () => {
    const { fetchFn } = mockFetch(() => new Response('secret draft body', { status: 500 }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    const error = await client
      .call('create_draft', { to: ['to@example.edu'], subject: 'Subject', body: 'Sensitive' })
      .catch((e: unknown) => e)
    expect(error).toBeInstanceOf(PowerAutomateAmbiguousWriteError)
    if (error instanceof PowerAutomateAmbiguousWriteError) {
      expect(error.message).toContain('write outcome is ambiguous')
      expect(error.message).toContain('Inspect Drafts before retrying')
      expect(error.message).not.toContain('secret draft body')
      expect(error.message).not.toContain('Sensitive')
    }
  })

  it('marks draft write timeouts as ambiguous', async () => {
    const fetchFn: typeof fetch = (input, init) =>
      new Promise((resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(init?.signal?.reason))
      })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
      timeoutMs: 20,
    })

    await expect(
      client.call('add_draft_attachment', {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      }),
    ).rejects.toThrow('write outcome is ambiguous')
  })

  it('rejects a 2xx response with a non-JSON body without leaking the URL', async () => {
    const { fetchFn } = mockFetch(() => new Response('not json', { status: 200 }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    const error = await client.call('list_messages', { top: 5 }).catch((e: unknown) => e)
    expect(error).toBeInstanceOf(PowerAutomateError)
    if (error instanceof PowerAutomateError) {
      expect(error.message).not.toContain('https://example.test')
    }
  })

  it('logs a structured entry without leaking the URL or body', async () => {
    const spy = vi.spyOn(console, 'log').mockImplementation(() => {})
    try {
      const { fetchFn } = mockFetch((record) =>
        successResponse(record, { value: [{ id: 'secret-body' }] }),
      )
      const client = new PowerAutomateClient({
        baseUrl: 'https://example.test/flow',
        gatewayKey: 'test-gateway-key',
        fetchFn,
      })

      await client.call('create_draft', {
        to: ['to@example.edu'],
        subject: 'Secret subject',
        body: 'secret-body',
      })

      expect(spy).toHaveBeenCalledTimes(1)
      const logged = JSON.parse(spy.mock.calls[0]?.[0] ?? '{}')
      expect(logged.type).toBe('power_automate_request')
      expect(logged.operation).toBe('create_draft')
      expect(logged.requestId).toBeTruthy()
      expect(logged.success).toBe(true)
      const raw = spy.mock.calls.map((args) => args.join(' ')).join('\n')
      expect(raw).not.toContain('https://example.test')
      expect(raw).not.toContain('secret-body')
      expect(raw).not.toContain('Secret subject')
      expect(raw).not.toContain('test-gateway-key')
    } finally {
      spy.mockRestore()
    }
  })

  it('links each operation to its args at the type level', async () => {
    const { fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    })

    // This is a compile-time-only assertion: the @ts-expect-error proves the
    // mismatched args fail typecheck. At runtime the mock echoes whatever
    // operation it receives, so nothing is actually verified here.
    const wrong = async (): Promise<void> => {
      // @ts-expect-error — get_message takes { messageId }, not { top }
      await client.call('get_message', { top: 20 })
    }
    await wrong()
  })

  it('invokes the global fetch correctly when no fetchFn is injected', async () => {
    const globalFetch = vi.fn<typeof fetch>(async (input, init) => {
      const body = JSON.parse(typeof init?.body === 'string' ? init.body : '{}')
      const { operation, requestId } = requestSchema.parse(body)
      return new Response(JSON.stringify({ ok: true, requestId, operation, data: { value: [] } }), {
        status: 200,
      })
    })
    vi.stubGlobal('fetch', globalFetch)
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
    })

    await client.call('list_messages', { top: 5 })

    expect(globalFetch).toHaveBeenCalledOnce()
    const call = globalFetch.mock.calls[0]
    expect(call?.[0]).toBe('https://example.test/flow')
    const headers = new Headers(call?.[1]?.headers)
    expect(headers.get('X-MCP-Gateway-Key')).toBe('test-gateway-key')
    expect(call?.[1]?.redirect).toBe('manual')
    const body = call?.[1]?.body
    if (typeof body !== 'string') throw new Error('expected a string request body')
    const parsed = requestSchema.parse(JSON.parse(body))
    expect(parsed.operation).toBe('list_messages')
    expect(parsed.args).toEqual({ top: 5 })
  })

  it('aborts on timeout', async () => {
    const fetchFn: typeof fetch = (input, init) =>
      new Promise((resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(init?.signal?.reason))
      })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
      timeoutMs: 20,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages timed out',
    )
  })
  it('rejects oversized Content-Length without reading the response', async () => {
    const cancel = vi.fn()
    const fetchFn: typeof fetch = async () =>
      new Response(new ReadableStream({ cancel }), {
        headers: { 'Content-Length': String(7 * 1024 * 1024) },
      })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test',
      gatewayKey: 'test',
      fetchFn,
    })
    await expect(
      client.call('get_attachment', { messageId: 'message', attachmentId: 'attachment' }),
    ).rejects.toThrow('size limit')
    expect(cancel).toHaveBeenCalledOnce()
  })
  it('caps bytes even without a truthful Content-Length', async () => {
    const cancel = vi.fn()
    const fetchFn: typeof fetch = async () =>
      new Response(
        new ReadableStream({
          start(controller) {
            controller.enqueue(new Uint8Array(256 * 1024 + 1))
          },
          cancel,
        }),
        { headers: { 'Content-Length': '1' } },
      )
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test',
      gatewayKey: 'test',
      fetchFn,
    })
    await expect(
      client.call('list_attachments', { messageId: 'message', top: 1, skip: 0 }),
    ).rejects.toThrow('size limit')
    expect(cancel).toHaveBeenCalledOnce()
  })
  it('bounds stalled response-body reads with the same request timeout', async () => {
    const cancel = vi.fn()
    const fetchFn: typeof fetch = async () => new Response(new ReadableStream({ cancel }))
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test',
      gatewayKey: 'test',
      fetchFn,
      timeoutMs: 20,
    })
    await expect(
      client.call('list_attachments', { messageId: 'message', top: 1, skip: 0 }),
    ).rejects.toThrow('timed out')
    expect(cancel).toHaveBeenCalledOnce()
  })
  it('rejects upstream redirects without forwarding gateway credentials', async () => {
    const fetchFn = vi.fn<typeof fetch>(
      async () =>
        new Response(null, {
          status: 302,
          headers: { Location: 'https://untrusted.example.test/' },
        }),
    )
    const client = new PowerAutomateClient({
      baseUrl: 'https://flow.example.test',
      gatewayKey: 'test',
      fetchFn,
    })
    await expect(
      client.call('list_attachments', { messageId: 'message', top: 1, skip: 0 }),
    ).rejects.toThrow('HTTP status 302')
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(fetchFn.mock.calls[0]?.[1]?.redirect).toBe('manual')
  })
})

describe.each(['create_draft', 'create_reply_draft', 'add_draft_attachment'] as const)(
  'ambiguous %s responses',
  (operation) => {
    it.each(['malformed', 'false-success', 'wrong-request', 'wrong-operation', 'network'] as const)(
      'sanitizes %s and makes no retry',
      async (failure) => {
        const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
          if (failure === 'network') throw new Error('PRIVATE-SENTINEL')
          if (failure === 'malformed') return new Response('PRIVATE-SENTINEL')
          const request = z
            .object({ requestId: z.string() })
            .parse(JSON.parse(z.string().parse(init?.body)))
          return Response.json({
            ok: failure !== 'false-success',
            requestId: failure === 'wrong-request' ? crypto.randomUUID() : request.requestId,
            operation: failure === 'wrong-operation' ? 'unrelated' : operation,
            data: { id: 'PRIVATE-SENTINEL', isDraft: true },
          })
        })
        const client = new PowerAutomateClient({
          baseUrl: 'https://example.test/flow',
          gatewayKey: 'test',
          fetchFn,
        })
        const args =
          operation === 'create_draft'
            ? { to: ['to@example.edu'], subject: 'Subject', body: 'Body' }
            : operation === 'create_reply_draft'
              ? { messageId: 'source', body: 'Body' }
              : { draftId: 'draft', name: 'a.txt', contentType: 'text/plain', contentBytes: 'YQ==' }
        const outcome = client.call(operation, args)
        await expect(outcome).rejects.toThrow('Inspect Drafts before retrying')
        await expect(outcome).rejects.not.toThrow('PRIVATE-SENTINEL')
        expect(fetchFn).toHaveBeenCalledOnce()
      },
    )
  },
)
