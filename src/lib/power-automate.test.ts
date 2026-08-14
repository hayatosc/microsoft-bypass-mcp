import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient, PowerAutomateError } from './power-automate.js'

afterEach(() => {
  vi.unstubAllGlobals()
})

const requestSchema = z.object({
  operation: z.enum(['list_messages', 'search_messages', 'get_message']),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})

type RequestRecord = { url: string; body: unknown }

function mockFetch(respond: (record: RequestRecord) => Response): {
  records: RequestRecord[]
  fetchFn: typeof fetch
} {
  const records: RequestRecord[] = []
  const fetchFn: typeof fetch = async (input, init) => {
    const record = {
      url: typeof input === 'string' ? input : input instanceof URL ? input.href : input.url,
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
    ]
    const { records, fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

    for (const { operation, args } of cases) {
      await client.call(operation, args)
    }

    expect(records).toHaveLength(cases.length)
    expect(records[0]?.url).toBe('https://example.test/flow')
    const sent = cases.map((_, i) => requestSchema.parse(records[i]?.body))
    expect(sent.map((s) => s.operation)).toEqual([
      'list_messages',
      'search_messages',
      'get_message',
    ])
    expect(sent.map((s) => s.args)).toEqual([
      { top: 5 },
      { query: 'PMDA', top: 10 },
      { messageId: 'msg-1' },
    ])
  })

  it('generates a fresh requestId per call', async () => {
    const { records, fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

    await client.call('list_messages', { top: 5 })
    await client.call('list_messages', { top: 5 })

    const requestIds = records.map((r) => requestSchema.parse(r.body).requestId)
    expect(new Set(requestIds).size).toBe(2)
  })

  it('parses the response JSON and unwraps the envelope', async () => {
    const data = { value: [{ id: 'msg-1' }] }
    const { fetchFn } = mockFetch((record) => successResponse(record, data))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

    await expect(client.call('get_message', { messageId: 'msg-1' })).resolves.toEqual(data)
  })

  it('throws a typed error when the 2xx envelope is invalid', async () => {
    const { fetchFn } = mockFetch(
      () => new Response(JSON.stringify({ ok: false }), { status: 200 }),
    )
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

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
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

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
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages response operation does not match',
    )
  })

  it('throws a typed error on non-2xx responses', async () => {
    const { fetchFn } = mockFetch(() => new Response('nope', { status: 500 }))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages failed with HTTP status 500',
    )
  })

  it('rejects a 2xx response with a non-JSON body without leaking the URL', async () => {
    const { fetchFn } = mockFetch(() => new Response('not json', { status: 200 }))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

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
      const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

      await client.call('list_messages', { top: 5 })

      expect(spy).toHaveBeenCalledTimes(1)
      const logged = JSON.parse(spy.mock.calls[0]?.[0] ?? '{}')
      expect(logged.type).toBe('power_automate_request')
      expect(logged.operation).toBe('list_messages')
      expect(logged.requestId).toBeTruthy()
      expect(logged.success).toBe(true)
      const raw = spy.mock.calls.map((args) => args.join(' ')).join('\n')
      expect(raw).not.toContain('https://example.test')
      expect(raw).not.toContain('secret-body')
    } finally {
      spy.mockRestore()
    }
  })

  it('links each operation to its args at the type level', async () => {
    const { fetchFn } = mockFetch((record) => successResponse(record, { value: [] }))
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow', fetchFn })

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
    const client = new PowerAutomateClient({ baseUrl: 'https://example.test/flow' })

    await client.call('list_messages', { top: 5 })

    expect(globalFetch).toHaveBeenCalledOnce()
    const call = globalFetch.mock.calls[0]
    expect(call?.[0]).toBe('https://example.test/flow')
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
      fetchFn,
      timeoutMs: 20,
    })

    await expect(client.call('list_messages', { top: 5 })).rejects.toThrow(
      'list_messages timed out',
    )
  })
})
