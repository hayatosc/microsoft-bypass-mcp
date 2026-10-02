/**
 * Stateless client for the Power Automate HTTP trigger. Each call POSTs a
 * JSON envelope with an operation, a fresh requestId, and args, and unwraps
 * the `{ ok, requestId, operation, data }` success envelope from the response.
 * The operation name is linked to its request args at the type level, so a
 * mismatched operation/args pair fails to compile.
 */
import { z } from 'zod'

/**
 * The fixed operation surface Power Automate implements. Each operation maps a
 * name to its request args. The Graph response shape is validated downstream by
 * the feature normalizers, so it stays `unknown` here.
 */
export interface PowerAutomateOperations {
  list_messages: { args: { top: number } }
  search_messages: { args: { query: string; top: number } }
  get_message: { args: { messageId: string } }
  list_attachments: { args: { messageId: string; top: number; skip: number } }
  get_attachment: { args: { messageId: string; attachmentId: string } }
}

/** Names of the operations Power Automate implements. */
export type PowerAutomateOperation = keyof PowerAutomateOperations

/** Options for constructing a {@link PowerAutomateClient}. */
export interface PowerAutomateClientOptions {
  baseUrl: string
  /** Gateway key sent as the `X-MCP-Gateway-Key` header required by the trigger. */
  gatewayKey: string
  /** Injectable for tests; defaults to the global fetch. */
  fetchFn?: typeof fetch
  /** Injectable for tests; defaults to 30s. */
  timeoutMs?: number
}

/**
 * Success envelope returned by the Power Automate flow for 2xx responses.
 * `requestId` and `operation` are echoes of the request and are checked below.
 */
const successEnvelopeSchema = z.object({
  ok: z.literal(true),
  requestId: z.string().uuid(),
  operation: z.string(),
  data: z.unknown(),
})

/**
 * Error for failed Power Automate calls. The message contains at most the
 * operation name and HTTP status — never the URL, never the response body.
 */
export class PowerAutomateError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'PowerAutomateError'
  }
}

/**
 * Stateless client that POSTs a single operation to the Power Automate HTTP
 * trigger and unwraps the validated success envelope.
 */
export class PowerAutomateClient {
  private readonly baseUrl: string
  private readonly gatewayKey: string
  private readonly fetchFn: typeof fetch
  private readonly timeoutMs: number

  constructor(options: PowerAutomateClientOptions) {
    this.baseUrl = options.baseUrl
    this.gatewayKey = options.gatewayKey
    // Wrap the global fetch so it is always invoked as a plain function call.
    // Calling the native fetch as a method (this.fetchFn(...)) loses its `this`
    // binding and throws "Illegal invocation" on the Workers runtime.
    this.fetchFn = options.fetchFn ?? ((input, init) => fetch(input, init))
    this.timeoutMs = options.timeoutMs ?? 30000
  }

  /**
   * POSTs an operation to Power Automate and returns the Graph response
   * (the `data` field of the success envelope).
   * @throws {PowerAutomateError} on non-2xx responses, timeouts, network errors,
   *   and envelope mismatches.
   */
  async call<K extends PowerAutomateOperation>(
    operation: K,
    args: PowerAutomateOperations[K]['args'],
  ): Promise<unknown> {
    const requestId = crypto.randomUUID()
    const startedAt = Date.now()
    const signal = AbortSignal.timeout(this.timeoutMs)
    let status = 0
    let success = false
    try {
      const response = await this.fetchFn(this.baseUrl, {
        method: 'POST',
        // Workers supports manual/follow only. Reject 3xx below without forwarding the gateway key.
        redirect: 'manual',
        headers: {
          'Content-Type': 'application/json',
          'X-MCP-Gateway-Key': this.gatewayKey,
        },
        body: JSON.stringify({ operation, requestId, args }),
        signal,
      })
      status = response.status
      if (!response.ok) {
        await response.body?.cancel()
        throw new PowerAutomateError(`${operation} failed with HTTP status ${response.status}`)
      }
      // Bound bytes before JSON parsing; Content-Length alone is untrusted.
      const body: unknown = await readBoundedJson(response, signal, operation)
      const parsed = successEnvelopeSchema.safeParse(body)
      if (!parsed.success) {
        throw new PowerAutomateError(`${operation} returned an unexpected response`)
      }
      if (parsed.data.requestId !== requestId) {
        throw new PowerAutomateError(`${operation} response requestId does not match`)
      }
      if (parsed.data.operation !== operation) {
        throw new PowerAutomateError(`${operation} response operation does not match`)
      }
      success = true
      return parsed.data.data
    } catch (error) {
      if (error instanceof PowerAutomateError) {
        throw error
      }
      if (signal.aborted) {
        throw new PowerAutomateError(`${operation} timed out`)
      }
      throw new PowerAutomateError(`${operation} failed to complete the request`)
    } finally {
      // Structured log: operation + requestId + timing only. Never the URL,
      // args, or response body.
      console.log(
        JSON.stringify({
          type: 'power_automate_request',
          requestId,
          operation,
          durationMs: Date.now() - startedAt,
          status,
          success,
        }),
      )
    }
  }
}

/** Prevent a malformed/upstream response from allocating an unbounded JSON body. */
async function readBoundedJson(
  response: Response,
  signal: AbortSignal,
  operation: PowerAutomateOperation,
): Promise<unknown> {
  const maxBytes = operation === 'list_attachments' ? 256 * 1024 : 6 * 1024 * 1024
  const length = Number(response.headers.get('content-length') ?? 0)
  if (length > maxBytes) {
    await response.body?.cancel()
    throw new PowerAutomateError(`${operation} response exceeds the size limit`)
  }
  if (!response.body) throw new PowerAutomateError(`${operation} returned an empty response`)
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let bytes = 0
  const abort = () => {
    void reader.cancel().catch(() => {})
  }
  signal.addEventListener('abort', abort, { once: true })
  try {
    if (signal.aborted) throw new PowerAutomateError(`${operation} timed out`)
    while (true) {
      const chunk = await reader.read()
      if (signal.aborted) throw new PowerAutomateError(`${operation} timed out`)
      if (chunk.done) break
      bytes += chunk.value.byteLength
      if (bytes > maxBytes) {
        await reader.cancel()
        throw new PowerAutomateError(`${operation} response exceeds the size limit`)
      }
      chunks.push(chunk.value)
    }
    const data = new Uint8Array(bytes)
    let offset = 0
    for (const chunk of chunks) {
      data.set(chunk, offset)
      offset += chunk.byteLength
    }
    return JSON.parse(new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(data))
  } finally {
    signal.removeEventListener('abort', abort)
    reader.releaseLock()
  }
}
