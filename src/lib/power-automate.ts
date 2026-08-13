/**
 * Stateless client for the Power Automate HTTP trigger. Each call POSTs a
 * JSON envelope with an operation, a fresh requestId, and args, and unwraps
 * the `{ ok: true, data }` success envelope from the response.
 */
import { z } from 'zod'

export type PowerAutomateOperation = 'list_messages' | 'search_messages' | 'get_message'

export type PowerAutomateArgs =
  | { top: number }
  | { query: string; top: number }
  | { messageId: string }

export interface PowerAutomateClientOptions {
  baseUrl: string
  /** Injectable for tests; defaults to the global fetch. */
  fetchFn?: typeof fetch
  /** Injectable for tests; defaults to 30s. */
  timeoutMs?: number
}

/** Success envelope returned by the Power Automate flow for 2xx responses. */
const successEnvelopeSchema = z.object({
  ok: z.literal(true),
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

export class PowerAutomateClient {
  private readonly baseUrl: string
  private readonly fetchFn: typeof fetch
  private readonly timeoutMs: number

  constructor(options: PowerAutomateClientOptions) {
    this.baseUrl = options.baseUrl
    // Wrap the global fetch so it is always invoked as a plain function call.
    // Calling the native fetch as a method (this.fetchFn(...)) loses its `this`
    // binding and throws "Illegal invocation" on the Workers runtime.
    this.fetchFn = options.fetchFn ?? ((input, init) => fetch(input, init))
    this.timeoutMs = options.timeoutMs ?? 30000
  }

  /**
   * POSTs an operation to Power Automate and returns the Graph response
   * (the `data` field of the success envelope).
   * @throws {PowerAutomateError} on non-2xx responses, timeouts, and network errors.
   */
  async call(operation: PowerAutomateOperation, args: PowerAutomateArgs): Promise<unknown> {
    const requestId = crypto.randomUUID()
    console.log(`[power-automate] operation=${operation} requestId=${requestId}`)
    const signal = AbortSignal.timeout(this.timeoutMs)
    try {
      const response = await this.fetchFn(this.baseUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ operation, requestId, args }),
        signal,
      })
      if (!response.ok) {
        throw new PowerAutomateError(`${operation} failed with HTTP status ${response.status}`)
      }
      const body: unknown = await response.json()
      const parsed = successEnvelopeSchema.safeParse(body)
      if (!parsed.success) {
        throw new PowerAutomateError(`${operation} returned an unexpected response`)
      }
      return parsed.data.data
    } catch (error) {
      if (error instanceof PowerAutomateError) {
        throw error
      }
      if (signal.aborted) {
        throw new PowerAutomateError(`${operation} timed out`)
      }
      throw new PowerAutomateError(`${operation} failed to complete the request`)
    }
  }
}
