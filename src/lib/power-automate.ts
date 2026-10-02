import { z } from 'zod'

export interface MailListFilters {
  isRead?: boolean | undefined
  hasAttachments?: boolean | undefined
  receivedAfter?: string | undefined
  receivedBefore?: string | undefined
}

export interface PowerAutomateOperations {
  list_messages: {
    args: {
      top: number
      skip?: number
      mailbox?: 'inbox' | 'sent' | 'all'
      folderId?: string
      filters?: MailListFilters
    }
  }
  search_messages: {
    args: { query: string; top: number; mailbox?: 'inbox' | 'sent' | 'all'; folderId?: string }
  }
  get_message: { args: { messageId: string } }
  list_mail_folders: { args: { top: number } }
  get_conversation: { args: { conversationId: string; top: number; skip?: number } }
  list_attachments: { args: { messageId: string; top: number; skip: number } }
  get_attachment: { args: { messageId: string; attachmentId: string } }
  create_draft: {
    args: {
      to: string[]
      cc?: string[] | undefined
      bcc?: string[] | undefined
      subject: string
      body: string
    }
  }
  create_reply_draft: { args: { messageId: string; body: string } }
  add_draft_attachment: {
    args: { draftId: string; name: string; contentType: string; contentBytes: string }
  }
  onedrive_search_files: { args: { query: string; top: number } }
  // Internal folder window top is 1..1000; MCP page limits stay 1..100.
  onedrive_list_folder: { args: { folderId?: string; top: number } }
  onedrive_get_metadata: { args: { fileId: string } }
  onedrive_get_content: { args: { fileId: string } }
}

export type PowerAutomateOperation = keyof PowerAutomateOperations

export interface PowerAutomateClientOptions {
  baseUrl: string
  gatewayKey: string
  fetchFn?: typeof fetch
  timeoutMs?: number
}

const successEnvelopeSchema = z.object({
  ok: z.literal(true),
  requestId: z.string().uuid(),
  operation: z.string(),
  data: z.unknown(),
})

const ambiguousWriteOperations = new Set<PowerAutomateOperation>([
  'create_draft',
  'create_reply_draft',
  'add_draft_attachment',
])

export class PowerAutomateError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'PowerAutomateError'
  }
}

export class PowerAutomateAmbiguousWriteError extends PowerAutomateError {
  constructor(operation: PowerAutomateOperation, cause: 'failed' | 'timed out') {
    super(
      `${operation} ${cause}; write outcome is ambiguous. Inspect Drafts before retrying and never blind retry.`,
    )
    this.name = 'PowerAutomateAmbiguousWriteError'
  }
}

export class PowerAutomateClient {
  private readonly baseUrl: string
  private readonly gatewayKey: string
  private readonly fetchFn: typeof fetch
  private readonly timeoutMs: number

  constructor(options: PowerAutomateClientOptions) {
    this.baseUrl = options.baseUrl
    this.gatewayKey = options.gatewayKey
    this.fetchFn = options.fetchFn ?? ((input, init) => fetch(input, init))
    this.timeoutMs = options.timeoutMs ?? 30000
  }

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
        redirect: 'manual',
        headers: { 'Content-Type': 'application/json', 'X-MCP-Gateway-Key': this.gatewayKey },
        body: JSON.stringify({ operation, requestId, args }),
        signal,
      })
      status = response.status
      if (!response.ok) {
        await response.body?.cancel()
        if (ambiguousWriteOperations.has(operation))
          throw new PowerAutomateAmbiguousWriteError(operation, 'failed')
        throw new PowerAutomateError(`${operation} failed with HTTP status ${response.status}`)
      }
      const body = await readBoundedJson(response, signal, operation)
      const parsed = successEnvelopeSchema.safeParse(body)
      if (!parsed.success)
        throw new PowerAutomateError(`${operation} returned an unexpected response`)
      if (parsed.data.requestId !== requestId)
        throw new PowerAutomateError(`${operation} response requestId does not match`)
      if (parsed.data.operation !== operation)
        throw new PowerAutomateError(`${operation} response operation does not match`)
      success = true
      return parsed.data.data
    } catch (error) {
      if (ambiguousWriteOperations.has(operation)) {
        if (error instanceof PowerAutomateAmbiguousWriteError) throw error
        throw new PowerAutomateAmbiguousWriteError(
          operation,
          signal.aborted ? 'timed out' : 'failed',
        )
      }
      if (error instanceof PowerAutomateError) throw error
      if (signal.aborted) throw new PowerAutomateError(`${operation} timed out`)
      throw new PowerAutomateError(`${operation} failed to complete the request`)
    } finally {
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

async function readBoundedJson(
  response: Response,
  signal: AbortSignal,
  operation: PowerAutomateOperation,
): Promise<unknown> {
  const listOperations = new Set<PowerAutomateOperation>([
    'list_messages',
    'search_messages',
    'list_mail_folders',
    'get_conversation',
    'list_attachments',
    'onedrive_search_files',
    'onedrive_list_folder',
    'onedrive_get_metadata',
  ])
  const maxBytes =
    operation === 'onedrive_list_folder'
      ? 4 * 1024 * 1024
      : listOperations.has(operation)
        ? 256 * 1024
        : 6 * 1024 * 1024
  const length = Number(response.headers.get('content-length') ?? 0)
  if (length > maxBytes) {
    await response.body?.cancel()
    throw new PowerAutomateError(`${operation} response exceeds the size limit`)
  }
  if (!response.body) throw new PowerAutomateError(`${operation} returned an empty response`)
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let bytes = 0
  const abort = () => void reader.cancel().catch(() => {})
  signal.addEventListener('abort', abort, { once: true })
  try {
    while (true) {
      if (signal.aborted) throw new PowerAutomateError(`${operation} timed out`)
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
    try {
      return JSON.parse(new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(data))
    } catch {
      throw new PowerAutomateError(`${operation} returned malformed JSON`)
    }
  } finally {
    signal.removeEventListener('abort', abort)
    reader.releaseLock()
  }
}
