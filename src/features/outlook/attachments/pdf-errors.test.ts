import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { getResolvedPDFJS } from 'unpdf'
import type * as Unpdf from 'unpdf'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('unpdf', async (importOriginal) => {
  const actual = await importOriginal<typeof Unpdf>()
  return { ...actual, getResolvedPDFJS: vi.fn(actual.getResolvedPDFJS) }
})

import { PowerAutomateClient } from '../../../lib/power-automate.js'
import { createOutlookMcpServer } from '../server.js'
import { makeEncryptedPdf, makePdf } from './fixtures-pdf.js'
import { getPdfDiagnosticCode, PDF_DIAGNOSTICS, PdfParseError } from './pdf-errors.js'
import { inspectPdf, readPdf } from './pdf.js'
import { syntheticAttachment } from './synthetic-fixtures.js'

const secret = 'SYNTHETIC_PRIVATE_ERROR https://private.invalid/file.pdf'
const encoder = new TextEncoder()

async function failure(run: () => Promise<unknown>) {
  try {
    await run()
  } catch (error) {
    return error
  }
  throw new Error('Expected rejection')
}

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('closed PDF diagnostics', () => {
  it.each([
    ['PDF_RAW_SIZE', () => new Uint8Array()],
    ['PDF_RAW_SIZE', () => new Uint8Array(4 * 1024 * 1024 + 1)],
  ] as const)('%s identifies pre-parser raw input limits', async (code, bytes) => {
    vi.mocked(getResolvedPDFJS).mockClear()
    const error = await failure(() => inspectPdf(bytes()))
    expect(error).toBeInstanceOf(PdfParseError)
    expect(getPdfDiagnosticCode(error)).toBe(code)
    expect(getResolvedPDFJS).not.toHaveBeenCalled()
  })

  it('does not trust mutated error properties, forged codes, or the code map prototype', () => {
    const error = new PdfParseError('PDF_LOAD')
    error.message = secret
    Object.defineProperty(error, 'code', { value: secret })
    expect(getPdfDiagnosticCode(error)).toBe('PDF_LOAD')
    expect(
      getPdfDiagnosticCode(Object.assign(new Error(secret), { code: 'PDF_LOAD' })),
    ).toBeUndefined()
    const invalid = Reflect.construct(PdfParseError, [secret])
    expect(getPdfDiagnosticCode(invalid)).toBe('PDF_UNKNOWN')
    const inherited = Reflect.construct(PdfParseError, ['toString'])
    expect(getPdfDiagnosticCode(inherited)).toBe('PDF_UNKNOWN')
    expect(Object.isFrozen(PDF_DIAGNOSTICS)).toBe(true)
  })

  it('labels module initialization failures without exposing their text', async () => {
    vi.mocked(getResolvedPDFJS).mockRejectedValueOnce(new Error(secret))
    const error = await failure(() => inspectPdf(makePdf()))
    expect(getPdfDiagnosticCode(error)).toBe('PDF_INITIALIZATION')
    expect(String(error)).not.toContain(secret)
  })

  it('sanitizes exceptions whose name accessor itself throws', async () => {
    const hostile = new Error(secret)
    Object.defineProperty(hostile, 'name', {
      get() {
        throw new Error(secret)
      },
    })
    vi.mocked(getResolvedPDFJS).mockRejectedValueOnce(hostile)
    const error = await failure(() => inspectPdf(makePdf()))
    expect(getPdfDiagnosticCode(error)).toBe('PDF_INITIALIZATION')
    expect(String(error)).not.toContain(secret)
  })

  it('rejects a genuine password-protected synthetic PDF through inspect and read', async () => {
    for (const parse of [inspectPdf, readPdf]) {
      const error = await failure(() => parse(makeEncryptedPdf()))
      expect(getPdfDiagnosticCode(error)).toBe('PDF_PASSWORD')
      expect(String(error)).not.toContain('synthetic-fixture-password')
    }
  })

  it('labels loading task creation failures without exposing their text', async () => {
    const pdfjs = await getResolvedPDFJS()
    vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
      ...pdfjs,
      getDocument: () => {
        throw new Error(secret)
      },
    })
    const error = await failure(() => inspectPdf(makePdf()))
    expect(getPdfDiagnosticCode(error)).toBe('PDF_CREATE_DOCUMENT')
    expect(String(error)).not.toContain(secret)
  })

  it.each(['load', 'extraction'] as const)(
    'labels %s failure and destroys the loading task',
    async (stage) => {
      const pdfjs = await getResolvedPDFJS()
      let destroyed = false
      vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
        ...pdfjs,
        getDocument: (parameters) => {
          const task = pdfjs.getDocument(parameters)
          const promise = task.promise.then((document) => {
            if (stage === 'load') throw new Error(secret)
            vi.spyOn(document, 'getPage').mockRejectedValue(new Error(secret))
            return document
          })
          Object.defineProperty(task, 'promise', { value: promise })
          const destroy = task.destroy.bind(task)
          vi.spyOn(task, 'destroy').mockImplementation(async () => {
            destroyed = true
            await destroy()
          })
          return task
        },
      })
      const error = await failure(() => readPdf(makePdf()))
      expect(getPdfDiagnosticCode(error)).toBe(stage === 'load' ? 'PDF_LOAD' : 'PDF_EXTRACTION')
      expect(String(error)).not.toContain(secret)
      expect(destroyed).toBe(true)
    },
  )

  it('labels invalid text chunks without exposing their contents', async () => {
    const pdfjs = await getResolvedPDFJS()
    let destroyed = false
    vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
      ...pdfjs,
      getDocument: (parameters) => {
        const task = pdfjs.getDocument(parameters)
        const promise = task.promise.then((document) => {
          const getPage = document.getPage.bind(document)
          vi.spyOn(document, 'getPage').mockImplementation(async (number) => {
            const page = await getPage(number)
            Object.defineProperty(page, 'streamTextContent', {
              value: () =>
                new ReadableStream({
                  start(controller) {
                    controller.enqueue({ private: secret })
                    controller.close()
                  },
                }),
            })
            return page
          })
          return document
        })
        Object.defineProperty(task, 'promise', { value: promise })
        const destroy = task.destroy.bind(task)
        vi.spyOn(task, 'destroy').mockImplementation(async () => {
          destroyed = true
          await destroy()
        })
        return task
      },
    })
    const error = await failure(() => readPdf(makePdf()))
    expect(getPdfDiagnosticCode(error)).toBe('PDF_TEXT_CHUNK')
    expect(String(error)).not.toContain(secret)
    expect(destroyed).toBe(true)
  })

  it('labels timeouts and preserves cleanup without changing the deadline', async () => {
    const pdfjs = await getResolvedPDFJS()
    let destroyed = false
    vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
      ...pdfjs,
      getDocument: (parameters) => {
        const task = pdfjs.getDocument(parameters)
        Object.defineProperty(task, 'promise', { value: new Promise(() => undefined) })
        const destroy = task.destroy.bind(task)
        vi.spyOn(task, 'destroy').mockImplementation(async () => {
          destroyed = true
          await destroy()
        })
        return task
      },
    })
    vi.useFakeTimers()
    const pending = failure(() => inspectPdf(makePdf()))
    await vi.advanceTimersByTimeAsync(10001)
    expect(getPdfDiagnosticCode(await pending)).toBe('PDF_TIMEOUT')
    expect(destroyed).toBe(true)
    expect(vi.getTimerCount()).toBe(0)
  })

  it('sanitizes malformed and encrypted-looking library failures', async () => {
    for (const bytes of [
      encoder.encode('not a PDF ' + secret),
      makePdf({ trailer: '/Encrypt 3 0 R' }),
    ]) {
      const error = await failure(() => inspectPdf(bytes))
      expect(error).toBeInstanceOf(PdfParseError)
      expect(getPdfDiagnosticCode(error)).toBeDefined()
      expect(String(error)).not.toContain(secret)
    }
  })
})

async function mcpFailure(
  tool: 'outlook_inspect_attachment' | 'outlook_read_attachment',
  bytes: Uint8Array,
) {
  const file = {
    ...syntheticAttachment('pdf'),
    size: bytes.length,
    contentBytes: btoa(String.fromCharCode(...bytes)),
  }
  const flow = new PowerAutomateClient({
    baseUrl: 'https://example.test/private',
    gatewayKey: 'synthetic',
    fetchFn: async (_url, init) => {
      const parsed: unknown = JSON.parse(typeof init?.body === 'string' ? init.body : '{}')
      if (typeof parsed !== 'object' || !parsed) throw new Error('Invalid synthetic request')
      return Response.json({ ...parsed, ok: true, data: file })
    },
  })
  const mcp = new Client({ name: 'diagnostic-test', version: '1' })
  const server = createOutlookMcpServer(flow)
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), server.connect(b)])
  try {
    return await mcp.request({
      method: 'tools/call',
      params: {
        name: tool,
        arguments: {
          messageId: 'msg-1',
          attachmentId: 'att-1',
          ...(tool === 'outlook_read_attachment'
            ? { selection: { format: 'pdf', pageStart: 1, pageEnd: 1 } }
            : {}),
        },
      },
    })
  } finally {
    await mcp.close()
    await server.close()
  }
}

describe('PDF diagnostics through actual MCP boundary', () => {
  it.each(['outlook_inspect_attachment', 'outlook_read_attachment'] as const)(
    '%s returns only generic text and a fixed reason',
    async (tool) => {
      const result = await mcpFailure(tool, encoder.encode('not a PDF ' + secret))
      expect(result.isError).toBe(true)
      expect(result.content).toEqual([
        {
          type: 'text',
          text: 'Attachment could not be read safely: unsupported, malformed, encrypted or over parser limits [PDF_LOAD]',
        },
      ])
      expect(result.structuredContent).toBeUndefined()
      expect(JSON.stringify(result)).not.toContain(secret)
    },
  )

  it('reconstructs typed errors instead of returning mutated messages or codes', async () => {
    const error = new PdfParseError('PDF_LOAD')
    error.message = secret
    Object.defineProperty(error, 'code', { value: secret })
    vi.mocked(getResolvedPDFJS).mockRejectedValueOnce(error)
    const result = await mcpFailure('outlook_inspect_attachment', makePdf())
    expect(JSON.stringify(result)).toContain('[PDF_LOAD]')
    expect(JSON.stringify(result)).not.toContain(secret)
  })

  it('unknown exceptions cannot forge a code or leak content and no parser logs occur', async () => {
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined)
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const errorLog = vi.spyOn(console, 'error').mockImplementation(() => undefined)
    vi.mocked(getResolvedPDFJS).mockRejectedValueOnce(
      Object.assign(new Error(secret), { code: secret }),
    )
    const result = await mcpFailure('outlook_inspect_attachment', makePdf())
    expect(JSON.stringify(result)).toContain('[PDF_INITIALIZATION]')
    expect(JSON.stringify(result)).not.toContain(secret)
    // PowerAutomateClient emits its existing structured transport event only.
    expect(warn).not.toHaveBeenCalled()
    expect(errorLog).not.toHaveBeenCalled()
    for (const call of log.mock.calls) expect(JSON.stringify(call)).not.toContain(secret)
  })
})

it.each(['outlook_inspect_attachment', 'outlook_read_attachment'] as const)(
  '%s sanitizes a genuine password-protected PDF',
  async (tool) => {
    const result = await mcpFailure(tool, makeEncryptedPdf())
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('[PDF_PASSWORD]')
    expect(JSON.stringify(result)).not.toContain('synthetic-fixture-password')
  },
)
