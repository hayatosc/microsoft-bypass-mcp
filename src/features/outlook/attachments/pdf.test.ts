import { getResolvedPDFJS } from 'unpdf'
import type * as Unpdf from 'unpdf'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('unpdf', async (importOriginal) => {
  const actual = await importOriginal<typeof Unpdf>()
  return { ...actual, getResolvedPDFJS: vi.fn(actual.getResolvedPDFJS) }
})

import { makeJpegStreamPdf, makeObjectStreamPdf, makePdf } from './fixtures-pdf.js'
import { inspectPdf, PdfRangeError, readPdf } from './pdf.js'

const encoder = new TextEncoder()

function replace(bytes: Uint8Array, search: string, replacement: string): Uint8Array {
  if (search.length !== replacement.length)
    throw new Error('Replacement must preserve xref offsets')
  const text = new TextDecoder().decode(bytes)
  if (!text.includes(search)) throw new Error('Fixture search failed')
  return encoder.encode(text.replace(search, replacement))
}

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('bounded PDF parsing', () => {
  it.each([false, true])('inspects and reads real %s-compressed PDF bytes', async (compressed) => {
    const bytes = makePdf({ compressed })
    expect(await inspectPdf(bytes)).toEqual({ pageCount: 2 })
    const result = await readPdf(bytes, { pageStart: 2, pageEnd: 2 })
    expect(result.pageCount).toBe(2)
    expect(result.pages).toEqual([{ page: 2, text: 'Hello attachment page 2', truncated: false }])
    expect(result.truncated).toBe(false)
  })

  it('lets PDF.js handle object streams, xref streams and JPEG image streams', async () => {
    expect(await inspectPdf(makeObjectStreamPdf())).toEqual({ pageCount: 1 })
    expect((await readPdf(makeObjectStreamPdf())).pages).toEqual([
      { page: 1, text: 'Object stream text', truncated: false },
    ])
    expect(await inspectPdf(makeJpegStreamPdf())).toEqual({ pageCount: 1 })
    expect((await readPdf(makeJpegStreamPdf())).pages).toEqual([
      { page: 1, text: 'JPEG-stream text', truncated: false },
    ])
  })

  it('enforces one aggregate character budget, with explicit per-page truncation', async () => {
    const result = await readPdf(makePdf({ pages: 3 }), { maxCharacters: 8 })
    expect(result.pages).toEqual([
      { page: 1, text: 'Hello at', truncated: true },
      { page: 2, text: '', truncated: true },
      { page: 3, text: '', truncated: true },
    ])
    expect(result.truncated).toBe(true)
    expect(result.pages.reduce((count, page) => count + page.text.length, 0)).toBe(8)
  })

  it('returns empty text for image-only/empty pages rather than claiming OCR', async () => {
    const result = await readPdf(makePdf({ pages: 1, content: 'q Q' }))
    expect(result.pages).toEqual([{ page: 1, text: '', truncated: false }])
  })

  it('does not fetch external fonts, images, CMaps or document URLs, or log content', async () => {
    const fetch = vi.fn(() => Promise.reject(new Error('Network must never be used')))
    vi.stubGlobal('fetch', fetch)
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined)
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const error = vi.spyOn(console, 'error').mockImplementation(() => undefined)
    await readPdf(makePdf({ text: 'Private attachment contents' }))
    expect(fetch).not.toHaveBeenCalled()
    expect(log).not.toHaveBeenCalled()
    expect(warn).not.toHaveBeenCalled()
    expect(error).not.toHaveBeenCalled()
  })

  it.each([
    { pageStart: 0 },
    { pageStart: 3 },
    { pageStart: 2, pageEnd: 1 },
    { pageEnd: 11 },
    { maxCharacters: 0 },
    { maxCharacters: 20001 },
  ])('rejects invalid page/character bounds %j', async (options) => {
    await expect(readPdf(makePdf(), options)).rejects.toThrow(PdfRangeError)
  })

  it('limits reads to ten pages and documents to 200 pages', async () => {
    const bytes = makePdf({ pages: 11 })
    expect((await readPdf(bytes)).pages).toHaveLength(10)
    await expect(readPdf(bytes, { pageEnd: 11 })).rejects.toThrow(PdfRangeError)
    await expect(inspectPdf(makePdf({ pages: 201 }))).rejects.toThrow('safety limits')
  })

  it('rejects empty and oversized raw inputs before PDF.js is initialized', async () => {
    vi.mocked(getResolvedPDFJS).mockClear()
    await expect(inspectPdf(new Uint8Array())).rejects.toThrow('[PDF_RAW_SIZE]')
    await expect(inspectPdf(new Uint8Array(4 * 1024 * 1024 + 1))).rejects.toThrow('[PDF_RAW_SIZE]')
    expect(getResolvedPDFJS).not.toHaveBeenCalled()
  })

  it('fails malformed and encrypted-looking inputs with sanitized parser diagnostics', async () => {
    await expect(inspectPdf(encoder.encode('not a PDF secret'))).rejects.toThrow('safety limits')
    await expect(inspectPdf(makePdf({ trailer: '/Encrypt 3 0 R' }))).rejects.toThrow(
      'safety limits',
    )
  })

  it('does not extract pages outside the selected bounds', async () => {
    const pdfjs = await getResolvedPDFJS()
    const requested: number[] = []
    let destroyed = false
    vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
      ...pdfjs,
      getDocument: (parameters) => {
        const task = pdfjs.getDocument(parameters)
        const promise = task.promise.then((document) => {
          const getPage = document.getPage.bind(document)
          vi.spyOn(document, 'getPage').mockImplementation(async (number) => {
            requested.push(number)
            return await getPage(number)
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
    const result = await readPdf(makePdf({ pages: 3 }), { pageStart: 2, pageEnd: 2 })
    expect(requested).toEqual([2])
    expect(result.pages).toEqual([{ page: 2, text: 'Hello attachment page 2', truncated: false }])
    expect(destroyed).toBe(true)
  })

  it('destroys the loading task on success, invalid selectors, and document-load failures', async () => {
    const pdfjs = await getResolvedPDFJS()
    const tasks: { destroyed: boolean }[] = []
    const instrument = () => {
      vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
        ...pdfjs,
        getDocument: (options) => {
          const task = pdfjs.getDocument(options)
          const state = { destroyed: false }
          tasks.push(state)
          const destroy = task.destroy.bind(task)
          vi.spyOn(task, 'destroy').mockImplementation(async () => {
            state.destroyed = true
            await destroy()
          })
          return task
        },
      })
    }
    instrument()
    await inspectPdf(makePdf())
    instrument()
    await expect(readPdf(makePdf(), { pageEnd: 3 })).rejects.toThrow(PdfRangeError)
    instrument()
    await expect(inspectPdf(replace(makePdf(), '/Root 1 0 R', '/Root 3 0 R'))).rejects.toThrow(
      'safety limits',
    )
    expect(tasks).toEqual([{ destroyed: true }, { destroyed: true }, { destroyed: true }])
  })

  it('cleans up a pending loading task on the extraction timeout', async () => {
    const pdfjs = await getResolvedPDFJS()
    let destroyed = false
    vi.mocked(getResolvedPDFJS).mockResolvedValueOnce({
      ...pdfjs,
      getDocument: (options) => {
        const task = pdfjs.getDocument(options)
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
    const failure = expect(inspectPdf(makePdf())).rejects.toThrow('safety limits')
    await vi.advanceTimersByTimeAsync(10001)
    await failure
    expect(destroyed).toBe(true)
    expect(vi.getTimerCount()).toBe(0)
  })
})
