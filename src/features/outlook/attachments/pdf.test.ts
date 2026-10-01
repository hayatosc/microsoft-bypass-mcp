import { getResolvedPDFJS } from 'unpdf'
import type * as Unpdf from 'unpdf'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('unpdf', async (importOriginal) => {
  const actual = await importOriginal<typeof Unpdf>()
  return { ...actual, getResolvedPDFJS: vi.fn(actual.getResolvedPDFJS) }
})

import { makePdf } from './fixtures-pdf.js'
import { inspectPdf, readPdf } from './pdf.js'

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
    await expect(readPdf(makePdf(), options)).rejects.toThrow('safety limits')
  })

  it('limits reads to ten pages and documents to 200 pages', async () => {
    const bytes = makePdf({ pages: 11 })
    expect((await readPdf(bytes)).pages).toHaveLength(10)
    await expect(readPdf(bytes, { pageEnd: 11 })).rejects.toThrow('safety limits')
    await expect(inspectPdf(makePdf({ pages: 201 }))).rejects.toThrow('safety limits')
  })

  it('rejects encrypted, active, incremental and object-stream forms before parsing', async () => {
    for (const trailer of [
      '/Encrypt 3 0 R',
      '/Prev 0',
      '/XRefStm 0',
      '/JS (secret)',
      '/JavaScript (secret)',
    ]) {
      await expect(inspectPdf(makePdf({ trailer }))).rejects.toThrow('safety limits')
    }
    for (const streamDictionary of [
      '/Type /ObjStm',
      '/Type /XRef',
      '/Filter /LZWDecode',
      '/Filter /DCTDecode',
      '/Filter 3 0 R',
      '/DecodeParms << /Predictor 12 >>',
      '/Length 3 0 R',
      '/F (https://example.invalid/secret)',
    ]) {
      await expect(inspectPdf(makePdf({ streamDictionary }))).rejects.toThrow('safety limits')
    }
  })

  it('rejects encoded-name bypasses, duplicate keys, invalid cross references and raw limits', async () => {
    await expect(inspectPdf(makePdf({ trailer: '/Encr#79pt 3 0 R' }))).rejects.toThrow(
      'safety limits',
    )
    await expect(inspectPdf(makePdf({ streamDictionary: '/L#65ngth 1' }))).rejects.toThrow(
      'safety limits',
    )
    await expect(
      inspectPdf(replace(makePdf(), '0000000009 00000 n', '0000000010 00000 n')),
    ).rejects.toThrow('safety limits')
    await expect(inspectPdf(new Uint8Array(4 * 1024 * 1024 + 1))).rejects.toThrow('safety limits')
    await expect(inspectPdf(encoder.encode('not a PDF secret'))).rejects.toThrow('safety limits')
  })

  it('rejects compressed stream bombs by actual decoded bytes, not PDF Length', async () => {
    const bytes = makePdf({ pages: 1, content: ' '.repeat(9 * 1024 * 1024), compressed: true })
    expect(bytes.length).toBeLessThan(20000)
    await expect(inspectPdf(bytes)).rejects.toThrow('safety limits')
  })

  it('rejects the aggregate stream budget and inline image codecs', async () => {
    await expect(
      inspectPdf(makePdf({ pages: 3, content: ' '.repeat(6 * 1024 * 1024), compressed: true })),
    ).rejects.toThrow('safety limits')
    await expect(
      inspectPdf(makePdf({ content: 'q BI /W 1 /H 1 /BPC 8 /CS /G ID x EI Q' })),
    ).rejects.toThrow('safety limits')
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
    await expect(readPdf(makePdf(), { pageEnd: 3 })).rejects.toThrow('safety limits')
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
        // Preserve a real task's lifecycle while simulating a stalled load.
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

  it('bounds CMap expansion independently of compressed stream sizes', async () => {
    const valid = makePdf({ content: '1 beginbfchar <0041> <0042> endbfchar', pages: 1 })
    expect(await inspectPdf(valid)).toEqual({ pageCount: 1 })
    const mappings = [
      '1 beginbfrange <00000000> <00ffffff> <0000> endbfrange',
      '2 beginbfrange <0000> <ffff> <0000> <0000> <ffff> <0000> endbfrange',
      `1 beginbfrange <0000> <ffff> <${'0041'.repeat(32)}> endbfrange`,
      '1 begincidrange <0000> <ffffff> 0 endcidrange',
      '1 beginbfchar <ffff0000> <0000> endbfchar',
    ]
    for (const content of mappings) {
      await expect(inspectPdf(makePdf({ pages: 1, content, compressed: true }))).rejects.toThrow(
        'safety limits',
      )
    }
  })

  it('bounds repeated content references and rejects Form/Type3 amplification', async () => {
    await expect(
      inspectPdf(
        makePdf({
          pages: 1,
          content: ' '.repeat(2 * 1024 * 1024),
          compressed: true,
          contentRepeats: 9,
        }),
      ),
    ).rejects.toThrow('safety limits')
    for (const streamDictionary of ['/Subtype /Form', '/Subtype /Type3', '/Type 3 0 R']) {
      await expect(inspectPdf(makePdf({ streamDictionary }))).rejects.toThrow('safety limits')
    }
  })

  it('bounds CID width ranges, sparse font indexes, and repeated mapping expansion before PDF.js', async () => {
    const pdfjs = vi.mocked(getResolvedPDFJS)
    pdfjs.mockClear()
    for (const streamDictionary of [
      '/W [0 1000000000 500]',
      '/W2 [0 1000000000 1 2 3]',
      '/W [0 65535 500 0 65535 500]',
      '/W [65535 [500 500]]',
      '/W2 [0 [1 2]]',
      '/FirstChar 1000000000',
      '/LastChar 1000000000',
      '/Differences [1000000000 /A]',
      '/Differences [65535 /A /B]',
    ]) {
      await expect(inspectPdf(makePdf({ streamDictionary }))).rejects.toThrow('safety limits')
    }
    expect(pdfjs).not.toHaveBeenCalled()
    expect(
      await inspectPdf(
        makePdf({ streamDictionary: '/W [0 255 500] /W2 [0 [1 2 3]] /Differences [65 /A /B]' }),
      ),
    ).toEqual({ pageCount: 2 })
  })
})
