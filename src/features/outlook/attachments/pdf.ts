import { getResolvedPDFJS } from 'unpdf'

import { getPdfDiagnosticCode, PdfParseError } from './pdf-errors.js'
import type { PdfDiagnosticCode } from './pdf-errors.js'

const MAX_RAW_BYTES = 4 * 1024 * 1024
const MAX_PAGES = 200
const MAX_READ_PAGES = 10
const MAX_CHARACTERS = 20000
const TIMEOUT_MS = 10000

/** A fixed selector message, never raw PDF.js errors or document contents. */
export class PdfRangeError extends Error {
  constructor() {
    super('The requested PDF page range or character limit is invalid.')
    this.name = 'PdfRangeError'
  }
}
function invalid(code: PdfDiagnosticCode): never {
  throw new PdfParseError(code)
}

class NoExternalData {
  fetch(): Promise<never> {
    return Promise.reject(new Error('External PDF resources are disabled'))
  }
}

type PdfModule = Awaited<ReturnType<typeof getResolvedPDFJS>>
type PdfDocument = Awaited<ReturnType<PdfModule['getDocument']>['promise']>

function classifyPdfFailure(error: unknown, fallback: PdfDiagnosticCode): PdfDiagnosticCode {
  const issued = getPdfDiagnosticCode(error)
  if (issued) return issued
  try {
    if (
      typeof error === 'object' &&
      error !== null &&
      'name' in error &&
      error.name === 'PasswordException'
    )
      return 'PDF_PASSWORD'
  } catch {
    // Unknown exception objects may have hostile accessors; retain the safe stage.
  }
  return fallback
}

async function withPdf<T>(
  bytes: Uint8Array,
  operation: (pdf: PdfDocument) => Promise<T>,
): Promise<T> {
  if (!bytes.length || bytes.length > MAX_RAW_BYTES) invalid('PDF_RAW_SIZE')

  let loading: ReturnType<PdfModule['getDocument']> | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  let stage: PdfDiagnosticCode = 'PDF_INITIALIZATION'
  try {
    const pdfjs = await getResolvedPDFJS()
    const parameters = {
      data: bytes.slice(),
      verbosity: 0,
      stopAtErrors: true,
      disableFontFace: true,
      // Defense in depth for PDF.js versions that support runtime codegen.
      isEvalSupported: false,
      useSystemFonts: true,
      useWorkerFetch: false,
      BinaryDataFactory: NoExternalData,
      useWasm: false,
      enableXfa: false,
      maxImageSize: 0,
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      disableAutoFetch: true,
      disableRange: true,
      disableStream: true,
      pdfBug: false,
    }
    stage = 'PDF_CREATE_DOCUMENT'
    loading = pdfjs.getDocument(parameters)
    const work = async () => {
      stage = 'PDF_LOAD'
      const pdf = await loading?.promise
      if (!pdf || !Number.isInteger(pdf.numPages) || pdf.numPages < 1) invalid('PDF_LOAD')
      if (pdf.numPages > MAX_PAGES) invalid('PDF_PAGE_LIMIT')
      stage = 'PDF_EXTRACTION'
      return await operation(pdf)
    }
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new PdfParseError('PDF_TIMEOUT')), TIMEOUT_MS)
    })
    return await Promise.race([work(), timeout])
  } catch (error) {
    if (error instanceof PdfRangeError) throw error
    return invalid(classifyPdfFailure(error, stage))
  } finally {
    if (timer !== undefined) clearTimeout(timer)
    // Destroy the loading task even if loading, parsing or extraction failed.
    if (loading) await loading.destroy().catch(() => undefined)
  }
}

export async function inspectPdf(bytes: Uint8Array): Promise<{ pageCount: number }> {
  return await withPdf(bytes, async (pdf) => ({ pageCount: pdf.numPages }))
}

export interface ReadPdfOptions {
  pageStart?: number
  pageEnd?: number
  maxCharacters?: number
}
export interface PdfPageText {
  page: number
  text: string
  truncated: boolean
}
export interface PdfText {
  pageCount: number
  pageStart: number
  pageEnd: number
  pages: PdfPageText[]
  truncated: boolean
}

export async function readPdf(bytes: Uint8Array, options: ReadPdfOptions = {}): Promise<PdfText> {
  const pageStart = options.pageStart ?? 1
  const maxCharacters = options.maxCharacters ?? MAX_CHARACTERS
  if (
    !Number.isInteger(pageStart) ||
    pageStart < 1 ||
    !Number.isInteger(maxCharacters) ||
    maxCharacters < 1 ||
    maxCharacters > MAX_CHARACTERS
  )
    throw new PdfRangeError()
  return await withPdf(bytes, async (pdf) => {
    const pageEnd = options.pageEnd ?? Math.min(pdf.numPages, pageStart + MAX_READ_PAGES - 1)
    if (
      !Number.isInteger(pageEnd) ||
      pageEnd < pageStart ||
      pageEnd > pdf.numPages ||
      pageEnd - pageStart + 1 > MAX_READ_PAGES
    )
      throw new PdfRangeError()
    const pages: PdfPageText[] = []
    let remaining = maxCharacters
    for (let number = pageStart; number <= pageEnd; number++) {
      if (remaining === 0) {
        pages.push({ page: number, text: '', truncated: true })
        continue
      }
      const page = await pdf.getPage(number)
      let text = ''
      let truncated = false
      const reader = page.streamTextContent().getReader()
      try {
        while (true) {
          const chunk: { done: boolean; value?: unknown } = await reader.read()
          if (chunk.done) break
          if (truncated) continue
          const value = chunk.value
          if (
            typeof value !== 'object' ||
            value === null ||
            !('items' in value) ||
            !Array.isArray(value.items)
          )
            invalid('PDF_TEXT_CHUNK')
          for (const item of value.items) {
            if (
              typeof item !== 'object' ||
              item === null ||
              !('str' in item) ||
              typeof item.str !== 'string'
            )
              continue
            const piece = item.str + ('hasEOL' in item && item.hasEOL ? '\n' : '')
            if (piece.length > remaining) {
              text += piece.slice(0, remaining)
              remaining = 0
              truncated = true
              break
            }
            text += piece
            remaining -= piece.length
          }
        }
      } finally {
        reader.releaseLock()
        page.cleanup()
      }
      pages.push({ page: number, text, truncated })
    }
    return {
      pageCount: pdf.numPages,
      pageStart,
      pageEnd,
      pages,
      truncated: pages.some((page) => page.truncated),
    }
  })
}
