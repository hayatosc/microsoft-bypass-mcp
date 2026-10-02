/** Fixed diagnostics only. Never add document values or original exceptions here. */
export const PDF_DIAGNOSTICS = Object.freeze({
  PDF_CREATE_DOCUMENT: 'PDF.js loading task creation failed',
  PDF_EXTRACTION: 'PDF.js inspection or text extraction failed',
  PDF_INITIALIZATION: 'PDF.js module initialization failed',
  PDF_LOAD: 'PDF.js document load failed',
  PDF_PAGE_LIMIT: 'Page count exceeded',
  PDF_PASSWORD: 'PDF.js reported an encrypted or password-protected document',
  PDF_RAW_SIZE: 'Empty PDF or raw byte limit exceeded',
  PDF_TEXT_CHUNK: 'PDF.js returned an invalid text chunk',
  PDF_TIMEOUT: 'PDF parse or extraction wall-time timeout',
  PDF_UNKNOWN: 'Unrecognized PDF diagnostic',
})
export type PdfDiagnosticCode = keyof typeof PDF_DIAGNOSTICS
const issuedCodes = new WeakMap<object, PdfDiagnosticCode>()

export class PdfParseError extends Error {
  constructor(code: PdfDiagnosticCode) {
    // Runtime validation also protects callers outside TypeScript.
    const safe =
      typeof code === 'string' && Object.hasOwn(PDF_DIAGNOSTICS, code) ? code : 'PDF_UNKNOWN'
    super(`PDF is invalid, unsupported, encrypted, or exceeds safety limits [${safe}]`)
    this.name = 'PdfParseError'
    issuedCodes.set(this, safe)
  }
}

/** Provenance is private: mutating error.message/code cannot change the public code. */
export function getPdfDiagnosticCode(error: unknown): PdfDiagnosticCode | undefined {
  return typeof error === 'object' && error !== null ? issuedCodes.get(error) : undefined
}
