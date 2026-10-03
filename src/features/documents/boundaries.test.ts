import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { oneDriveFileIdSchema } from '../onedrive/schema.js'
import { normalizeOneDriveList, normalizeOneDriveMetadata } from '../onedrive/service.js'
import {
  attachmentIdSchema,
  attachmentTargetSchema,
  listAttachmentsInputSchema,
  MAX_OUTPUT_CHARACTERS,
  readAttachmentInputSchema,
} from '../outlook/attachments/schema.js'
import { normalizeAttachments } from '../outlook/attachments/service.js'
import { syntheticAttachment, syntheticXlsx } from '../outlook/attachments/synthetic-fixtures.js'
import { createOutlookMcpServer } from '../outlook/server.js'
import { boundedToolResult, DocumentSourceError, formatOf, readDocument } from './source.js'

const formats = [
  ['pdf', 'application/pdf'],
  ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  ['xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
] as const
const metadata = {
  Id: 'file-1',
  Name: 'sample.pdf',
  Size: 3,
  MediaType: 'application/pdf',
  IsFolder: false,
}
const target = { messageId: 'msg-1', attachmentId: 'att-1' }
const selection = { format: 'xlsx', sheet: 'Budget', range: 'A1' } as const
const clients: Client[] = []
afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()))
  vi.restoreAllMocks()
})
async function flow(data: unknown) {
  const requestSchema = z.object({ operation: z.string(), requestId: z.string() })
  const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
    const request = requestSchema.parse(
      JSON.parse(typeof init?.body === 'string' ? init.body : '{}'),
    )
    return Response.json({ ...request, ok: true, data })
  })
  const client = new PowerAutomateClient({
    baseUrl: 'https://example.test/synthetic-flow',
    gatewayKey: 'synthetic',
    fetchFn,
  })
  const mcp = new Client({ name: 'document-boundary-tests', version: '1' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), createOutlookMcpServer(client).connect(b)])
  clients.push(mcp)
  return { mcp, fetchFn }
}
function call(mcp: Client, name: string, args: Record<string, unknown>) {
  return mcp.request({ method: 'tools/call', params: { name, arguments: args } })
}
function workbook(rows: string) {
  const parts = unzipSync(syntheticXlsx())
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
  )
  return { source: { format: 'xlsx' as const }, bytes: zipSync(parts) }
}
function cell(data: string, attributes = '') {
  return workbook(`<row r="1"><c r="A1" ${attributes}>${data}</c></row>`)
}

// Each case runs in the repository's official Cloudflare Workers Vitest integration.
describe('document extension boundary', () => {
  it.each(formats)('requires the dot in the %s extension', (format, media) => {
    for (const contentType of [media, '', 'application/octet-stream']) {
      expect(formatOf(format, contentType)).toBeNull()
      expect(formatOf(`sample.${format}`, contentType)).toBe(format)
      expect(formatOf(`sample.${format.toUpperCase()}`, contentType)).toBe(format)
    }
    expect(formatOf(`sample.${format}`, 'image/png')).toBeNull()
  })

  it.each(formats)('marks dotless %s metadata unreadable in both providers', (format, media) => {
    const native = { ...metadata, Name: format, MediaType: media }
    expect(normalizeOneDriveMetadata(native)).toMatchObject({
      supportedFormat: null,
      readable: false,
    })
    expect(normalizeOneDriveList({ value: [native], truncated: false }, 1).files[0]).toMatchObject({
      supportedFormat: null,
      readable: false,
    })
    expect(
      normalizeAttachments(
        { value: [{ ...syntheticAttachment(format), name: format, contentType: media }] },
        target.messageId,
        1,
        0,
      ).attachments[0],
    ).toMatchObject({ supportedFormat: null, readable: false })
  })
})

describe('attachment NEL input boundary', () => {
  const invalidId = 'x\u0085y'
  it('rejects NEL in the ID, target, list and read schemas while preserving opaque IDs', () => {
    expect(attachmentIdSchema.safeParse(invalidId).success).toBe(false)
    for (const field of ['messageId', 'attachmentId']) {
      const invalidTarget = { ...target, [field]: invalidId }
      expect(attachmentTargetSchema.safeParse(invalidTarget).success).toBe(false)
      expect(readAttachmentInputSchema.safeParse({ ...invalidTarget, selection }).success).toBe(
        false,
      )
    }
    expect(listAttachmentsInputSchema.safeParse({ messageId: invalidId }).success).toBe(false)
    expect(attachmentIdSchema.safeParse('AAMk+/=').success).toBe(true)
  })

  it.each([
    ['outlook_list_attachments', { messageId: invalidId, limit: 1, offset: 0 }],
    ['outlook_inspect_attachment', { ...target, messageId: invalidId }],
    ['outlook_inspect_attachment', { ...target, attachmentId: invalidId }],
    ['outlook_read_attachment', { ...target, messageId: invalidId, selection }],
    ['outlook_read_attachment', { ...target, attachmentId: invalidId, selection }],
  ])('rejects NEL before calling the flow via %s with %j', async (name, args) => {
    const { mcp, fetchFn } = await flow({})
    const result = await call(mcp, name, args)
    expect(result.isError).toBe(true)
    expect(fetchFn).not.toHaveBeenCalled()
  })
})

describe('OneDrive native output ID boundary', () => {
  it.each(['bad id', '.', '..', 'x\u0085y', 'x\ufeffy', 'x\ny'])(
    'rejects native ID %j rather than returning an unusable follow-up ID',
    (Id) => {
      expect(oneDriveFileIdSchema.safeParse(Id).success).toBe(false)
      expect(() =>
        normalizeOneDriveList({ value: [{ ...metadata, Id }], truncated: false }, 1),
      ).toThrow('Invalid OneDrive metadata response')
      expect(() => normalizeOneDriveMetadata({ ...metadata, Id })).toThrow(
        'Invalid OneDrive metadata response',
      )
    },
  )

  it('preserves a valid opaque native ID', () => {
    const Id = 'AAMk+/='
    expect(oneDriveFileIdSchema.safeParse(Id).success).toBe(true)
    expect(normalizeOneDriveList([{ ...metadata, Id }], 1).files[0]?.fileId).toBe(Id)
  })

  it('returns a fixed search error without the malformed ID or upstream URL', async () => {
    const canary = 'SYNTHETIC_INVALID_ID with space'
    const { mcp, fetchFn } = await flow({
      value: [{ ...metadata, Id: canary }],
      nextLink: 'https://example.test/private-next-link',
    })
    const result = await call(mcp, 'onedrive_search_files', { query: 'synthetic', limit: 1 })
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([{ type: 'text', text: 'Invalid OneDrive metadata response' }])
    expect(result.structuredContent).toBeUndefined()
    expect(JSON.stringify(result)).not.toContain(canary)
    expect(JSON.stringify(result)).not.toContain('example.test')
    expect(fetchFn).toHaveBeenCalledOnce()
  })
})

describe('XLSX selected text budget', () => {
  it.each([
    ['boolean raw text', 't="b"', `<v>${' '.repeat(MAX_OUTPUT_CHARACTERS)}1</v>`],
    [
      'missing formula cache raw text',
      '',
      `<f>1+1</f><v>${' '.repeat(MAX_OUTPUT_CHARACTERS + 1)}</v>`,
    ],
  ])(
    'rejects oversized %s through the shared document adapter',
    async (_label, attributes, data) => {
      await expect(readDocument(cell(data, attributes), selection)).rejects.toThrow(
        new DocumentSourceError('Selected cells exceed the text limit; request a smaller range'),
      )
    },
  )

  it.each([
    ['string', 't="str"', 'a'.repeat(MAX_OUTPUT_CHARACTERS)],
    ['numeric string', '', '1'.repeat(MAX_OUTPUT_CHARACTERS)],
    ['boolean', 't="b"', ' '.repeat(MAX_OUTPUT_CHARACTERS - 1) + '1'],
  ])(
    'preserves an exact-budget %s without counting identical value/rawValue twice',
    async (_label, attributes, raw) => {
      const result = await readDocument(cell(`<v>${raw}</v>`, attributes), selection)
      expect(result.data).toMatchObject({
        format: 'xlsx',
        cells: [{ value: attributes === 't="b"' ? true : raw, rawValue: raw }],
      })
    },
  )

  it('counts distinct decoded and raw strings as well as UTF-16 code units', async () => {
    // OOXML escapes preserve a distinct raw string; emoji use two UTF-16 units.
    const raw = '_x0041_'.repeat(2499) + '😀'.repeat(2)
    const result = await readDocument(cell(`<v>${raw}</v>`, 't="str"'), selection)
    expect(result.data).toMatchObject({
      cells: [{ value: 'A'.repeat(2499) + '😀'.repeat(2), rawValue: raw }],
    })
    const over = '_x0041_'.repeat(2499) + '😀'.repeat(3)
    await expect(readDocument(cell(`<v>${over}</v>`, 't="str"'), selection)).rejects.toThrow(
      'text limit',
    )
  })

  it('preserves boolean, precise numeric, blank and missing-cache values within budget', async () => {
    const source = workbook(
      '<row r="1"><c r="A1" t="b"><v>0</v></c><c r="B1"><v>9007199254740993</v></c><c r="C1"><f>1+1</f></c><c r="D1" t="str"><f>""</f><v/></c></row>',
    )
    const result = await readDocument(source, { ...selection, range: 'A1:E1' })
    expect(result.data).toMatchObject({
      cells: [
        { type: 'boolean', value: false, rawValue: '0' },
        { type: 'number', value: '9007199254740993', rawValue: '9007199254740993' },
        { type: 'blank', value: null, rawValue: null, hasFormula: true, cachedValueMissing: true },
        { type: 'string', value: '', rawValue: '', hasFormula: true, cachedValueMissing: false },
        { type: 'blank', value: null, rawValue: null, present: false },
      ],
    })
  })

  it('keeps the independent 128 KiB JSON guard even within the selected text budget', async () => {
    const rows = Array.from(
      { length: 500 },
      (_, index) =>
        `<row r="${index + 1}"><c r="A${index + 1}" t="str"><v>${'日'.repeat(40)}</v></c></row>`,
    ).join('')
    const result = await readDocument(workbook(rows), { ...selection, range: 'A1:A500' })
    expect(() => boundedToolResult(result)).toThrow('output limit')
  })

  it('returns a safe MCP text-limit error without raw document values', async () => {
    const raw = ' '.repeat(MAX_OUTPUT_CHARACTERS) + '1'
    const document = cell(`<v>${raw}</v>`, 't="b"')
    let binary = ''
    for (const byte of document.bytes) binary += String.fromCharCode(byte)
    const file = {
      ...syntheticAttachment('xlsx'),
      size: document.bytes.length,
      contentBytes: btoa(binary),
    }
    const { mcp, fetchFn } = await flow(file)
    const result = await call(mcp, 'outlook_read_attachment', { ...target, selection })
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([
      { type: 'text', text: 'Selected cells exceed the text limit; request a smaller range' },
    ])
    expect(result.structuredContent).toBeUndefined()
    expect(JSON.stringify(result)).not.toContain(raw)
    expect(JSON.stringify(result)).not.toContain(file.contentBytes)
    expect(JSON.stringify(result)).not.toContain('example.test')
    expect(fetchFn).toHaveBeenCalledOnce()
  })
})
