import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../../lib/power-automate.js'
import { createOutlookMcpServer } from '../server.js'
import { attachmentIdSchema, MAX_ATTACHMENT_BYTES, readAttachmentInputSchema } from './schema.js'
import {
  attachmentResult,
  inspectAttachment,
  normalizeAttachments,
  readAttachment,
} from './service.js'
import { syntheticAttachment, syntheticGraphSizeAttachment } from './synthetic-fixtures.js'

const requestSchema = z.object({ operation: z.string(), requestId: z.string(), args: z.unknown() })
function flow(data: unknown) {
  const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
    const request = requestSchema.parse(
      JSON.parse(typeof init?.body === 'string' ? init.body : '{}'),
    )
    return Response.json({ ...request, ok: true, data })
  })
  return {
    client: new PowerAutomateClient({
      baseUrl: 'https://example.test/private',
      gatewayKey: 'synthetic',
      fetchFn,
    }),
    fetchFn,
  }
}
const target = { messageId: 'msg-1', attachmentId: 'att-1' }
async function connect(client: PowerAutomateClient) {
  const mcp = new Client({ name: 'test', version: '1' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), createOutlookMcpServer(client).connect(b)])
  return mcp
}

describe('bounded attachment service', () => {
  it.each(['pdf', 'docx', 'xlsx'] as const)(
    'inspects a real synthetic %s file without returning bytes',
    async (format) => {
      const file = syntheticAttachment(format)
      const { client } = flow(file)
      const result = await inspectAttachment(client, target)
      expect(result.source.format).toBe(format)
      expect(result.untrustedContent).toBe(true)
      expect(JSON.stringify(result)).not.toContain(file.contentBytes)
      expect(result.structure.format).toBe(format)
    },
  )
  it('reads PDF page text with provenance', async () => {
    const { client } = flow(syntheticAttachment('pdf'))
    const result = await readAttachment(
      client,
      readAttachmentInputSchema.parse({
        ...target,
        selection: { format: 'pdf', pageStart: 1, pageEnd: 1 },
      }),
    )
    expect(result.source).toMatchObject(target)
    expect(result.data).toMatchObject({
      format: 'pdf',
      pages: [{ page: 1, text: expect.stringContaining('Synthetic PDF page') }],
    })
  })
  it('reads a bounded Word text range with source offsets', async () => {
    const { client } = flow(syntheticAttachment('docx'))
    const result = await readAttachment(
      client,
      readAttachmentInputSchema.parse({
        ...target,
        selection: { format: 'docx', offset: 0, length: 8 },
      }),
    )
    expect(result.data).toMatchObject({
      format: 'docx',
      text: 'Overview',
      sourceStart: 0,
      sourceEnd: 8,
      nextOffset: 8,
    })
  })
  it('inspects and reads raw 889-byte DOCX with Graph size 1223 independently', async () => {
    const file = syntheticGraphSizeAttachment()
    expect(atob(file.contentBytes)).toHaveLength(889)
    expect(file.size).toBe(1223)
    const { client, fetchFn } = flow(file)
    const inspection = await inspectAttachment(client, target)
    expect(inspection.source.size).toBe(1223)
    expect(inspection.structure.format).toBe('docx')
    const read = await readAttachment(
      client,
      readAttachmentInputSchema.parse({
        ...target,
        selection: { format: 'docx', offset: 0, length: 100 },
      }),
    )
    expect(read.source.size).toBe(1223)
    expect(read.data).toMatchObject({
      format: 'docx',
      text: expect.stringContaining('Synthetic document body'),
    })
    expect(fetchFn).toHaveBeenCalledTimes(2)
    for (const result of [inspection, read]) {
      expect(JSON.stringify(result)).not.toContain('contentBytes')
      expect(JSON.stringify(result)).not.toContain(file.contentBytes)
    }
  })
  it('reads exact cells and cached values without formulas', async () => {
    const { client } = flow(syntheticAttachment('xlsx'))
    const result = await readAttachment(
      client,
      readAttachmentInputSchema.parse({
        ...target,
        selection: { format: 'xlsx', sheet: 'Budget', range: 'A1:B2' },
      }),
    )
    expect(result.data).toMatchObject({
      format: 'xlsx',
      sheet: 'Budget',
      cells: expect.arrayContaining([
        expect.objectContaining({ address: 'A1', value: 'Synthetic' }),
        expect.objectContaining({ address: 'B2', value: '42', hasFormula: true }),
      ]),
    })
    expect(JSON.stringify(result)).not.toContain('SUM(')
  })
  it('paginates metadata only and never exposes or follows nextLink', async () => {
    const file = syntheticAttachment('pdf')
    const { client, fetchFn } = flow({
      value: [file],
      '@odata.nextLink': 'https://evil.test/secret',
    })
    const mcp = await connect(client)
    const result = await mcp.callTool({
      name: 'outlook_list_attachments',
      arguments: { messageId: 'msg-1', limit: 1, offset: 3 },
    })
    expect(result.structuredContent).toMatchObject({
      hasMore: true,
      nextOffset: 4,
      attachments: [{ attachmentId: 'att-1', readable: true }],
    })
    expect(fetchFn.mock.calls[0]?.[1]?.body).toContain('"skip":3')
    expect(fetchFn).toHaveBeenCalledOnce()
    expect(JSON.stringify(result)).not.toContain('evil.test')
    expect(JSON.stringify(result)).not.toContain('contentBytes')
  })
  it('treats the flow empty nextLink placeholder as the final page', () => {
    expect(
      normalizeAttachments({ value: [], '@odata.nextLink': '' }, 'msg-1', 20, 0),
    ).toMatchObject({ hasMore: false, nextOffset: null })
  })
  it('rejects an empty continuation and a page larger than requested', () => {
    expect(() =>
      normalizeAttachments({ value: [], '@odata.nextLink': 'anything' }, 'msg-1', 10, 0),
    ).toThrow('pagination')
    expect(() =>
      normalizeAttachments({ value: [syntheticAttachment('pdf')] }, 'msg-1', 0, 0),
    ).toThrow('list response')
  })
  it.each([
    { '@odata.type': '#microsoft.graph.referenceAttachment' },
    { '@odata.type': '#microsoft.graph.itemAttachment' },
    { size: MAX_ATTACHMENT_BYTES + 1 },
    { contentBytes: 'not base64 !' },
    { id: 'other-id' },
    { name: 'legacy.xls' },
    { contentType: 'image/png' },
  ])('rejects unsafe metadata or encoding %j', async (change) => {
    const { client } = flow({ ...syntheticAttachment('pdf'), ...change })
    await expect(inspectAttachment(client, target)).rejects.toThrow()
  })
  it('validates size after base64 padding and rejects excessive actual bytes', async () => {
    const contentBytes = btoa('a'.repeat(MAX_ATTACHMENT_BYTES + 1))
    const { client } = flow({ ...syntheticAttachment('pdf'), size: 1, contentBytes })
    await expect(inspectAttachment(client, target)).rejects.toThrow('size limit')
  })
  it('rejects unsupported/malformed files without exposing contents through MCP errors', async () => {
    const secret = 'SYNTHETIC_PRIVATE_TEXT'
    const { client } = flow({ ...syntheticAttachment('pdf'), contentBytes: btoa(secret) })
    const mcp = await connect(client)
    const result = await mcp.request({
      method: 'tools/call',
      params: { name: 'outlook_inspect_attachment', arguments: target },
    })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).not.toContain(secret)
    expect(JSON.stringify(result)).not.toContain('example.test')
  })
  it.each([
    { format: 'docx', sectionId: 'SYNTHETIC_PRIVATE_SECTION', offset: 0, length: 8 },
    { format: 'docx', offset: 1000, length: 8 },
    { format: 'xlsx', sheet: 'SYNTHETIC_PRIVATE_SHEET', range: 'A1' },
    { format: 'xlsx', sheet: 'Budget', range: 'B2:A1' },
    { format: 'xlsx', sheet: 'Budget', range: 'A1:A501' },
    { format: 'pdf', pageStart: 1, pageEnd: 2 },
  ] as const)('returns a sanitized selector error through MCP for %j', async (selection) => {
    const file = syntheticAttachment(selection.format)
    const { client } = flow(file)
    const mcp = await connect(client)
    const result = await mcp.request({
      method: 'tools/call',
      params: { name: 'outlook_read_attachment', arguments: { ...target, selection } },
    })
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([
      {
        type: 'text',
        text:
          selection.format === 'pdf'
            ? 'The requested PDF page range or character limit is invalid.'
            : 'The requested document section or cell range is invalid.',
      },
    ])
    expect(result.structuredContent).toBeUndefined()
    expect(JSON.stringify(result)).not.toContain(file.contentBytes)
    expect(JSON.stringify(result)).not.toContain('SYNTHETIC_PRIVATE')
    expect(JSON.stringify(result)).not.toContain('example.test')
  })
  it('rejects format mismatch and unbounded selectors', async () => {
    const { client } = flow(syntheticAttachment('docx'))
    await expect(
      readAttachment(
        client,
        readAttachmentInputSchema.parse({
          ...target,
          selection: { format: 'pdf', pageStart: 1, pageEnd: 1 },
        }),
      ),
    ).rejects.toThrow('does not match')
    expect(
      readAttachmentInputSchema.safeParse({
        ...target,
        selection: { format: 'pdf', pageStart: 1, pageEnd: 11 },
      }).success,
    ).toBe(false)
    expect(
      readAttachmentInputSchema.safeParse({
        ...target,
        selection: { format: 'docx', length: 20001 },
      }).success,
    ).toBe(false)
  })
  it('validates IDs and encoded JSON byte limits', () => {
    for (const id of ['.', '..', '', 'has space', 'with\ncontrol'])
      expect(attachmentIdSchema.safeParse(id).success).toBe(false)
    expect(attachmentIdSchema.safeParse('AAMk+/=').success).toBe(true)
    expect(() => attachmentResult({ text: '日'.repeat(50000) })).toThrow('output limit')
  })
})
