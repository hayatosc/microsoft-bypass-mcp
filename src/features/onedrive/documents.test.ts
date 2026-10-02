import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { MAX_ATTACHMENT_BYTES } from '../outlook/attachments/schema.js'
import { syntheticAttachment } from '../outlook/attachments/synthetic-fixtures.js'
import { createOutlookMcpServer } from '../outlook/server.js'
import { oneDriveInspectOutputSchema, oneDriveReadOutputSchema } from './schema.js'

const CANARY = 'SYNTHETIC_PRIVATE_ONEDRIVE_CANARY'
const URL_CANARY = `https://example.test/download/${CANARY}`
const requestSchema = z.object({
  operation: z.string(),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})
type Request = z.infer<typeof requestSchema>
type Format = 'pdf' | 'docx' | 'xlsx'
const clients: Client[] = []
afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()))
  vi.restoreAllMocks()
})
function nativeFixture(format: Format = 'pdf') {
  const attachment = syntheticAttachment(format)
  const metadata = {
    Id: 'file-1',
    Name: attachment.name,
    NameNoExt: 'synthetic',
    DisplayName: attachment.name,
    Size: attachment.size,
    MediaType: attachment.contentType,
    IsFolder: false,
    LastModified: '2026-01-01T12:00:00Z',
    ETag: 'synthetic-etag',
    Path: URL_CANARY,
    FileLocator: { downloadUrl: URL_CANARY },
    WebUrl: URL_CANARY,
    '@microsoft.graph.downloadUrl': URL_CANARY,
  }
  return { attachment, metadata, content: { metadata, contentBytes: attachment.contentBytes } }
}
async function flow(respond: (request: Request) => unknown) {
  const records: Request[] = []
  const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
    const request = requestSchema.parse(
      JSON.parse(typeof init?.body === 'string' ? init.body : '{}'),
    )
    records.push(request)
    return Response.json({
      ok: true,
      requestId: request.requestId,
      operation: request.operation,
      data: respond(request),
    })
  })
  const client = new PowerAutomateClient({ baseUrl: URL_CANARY, gatewayKey: CANARY, fetchFn })
  const mcp = new Client({ name: 'native-onedrive-tests', version: '1' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), createOutlookMcpServer(client).connect(b)])
  clients.push(mcp)
  return { mcp, records, fetchFn }
}
function call(mcp: Client, name: string, args: Record<string, unknown>) {
  return mcp.request({ method: 'tools/call', params: { name, arguments: args } })
}
function selection(format: Format) {
  if (format === 'pdf') return { format, pageStart: 1, pageEnd: 1, maxCharacters: 1000 }
  if (format === 'docx') return { format, offset: 0, length: 16 }
  return { format, sheet: 'Budget', range: 'A1:B2' }
}
function expectSanitized(result: unknown, base64?: string) {
  const serialized = JSON.stringify(result)
  expect(serialized).not.toContain(CANARY)
  expect(serialized).not.toContain('contentBytes')
  expect(serialized).not.toContain('downloadUrl')
  expect(serialized).not.toContain('FileLocator')
  if (base64 !== undefined) expect(serialized).not.toContain(base64)
}

describe('OneDrive native document adapter boundaries', () => {
  it('normalizes upper-case native metadata and excludes URL and locator fields', async () => {
    const { metadata, attachment } = nativeFixture()
    const { mcp, records } = await flow(() => metadata)
    const result = await call(mcp, 'onedrive_get_metadata', { fileId: 'file-1' })
    expect(result.structuredContent).toEqual({
      fileId: 'file-1',
      name: 'synthetic.pdf',
      size: attachment.size,
      contentType: 'application/octet-stream',
      isFolder: false,
      supportedFormat: 'pdf',
      readable: true,
      lastModifiedDateTime: metadata.LastModified,
      eTag: metadata.ETag,
      limitation: null,
    })
    expect(records.map(({ operation, args }) => ({ operation, args }))).toEqual([
      { operation: 'onedrive_get_metadata', args: { fileId: 'file-1' } },
    ])
    expectSanitized(result)
  })

  it.each([
    ['native continuation', { nextLink: URL_CANARY }, true, false],
    ['native truncation', { truncated: true }, true, true],
    ['both native flags', { nextLink: URL_CANARY, truncated: true }, true, true],
    ['empty nextLink', { nextLink: '', truncated: false }, false, false],
    ['null nextLink', { nextLink: null }, false, false],
    ['no continuation flags', {}, false, false],
  ] as const)(
    'preserves %s without inventing or exposing a cursor',
    async (_label, page, hasMore, truncated) => {
      const { metadata } = nativeFixture()
      const { mcp, fetchFn } = await flow(() => ({ value: [metadata], ...page }))
      const result = await call(mcp, 'onedrive_list_folder', { limit: 1 })
      expect(result.structuredContent).toMatchObject({
        folderId: null,
        files: [{ fileId: 'file-1' }],
        hasMore,
        truncated,
        nextCursor: null,
        incompleteReason: hasMore ? expect.any(String) : null,
      })
      expect(fetchFn).toHaveBeenCalledOnce()
      expectSanitized(result)
    },
  )

  it.each([
    ['lower-case Graph shape', { value: [{ id: 'file-1', name: 'report.pdf', size: 1 }] }],
    ['wrong native size type', { value: [{ ...nativeFixture().metadata, Size: '12' }] }],
    ['wrong native folder flag', { value: [{ ...nativeFixture().metadata, IsFolder: 'false' }] }],
    ['malformed page flag', { value: [], truncated: 'true' }],
  ])('rejects %s instead of silently assuming native metadata', async (_label, data) => {
    const { mcp } = await flow(() => data)
    const result = await call(mcp, 'onedrive_search_files', { query: CANARY })
    expect(result.isError).toBe(true)
    expect(result.structuredContent).toBeUndefined()
    expectSanitized(result)
  })

  it('rejects a native page larger than the requested limit', async () => {
    const { metadata } = nativeFixture()
    const { mcp } = await flow(() => [metadata, { ...metadata, Id: 'file-2' }])
    const result = await call(mcp, 'onedrive_search_files', { query: 'synthetic', limit: 1 })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('Invalid OneDrive page response')
  })

  it.each([
    ['folder', { IsFolder: true }],
    ['unsupported extension', { Name: 'unsupported.exe' }],
    ['inconsistent media type', { MediaType: 'image/png' }],
    ['empty file', { Size: 0 }],
    ['oversized file', { Size: MAX_ATTACHMENT_BYTES + 1 }],
    ['wrong identity', { Id: 'other-file' }],
  ] as const)(
    'gates %s metadata before fetching any content for inspect or read',
    async (_label, patch) => {
      const fixture = nativeFixture()
      const { mcp, records } = await flow(({ operation }) => {
        if (operation !== 'onedrive_get_metadata') throw new Error('Content must not be fetched')
        return { ...fixture.metadata, ...patch }
      })
      const inspect = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
      const read = await call(mcp, 'onedrive_read_file', {
        fileId: 'file-1',
        selection: selection('pdf'),
      })
      expect(inspect.isError).toBe(true)
      expect(read.isError).toBe(true)
      expect(records.map(({ operation }) => operation)).toEqual([
        'onedrive_get_metadata',
        'onedrive_get_metadata',
      ])
      expectSanitized(inspect)
      expectSanitized(read)
    },
  )

  it.each([
    ['identity', { Id: 'other-file' }],
    ['name', { Name: 'changed.pdf' }],
    ['size', { Size: nativeFixture().metadata.Size + 1 }],
    ['media type', { MediaType: 'application/pdf' }],
    ['folder type', { IsFolder: true }],
    ['ETag', { ETag: 'changed-etag' }],
    ['removed ETag', { ETag: null }],
    ['modified time', { LastModified: '2026-01-02T12:00:00Z' }],
    ['removed modified time', { LastModified: null }],
  ] as const)('rejects a %s race between metadata and content', async (_label, patch) => {
    const fixture = nativeFixture()
    const { mcp, records } = await flow(({ operation }) =>
      operation === 'onedrive_get_metadata'
        ? fixture.metadata
        : { ...fixture.content, metadata: { ...fixture.metadata, ...patch } },
    )
    const result = await call(mcp, 'onedrive_read_file', {
      fileId: 'file-1',
      selection: selection('pdf'),
    })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toMatch(/metadata (?:changed|ID does not match)/)
    expect(records.map(({ operation, args }) => ({ operation, args }))).toEqual([
      { operation: 'onedrive_get_metadata', args: { fileId: 'file-1' } },
      { operation: 'onedrive_get_content', args: { fileId: 'file-1' } },
    ])
    expectSanitized(result, fixture.attachment.contentBytes)
  })

  it.each(['!AAA', 'A===', 'AAAA=', 'AA A', 'AA-A'])(
    'rejects malformed native content base64 %s',
    async (contentBytes) => {
      const fixture = nativeFixture()
      const { mcp } = await flow(({ operation }) =>
        operation === 'onedrive_get_metadata'
          ? fixture.metadata
          : { metadata: fixture.metadata, contentBytes },
      )
      const result = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
      expect(result.isError).toBe(true)
      expect(JSON.stringify(result)).toContain('Invalid OneDrive content response')
      expectSanitized(result)
    },
  )

  it.each([-1, 1])(
    'enforces native exact size when decoded content differs by %i byte',
    async (delta) => {
      const fixture = nativeFixture()
      const metadata = { ...fixture.metadata, Size: fixture.metadata.Size + delta }
      const { mcp } = await flow(({ operation }) =>
        operation === 'onedrive_get_metadata'
          ? metadata
          : { metadata, contentBytes: fixture.attachment.contentBytes },
      )
      const result = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
      expect(result.isError).toBe(true)
      expect(JSON.stringify(result)).toContain('OneDrive content size does not match metadata')
      expectSanitized(result, fixture.attachment.contentBytes)
    },
  )

  it.each(['pdf', 'docx', 'xlsx'] as const)(
    'matches attachment %s parsing with distinct OneDrive provenance',
    async (format) => {
      const fixture = nativeFixture(format)
      const logs = vi.spyOn(console, 'log').mockImplementation(() => {})
      const { mcp, records, fetchFn } = await flow(({ operation }) => {
        if (operation === 'onedrive_get_metadata') return fixture.metadata
        if (operation === 'onedrive_get_content') return fixture.content
        if (operation === 'get_attachment') return fixture.attachment
        throw new Error(`Unexpected operation: ${operation}`)
      })
      const target = { messageId: 'message-1', attachmentId: 'att-1' }
      const nativeInspect = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
      const attachmentInspect = await call(mcp, 'outlook_inspect_attachment', target)
      const nativeRead = await call(mcp, 'onedrive_read_file', {
        fileId: 'file-1',
        selection: selection(format),
      })
      const attachmentRead = await call(mcp, 'outlook_read_attachment', {
        ...target,
        selection: selection(format),
      })
      const inspected = oneDriveInspectOutputSchema.parse(nativeInspect.structuredContent)
      const read = oneDriveReadOutputSchema.parse(nativeRead.structuredContent)
      expect(inspected.structure).toEqual(
        z.object({ structure: z.unknown() }).parse(attachmentInspect.structuredContent).structure,
      )
      expect(read.data).toEqual(
        z.object({ data: z.unknown() }).parse(attachmentRead.structuredContent).data,
      )
      for (const result of [inspected, read]) {
        expect(result.untrustedContent).toBe(true)
        expect(result.source).toEqual({
          provider: 'onedrive',
          fileId: 'file-1',
          name: fixture.metadata.Name,
          contentType: fixture.metadata.MediaType,
          size: fixture.metadata.Size,
          lastModifiedDateTime: fixture.metadata.LastModified,
          eTag: fixture.metadata.ETag,
          format,
        })
      }
      expect(
        z.object({ source: z.unknown() }).parse(attachmentRead.structuredContent).source,
      ).toMatchObject(target)
      expect(records.map(({ operation }) => operation)).toEqual([
        'onedrive_get_metadata',
        'onedrive_get_content',
        'get_attachment',
        'onedrive_get_metadata',
        'onedrive_get_content',
        'get_attachment',
      ])
      expect(fetchFn.mock.calls.every(([url]) => url === URL_CANARY)).toBe(true)
      for (const result of [nativeInspect, nativeRead, attachmentInspect, attachmentRead]) {
        expect(result.isError).not.toBe(true)
        expectSanitized(result, fixture.attachment.contentBytes)
      }
      expect(logs).toHaveBeenCalledTimes(6)
      for (const [entry] of logs.mock.calls) {
        const log = z.record(z.string(), z.unknown()).parse(JSON.parse(z.string().parse(entry)))
        expect(Object.keys(log).sort()).toEqual([
          'durationMs',
          'operation',
          'requestId',
          'status',
          'success',
          'type',
        ])
        expect(log.success).toBe(true)
      }
      expectSanitized(logs.mock.calls, fixture.attachment.contentBytes)
    },
  )

  it.each([
    ['docx', { format: 'docx', sectionId: CANARY, offset: 0, length: 8 }],
    ['xlsx', { format: 'xlsx', sheet: CANARY, range: 'A1' }],
    ['pdf', { format: 'pdf', pageStart: 2, pageEnd: 2 }],
  ] as const)(
    'sanitizes invalid %s selectors at the MCP boundary',
    async (format, invalidSelection) => {
      const fixture = nativeFixture(format)
      const { mcp } = await flow(({ operation }) =>
        operation === 'onedrive_get_metadata' ? fixture.metadata : fixture.content,
      )
      const result = await call(mcp, 'onedrive_read_file', {
        fileId: 'file-1',
        selection: invalidSelection,
      })
      expect(result.isError).toBe(true)
      expect(result.content).toEqual([
        {
          type: 'text',
          text:
            format === 'pdf'
              ? 'The requested PDF page range or character limit is invalid.'
              : 'The requested document section or cell range is invalid.',
        },
      ])
      expect(result.structuredContent).toBeUndefined()
      expectSanitized(result, fixture.attachment.contentBytes)
    },
  )

  it('does not leak malformed document text, base64, filenames or URLs through parser errors or logs', async () => {
    const logs = vi.spyOn(console, 'log').mockImplementation(() => {})
    const fixture = nativeFixture()
    const metadata = { ...fixture.metadata, Name: `${CANARY}.pdf`, Size: CANARY.length }
    const contentBytes = btoa(CANARY)
    const { mcp } = await flow(({ operation }) =>
      operation === 'onedrive_get_metadata' ? metadata : { metadata, contentBytes },
    )
    const result = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
    expect(result.isError).toBe(true)
    expect(result.structuredContent).toBeUndefined()
    expectSanitized(result, contentBytes)
    expectSanitized(logs.mock.calls, contentBytes)
  })

  it('rejects extra native content envelope fields without returning their values', async () => {
    const fixture = nativeFixture()
    const { mcp } = await flow(({ operation }) =>
      operation === 'onedrive_get_metadata'
        ? fixture.metadata
        : { ...fixture.content, downloadUrl: URL_CANARY },
    )
    const result = await call(mcp, 'onedrive_inspect_file', { fileId: 'file-1' })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(result)).toContain('Invalid OneDrive content response')
    expectSanitized(result, fixture.attachment.contentBytes)
  })
})
