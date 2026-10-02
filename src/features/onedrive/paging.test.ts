import { Client, InMemoryTransport } from '@modelcontextprotocol/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { createOutlookMcpServer } from '../outlook/server.js'
import {
  oneDriveListFolderInputSchema,
  oneDriveListFolderOutputSchema,
  oneDriveSearchInputSchema,
} from './schema.js'
import { listOneDriveFolder, normalizeOneDriveMetadata, searchOneDrive } from './service.js'

const CANARY = 'SYNTHETIC_PRIVATE_PAGING_CANARY'
const URL_CANARY = `https://outside.invalid/${CANARY}`
const FLOW_URL = 'https://example.test/fixed-flow'
const requestSchema = z.object({
  operation: z.string(),
  requestId: z.string().uuid(),
  args: z.record(z.string(), z.unknown()),
})
type Request = z.infer<typeof requestSchema>
const clients: Client[] = []
afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()))
  vi.restoreAllMocks()
})
function items(count: number) {
  return Array.from({ length: count }, (_, index) => ({
    Id: `owned-file-${index}`,
    Name: `report-${index}.pdf`,
    Size: index + 1,
    MediaType: 'application/pdf',
    IsFolder: false,
    LastModified: '2026-01-01T00:00:00Z',
    ETag: `etag-${index}`,
    Path: URL_CANARY,
    FileLocator: { downloadUrl: URL_CANARY },
    contentBytes: btoa(CANARY),
  }))
}
function flow(respond: (request: Request) => unknown) {
  const records: Request[] = []
  const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
    const request = requestSchema.parse(JSON.parse(z.string().parse(init?.body)))
    records.push(request)
    return Response.json({
      ok: true,
      requestId: request.requestId,
      operation: request.operation,
      data: respond(request),
    })
  })
  const client = new PowerAutomateClient({ baseUrl: FLOW_URL, gatewayKey: 'test', fetchFn })
  return { client, records, fetchFn }
}
async function connect(client: PowerAutomateClient) {
  const mcp = new Client({ name: 'onedrive-paging-tests', version: '1' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await Promise.all([mcp.connect(a), createOutlookMcpServer(client).connect(b)])
  clients.push(mcp)
  return mcp
}
function call(mcp: Client, args: Record<string, unknown>, name = 'onedrive_list_folder') {
  return mcp.request({ method: 'tools/call', params: { name, arguments: args } })
}
function cursorOf(page: { nextCursor: string | null }): string {
  if (page.nextCursor === null) throw new Error('Expected a continuation cursor')
  return page.nextCursor
}
function decode(cursor: string): Record<string, unknown> {
  const base64 = cursor.replace(/-/g, '+').replace(/_/g, '/')
  return z
    .record(z.string(), z.unknown())
    .parse(JSON.parse(atob(base64 + '='.repeat((4 - (base64.length % 4)) % 4))))
}
function encodeText(text: string) {
  return btoa(text).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}
function encode(payload: Record<string, unknown>) {
  return encodeText(JSON.stringify(payload))
}
function expectSanitized(result: unknown) {
  const text = JSON.stringify(result)
  for (const marker of [
    CANARY,
    'contentBytes',
    btoa(CANARY),
    'FileLocator',
    'downloadUrl',
    'nextLink',
  ])
    expect(text).not.toContain(marker)
}

describe('OneDrive stateless bounded folder pagination', () => {
  it.each([undefined, 'owned-folder'])(
    'pages the full 250-item window for scope %s',
    async (folderId) => {
      // The flow has already aggregated multiple actual native connector pages.
      const native = items(250)
      const { client, records, fetchFn } = flow(() => ({ value: native, truncated: false }))
      const mcp = await connect(client)
      const ids: string[] = []
      let cursor: string | undefined
      for (let index = 0; index < 5; index++) {
        const result = await call(mcp, { folderId, limit: 60, cursor })
        expect(result.isError).not.toBe(true)
        const page = oneDriveListFolderOutputSchema.parse(result.structuredContent)
        ids.push(...page.files.map((file) => file.fileId))
        expect(page.files).toHaveLength(index === 4 ? 10 : 60)
        expect(page.folderId).toBe(folderId ?? null)
        expect(page.hasMore).toBe(index < 4)
        expect(page.truncated).toBe(false)
        expect(page.incompleteReason).toBeNull()
        expectSanitized(result)
        cursor = page.nextCursor ?? undefined
      }
      expect(cursor).toBeUndefined()
      expect(ids).toEqual(native.map((item) => item.Id))
      expect(new Set(ids).size).toBe(250)
      expect(records.map(({ operation, args }) => ({ operation, args }))).toEqual(
        Array.from({ length: 5 }, () => ({
          operation: 'onedrive_list_folder',
          args: { ...(folderId === undefined ? {} : { folderId }), top: 1000 },
        })),
      )
      expect(
        fetchFn.mock.calls.every(([url, init]) => url === FLOW_URL && init?.redirect === 'manual'),
      ).toBe(true)
    },
  )

  it('binds only hashes and bounded offsets, fingerprints the entire normalized ordered window, and needs no stored session', async () => {
    const native = items(250)
    const first = flow(() => native)
    const page = await listOneDriveFolder(first.client, 'private-folder-id', 100)
    const cursor = cursorOf(page)
    const payload = decode(cursor)
    const digest = await crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(JSON.stringify(native.map(normalizeOneDriveMetadata))),
    )
    const hash = [...new Uint8Array(digest)]
      .map((byte) => byte.toString(16).padStart(2, '0'))
      .join('')
    expect(payload).toEqual({
      v: 1,
      operation: 'onedrive_list_folder',
      scope: expect.stringMatching(/^[a-f0-9]{64}$/),
      window: hash,
      offset: 100,
      limit: 100,
    })
    expect(cursor.length).toBeLessThanOrEqual(1024)
    expect(cursor).toMatch(/^[A-Za-z0-9_-]+$/)
    const decoded = JSON.stringify(payload)
    for (const marker of ['private-folder-id', 'owned-file-', 'report-', 'etag-', 'http', CANARY])
      expect(decoded).not.toContain(marker)
    // An entirely new client can resume; no server-side collection storage.
    const second = flow(() => ({
      value: native.map((item) => ({ ...item, Path: 'changed-raw-url' })),
    }))
    const resumed = await listOneDriveFolder(second.client, 'private-folder-id', 100, cursor)
    expect(resumed.files.map((file) => file.fileId)).toEqual(
      native.slice(100, 200).map((item) => item.Id),
    )
    expect(first.fetchFn).toHaveBeenCalledOnce()
    expect(second.fetchFn).toHaveBeenCalledOnce()
  })

  it.each([
    ['owned-folder', undefined, 10],
    [undefined, 'owned-folder', 10],
    ['owned-folder', 'different-folder', 10],
    ['owned-folder', 'owned-folder', 11],
  ] as const)(
    'rejects changed folder/limit scope before fetching: %s -> %s, %i',
    async (before, after, limit) => {
      const { client, fetchFn } = flow(() => items(25))
      const cursor = cursorOf(await listOneDriveFolder(client, before, 10))
      fetchFn.mockClear()
      await expect(listOneDriveFolder(client, after, limit, cursor)).rejects.toThrow(
        'cursor for this request',
      )
      expect(fetchFn).not.toHaveBeenCalled()
    },
  )

  it('rejects query changes and using a folder cursor for search at the closed MCP input boundary', async () => {
    const { client, fetchFn } = flow(() => items(25))
    const cursor = cursorOf(await listOneDriveFolder(client, undefined, 10))
    const mcp = await connect(client)
    fetchFn.mockClear()
    for (const query of ['original-query', 'changed-query']) {
      expect((await call(mcp, { cursor, query, limit: 10 })).isError).toBe(true)
      expect((await call(mcp, { cursor, query, limit: 10 }, 'onedrive_search_files')).isError).toBe(
        true,
      )
    }
    expect(fetchFn).not.toHaveBeenCalled()
  })

  it.each([
    { v: 2 },
    { operation: 'onedrive_search_files' },
    { scope: 'a'.repeat(64) },
    { scope: 'A'.repeat(64) },
    { scope: 'a'.repeat(64) + '\n' },
    { window: 'https://outside.invalid/' },
    { window: 'a'.repeat(64) + '\n' },
    { offset: 0 },
    { offset: -1 },
    { offset: 1000 },
    { offset: 1.5 },
    { offset: '1' },
    { limit: 0 },
    { limit: 101 },
    { limit: 11 },
    { url: URL_CANARY },
    { folderId: CANARY },
  ])('rejects invalid typed cursor payload %j before fetching', async (patch) => {
    const { client, fetchFn } = flow(() => items(25))
    const cursor = cursorOf(await listOneDriveFolder(client, undefined, 10))
    fetchFn.mockClear()
    await expect(
      listOneDriveFolder(client, undefined, 10, encode({ ...decode(cursor), ...patch })),
    ).rejects.toThrow('cursor')
    expect(fetchFn).not.toHaveBeenCalled()
  })

  it('rejects malformed, oversized and noncanonical cursor encodings before fetching', async () => {
    const { client, fetchFn } = flow(() => items(25))
    const cursor = cursorOf(await listOneDriveFolder(client, undefined, 10))
    const payload = decode(cursor)
    const text = JSON.stringify(payload)
    const invalid = [
      '',
      'A'.repeat(1025),
      'A',
      '_',
      '!!!',
      'é',
      cursor + '=',
      cursor + '\n',
      encodeText('not json'),
      encodeText('null'),
      encodeText('[]'),
      encode(Object.fromEntries(Object.entries(payload).reverse())),
      encodeText(JSON.stringify(payload, null, 2)),
      encodeText(text.replace('"v":1', '"v":1,"v":1')),
      encodeText(text.replace('"v":1', '"v":1e0')),
    ]
    // A different base64 pad-bit spelling must not alias the canonical token.
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'
    const last = cursor.at(-1)
    if (last !== undefined && cursor.length % 4 !== 0)
      invalid.push(cursor.slice(0, -1) + alphabet[alphabet.indexOf(last) + 1])
    fetchFn.mockClear()
    for (const token of invalid)
      await expect(listOneDriveFolder(client, undefined, 10, token)).rejects.toThrow(
        /cursor|request/,
      )
    expect(fetchFn).not.toHaveBeenCalled()
  })

  it.each([
    'reordered',
    'name',
    'size',
    'type',
    'folder',
    'etag',
    'modified',
    'removed',
    'added',
  ] as const)(
    'rejects a stale %s fingerprint, including changes outside the requested page',
    async (change) => {
      let native = items(250)
      const { client, fetchFn } = flow(() => ({ value: native }))
      const cursor = cursorOf(await listOneDriveFolder(client, undefined, 10))
      const last = native[249]
      if (last === undefined) throw new Error('Missing fixture')
      if (change === 'reordered') native = [...native].reverse()
      else if (change === 'removed') native = native.slice(0, -1)
      else if (change === 'added') native = items(251)
      else {
        const patch =
          change === 'name'
            ? { Name: 'renamed.pdf' }
            : change === 'size'
              ? { Size: 999 }
              : change === 'type'
                ? { MediaType: 'application/octet-stream' }
                : change === 'folder'
                  ? { IsFolder: true }
                  : change === 'etag'
                    ? { ETag: 'changed' }
                    : { LastModified: '2026-02-01T00:00:00Z' }
        native = [...native.slice(0, -1), { ...last, ...patch }]
      }
      await expect(listOneDriveFolder(client, undefined, 10, cursor)).rejects.toThrow(
        'restart from the first page',
      )
      expect(fetchFn).toHaveBeenCalledTimes(2)
    },
  )

  it('rejects offsets outside the known window after re-fetching, but bounded offsets are not authorization', async () => {
    const { client, fetchFn } = flow(() => items(25))
    const cursor = cursorOf(await listOneDriveFolder(client, undefined, 10))
    const payload = decode(cursor)
    for (const offset of [25, 999])
      await expect(
        listOneDriveFolder(client, undefined, 10, encode({ ...payload, offset })),
      ).rejects.toThrow('cursor offset')
    const page = await listOneDriveFolder(client, undefined, 10, encode({ ...payload, offset: 24 }))
    expect(page.files.map((file) => file.fileId)).toEqual(['owned-file-24'])
    expect(page.hasMore).toBe(false)
    expect(page.nextCursor).toBeNull()
    expect(fetchFn).toHaveBeenCalledTimes(4)
  })

  it.each([undefined, 'owned-folder'])(
    'rejects oversized raw arrays and envelopes for %s',
    async (folderId) => {
      for (const raw of [items(1001), { value: items(1001) }]) {
        const { client, fetchFn } = flow(() => raw)
        await expect(listOneDriveFolder(client, folderId, 100)).rejects.toThrow('metadata')
        expect(fetchFn).toHaveBeenCalledOnce()
      }
    },
  )

  it('allows offset 999 only inside a known 1000-item window, without implying complete enumeration', async () => {
    const { client } = flow(() => items(1000))
    const first = await listOneDriveFolder(client, undefined, 1)
    const payload = decode(cursorOf(first))
    const last = await listOneDriveFolder(client, undefined, 1, encode({ ...payload, offset: 999 }))
    expect(last.files.map((file) => file.fileId)).toEqual(['owned-file-999'])
    expect(last.hasMore).toBe(true)
    expect(last.nextCursor).toBeNull()
    expect(last.truncated).toBe(true)
    expect(last.incompleteReason).toEqual(expect.any(String))
  })

  it('recomputes upstream incompleteness independently from an unchanged metadata fingerprint', async () => {
    let truncated = false
    const { client } = flow(() => ({ value: items(3), truncated }))
    const first = await listOneDriveFolder(client, undefined, 2)
    truncated = true
    const last = await listOneDriveFolder(client, undefined, 2, cursorOf(first))
    expect(last.files.map((file) => file.fileId)).toEqual(['owned-file-2'])
    expect(last.hasMore).toBe(true)
    expect(last.nextCursor).toBeNull()
    expect(last.incompleteReason).toEqual(expect.any(String))
  })

  it.each([
    ['duplicates', [items(1)[0], items(1)[0]]],
    ['over 1000', items(1001)],
    ['lowercase metadata', [{ id: 'file-id' }]],
    ['invalid ID', [{ ...items(1)[0], Id: 'bad id' }]],
    ['oversized ID', [{ ...items(1)[0], Id: 'x'.repeat(2049) }]],
    ['bad size', [{ ...items(1)[0], Size: '123' }]],
    ['bad ETag', [{ ...items(1)[0], ETag: {} }]],
    ['unbounded ETag', [{ ...items(1)[0], ETag: 'x'.repeat(2049) }]],
    ['malformed nextLink', { value: [], nextLink: {} }],
    ['malformed truncated', { value: [], truncated: 'true' }],
  ])('rejects %s safely at the MCP boundary', async (_label, response) => {
    const { client } = flow(() => (Array.isArray(response) ? { value: response } : response))
    const mcp = await connect(client)
    const result = await call(mcp, { limit: 1 })
    expect(result.isError).toBe(true)
    expect(result.structuredContent).toBeUndefined()
    expectSanitized(result)
  })

  it.each([
    ['native URL', { nextLink: URL_CANARY }, true, false],
    ['native truncated', { truncated: true }, true, true],
    ['both flags', { nextLink: URL_CANARY, truncated: true }, true, true],
    ['empty URL', { nextLink: '', truncated: false }, false, false],
    ['null URL', { nextLink: null }, false, false],
    ['no flags', {}, false, false],
  ] as const)(
    'distinguishes local exhaustion from %s',
    async (_label, flags, incomplete, truncated) => {
      const { client, fetchFn } = flow(() => ({ value: items(3), ...flags }))
      const first = await listOneDriveFolder(client, undefined, 2)
      expect(first.hasMore).toBe(true)
      expect(first.nextCursor).toEqual(expect.any(String))
      expect(first.incompleteReason).toEqual(incomplete ? expect.any(String) : null)
      expect(first.truncated).toBe(truncated)
      const last = await listOneDriveFolder(client, undefined, 2, cursorOf(first))
      expect(last.files).toHaveLength(1)
      expect(last.hasMore).toBe(incomplete)
      expect(last.nextCursor).toBeNull()
      expect(last.incompleteReason).toEqual(incomplete ? expect.any(String) : null)
      expect(last.truncated).toBe(truncated)
      expect(fetchFn).toHaveBeenCalledTimes(2)
      expect(fetchFn.mock.calls.every(([url]) => url === FLOW_URL)).toBe(true)
      expectSanitized([first, last])
    },
  )

  it.each([undefined, 'owned-folder'])(
    'conservatively marks the 1000-item boundary for %s, even without native flags',
    async (folderId) => {
      const { client } = flow(() => items(1000))
      let cursor: string | undefined
      for (let index = 0; index < 10; index++) {
        const page = await listOneDriveFolder(client, folderId, 100, cursor)
        expect(page.files).toHaveLength(100)
        expect(page.files[0]?.fileId).toBe(`owned-file-${index * 100}`)
        expect(page.hasMore).toBe(true)
        expect(page.truncated).toBe(true)
        expect(page.incompleteReason).toContain('minimum threshold')
        expect(page.incompleteReason).toContain('not beyond it')
        if (index < 9) expect(page.nextCursor).toEqual(expect.any(String))
        else expect(page.nextCursor).toBeNull()
        cursor = page.nextCursor ?? undefined
      }
    },
  )

  it('keeps root/child empty windows complete and defaults to an outer 50-item page', async () => {
    const { client, records } = flow(() => [])
    for (const folderId of [undefined, 'owned-folder']) {
      const page = await listOneDriveFolder(client, folderId, 50)
      expect(page).toEqual({
        folderId: folderId ?? null,
        files: [],
        hasMore: false,
        nextCursor: null,
        incompleteReason: null,
        truncated: false,
      })
    }
    expect(records.every(({ args }) => args.top === 1000)).toBe(true)
    expect(oneDriveListFolderInputSchema.parse({})).toEqual({ limit: 50 })
    for (const limit of [0, 101, 1000])
      expect(oneDriveListFolderInputSchema.safeParse({ limit }).success).toBe(false)
  })

  it('keeps search capped at 100 with no cursor or folder-window behavior', async () => {
    const { client, records } = flow(() => ({ value: items(100), truncated: true }))
    const page = await searchOneDrive(client, 'synthetic', 100)
    expect(page.files).toHaveLength(100)
    expect(page.hasMore).toBe(true)
    expect(page.nextCursor).toBeNull()
    expect(records[0]?.args).toEqual({ query: 'synthetic', top: 100 })
    expect(oneDriveSearchInputSchema.safeParse({ query: 'synthetic', limit: 101 }).success).toBe(
      false,
    )
    const oversized = flow(() => items(101))
    await expect(searchOneDrive(oversized.client, 'synthetic', 100)).rejects.toThrow('metadata')
  })

  it('preserves the bounded outer tool-result budget rather than returning a huge 1000-item MCP result', async () => {
    const response = items(250).map((item) => ({
      ...item,
      Id: 'x'.repeat(2040) + item.Id.slice(-5),
      Name: 'n'.repeat(512),
    }))
    const { client } = flow(() => response)
    const mcp = await connect(client)
    const tooLarge = await call(mcp, { limit: 100 })
    expect(tooLarge.isError).toBe(true)
    expect(JSON.stringify(tooLarge)).toContain('output limit')
    const bounded = await call(mcp, { limit: 1 })
    expect(bounded.isError).not.toBe(true)
    expect(oneDriveListFolderOutputSchema.parse(bounded.structuredContent).files).toHaveLength(1)
  })
})

describe('OneDrive folder-only transport budget', () => {
  it('accepts folder metadata above 256 KiB, while search and other lists retain that limit', async () => {
    const { client } = flow(() => ({ value: items(250), padding: 'x'.repeat(300 * 1024) }))
    await expect(client.call('onedrive_list_folder', { top: 1000 })).resolves.toBeDefined()
    for (const operation of ['onedrive_search_files', 'list_mail_folders'] as const) {
      const args =
        operation === 'onedrive_search_files' ? { query: 'synthetic', top: 100 } : { top: 50 }
      await expect(client.call(operation, args)).rejects.toThrow('size limit')
    }
    await expect(client.call('onedrive_get_metadata', { fileId: 'owned-file-0' })).rejects.toThrow(
      'size limit',
    )
    // Binary transports retain their existing 6 MiB bound.
    await expect(
      client.call('onedrive_get_content', { fileId: 'owned-file-0' }),
    ).resolves.toBeDefined()
  })

  it.each(['content-length', 'stream'] as const)(
    'rejects over 4 MiB by %s and cancels the body',
    async (mode) => {
      const cancel = vi.fn()
      const fetchFn = vi.fn<typeof fetch>(
        async () =>
          new Response(
            new ReadableStream({
              start(controller) {
                if (mode === 'stream') controller.enqueue(new Uint8Array(4 * 1024 * 1024 + 1))
              },
              cancel,
            }),
            {
              headers: {
                'Content-Length': mode === 'content-length' ? String(4 * 1024 * 1024 + 1) : '1',
              },
            },
          ),
      )
      const client = new PowerAutomateClient({ baseUrl: FLOW_URL, gatewayKey: 'test', fetchFn })
      await expect(client.call('onedrive_list_folder', { top: 1000 })).rejects.toThrow('size limit')
      expect(cancel).toHaveBeenCalledOnce()
      expect(fetchFn).toHaveBeenCalledOnce()
    },
  )
})
