import { describe, expect, it, vi } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../lib/power-automate.js'
import { syntheticAttachment } from '../outlook/attachments/synthetic-fixtures.js'
import { readOneDriveFile } from './service.js'
import {
  normalizeOneDriveList,
  normalizeOneDriveMetadata,
  normalizeOneDriveMetadataFor,
} from './service.js'

const pdf = {
  Id: 'file-1',
  Name: 'report.pdf',
  Size: 123,
  MediaType: 'application/pdf',
  IsFolder: false,
  LastModified: '2026-01-01T00:00:00Z',
  ETag: 'etag',
}

describe('OneDrive normalization', () => {
  it('normalizes native connector arrays without exposing URLs', () => {
    expect(normalizeOneDriveList([pdf], 10)).toEqual({
      files: [
        {
          fileId: 'file-1',
          name: 'report.pdf',
          contentType: 'application/pdf',
          size: 123,
          isFolder: false,
          supportedFormat: 'pdf',
          readable: true,
          lastModifiedDateTime: '2026-01-01T00:00:00Z',
          eTag: 'etag',
          limitation: null,
        },
      ],
      hasMore: false,
      nextCursor: null,
      incompleteReason: null,
      truncated: false,
    })
  })
  it('normalizes native BlobMetadataPage and preserves hasMore from nextLink', () => {
    expect(normalizeOneDriveList({ value: [pdf], nextLink: 'native-token' }, 10)).toMatchObject({
      hasMore: true,
      nextCursor: null,
      truncated: false,
      incompleteReason: expect.stringContaining('bounded page'),
    })
  })
  it('marks folders and unsupported files unreadable', () => {
    expect(
      normalizeOneDriveMetadata({
        Id: 'folder-1',
        Name: 'Folder',
        Size: 0,
        MediaType: '',
        IsFolder: true,
      }),
    ).toMatchObject({
      isFolder: true,
      readable: false,
      limitation: 'Folders must be listed, not read as documents.',
    })
    expect(
      normalizeOneDriveMetadata({
        Id: 'img-1',
        Name: 'photo.png',
        Size: 12,
        MediaType: 'image/png',
        IsFolder: false,
      }),
    ).toMatchObject({ supportedFormat: null, readable: false })
  })
  it('rejects malformed native connector envelopes and mismatched requested identity', () => {
    expect(() => normalizeOneDriveList({ value: [{ id: 1 }] }, 10)).toThrow('metadata')
    expect(() => normalizeOneDriveMetadataFor(pdf, 'other')).toThrow('ID does not match')
  })
  it('rejects metadata/content races before decoding content', async () => {
    const file = syntheticAttachment('pdf')
    const records: unknown[] = []
    const fetchFn = vi.fn<typeof fetch>(async (_url, init) => {
      if (typeof init?.body !== 'string') throw new Error('Expected a JSON request body')
      const request = z
        .object({ operation: z.string(), requestId: z.string(), args: z.unknown() })
        .parse(JSON.parse(init.body))
      records.push(request)
      const data =
        request.operation === 'onedrive_get_metadata'
          ? { ...pdf, Size: file.size }
          : { metadata: { ...pdf, Size: file.size + 1 }, contentBytes: file.contentBytes }
      return Response.json({
        ok: true,
        requestId: request.requestId,
        operation: request.operation,
        data,
      })
    })
    const client = new PowerAutomateClient({
      baseUrl: 'https://example.test',
      gatewayKey: 'key',
      fetchFn,
    })
    await expect(
      readOneDriveFile(client, {
        fileId: 'file-1',
        selection: { format: 'pdf', pageStart: 1, pageEnd: 1, maxCharacters: 1000 },
      }),
    ).rejects.toThrow('metadata changed')
    expect(
      records.map((record) => z.object({ operation: z.string() }).parse(record).operation),
    ).toEqual(['onedrive_get_metadata', 'onedrive_get_content'])
  })
})
