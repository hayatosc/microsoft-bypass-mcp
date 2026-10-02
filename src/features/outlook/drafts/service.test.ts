import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { PowerAutomateClient } from '../../../lib/power-automate.js'
import { addDraftAttachment, createDraft, createReplyDraft } from './service.js'

function mockClient(data: unknown): { records: unknown[]; client: PowerAutomateClient } {
  const records: unknown[] = []
  const fetchFn: typeof fetch = async (_input, init) => {
    const body = typeof init?.body === 'string' ? JSON.parse(init.body) : undefined
    records.push(body)
    const request = z.object({ operation: z.string(), requestId: z.string() }).parse(body)
    return Response.json({
      ok: true,
      requestId: request.requestId,
      operation: request.operation,
      data,
    })
  }
  return {
    records,
    client: new PowerAutomateClient({
      baseUrl: 'https://example.test/flow',
      gatewayKey: 'test-gateway-key',
      fetchFn,
    }),
  }
}

describe('draft service', () => {
  it('normalizes create draft summaries only', async () => {
    const { records, client } = mockClient({ id: 'draft-1', isDraft: true, webLink: 'secret' })

    await expect(
      createDraft(client, {
        to: ['to@example.com'],
        cc: ['cc@example.com'],
        bcc: ['bcc@example.com'],
        subject: 'Subject',
        body: 'Body',
      }),
    ).resolves.toEqual({ draftId: 'draft-1', isDraft: true })

    expect(records).toHaveLength(1)
    expect(records[0]).toMatchObject({
      operation: 'create_draft',
      args: {
        to: ['to@example.com'],
        cc: ['cc@example.com'],
        bcc: ['bcc@example.com'],
        subject: 'Subject',
        body: 'Body',
      },
    })
  })

  it('normalizes reply draft summaries only', async () => {
    const { records, client } = mockClient({ id: 'reply-draft-1', isDraft: true })

    await expect(
      createReplyDraft(client, { messageId: 'message-1', body: 'Thanks' }),
    ).resolves.toEqual({ draftId: 'reply-draft-1', isDraft: true })

    expect(records[0]).toMatchObject({
      operation: 'create_reply_draft',
      args: { messageId: 'message-1', body: 'Thanks' },
    })
  })

  it('normalizes attachment metadata and verifies requested draft, name, and size', async () => {
    const { records, client } = mockClient({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
      contentBytes: 'secret',
    })

    await expect(
      addDraftAttachment(client, {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      }),
    ).resolves.toEqual({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
    })

    expect(records[0]).toMatchObject({
      operation: 'add_draft_attachment',
      args: {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      },
    })
  })

  it('rejects attachment responses for another draft', async () => {
    const { client } = mockClient({
      draftId: 'draft-2',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
    })

    await expect(
      addDraftAttachment(client, {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      }),
    ).rejects.toThrow('draftId does not match')
  })

  it.each([889, 2097152])('accepts verified raw-file size %i from the flow', async (size) => {
    const { records, client } = mockClient({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size,
    })
    await expect(
      addDraftAttachment(client, {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: btoa('x'.repeat(size)),
      }),
    ).resolves.toEqual({ draftId: 'draft-1', attachmentId: 'attachment-1', name: 'note.txt', size })
    expect(records).toHaveLength(1)
  })

  it.each([
    { attachmentId: 'bad id' },
    { attachmentId: '.' },
    { name: 'NOTE.txt' },
    { size: 0 },
    { size: -1 },
    { size: 3.5 },
    { size: 2097153 },
    { size: 2147483648 },
  ])('rejects forged or out-of-bound raw-file output without retrying %j', async (change) => {
    const { records, client } = mockClient({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 5,
      ...change,
    })
    await expect(
      addDraftAttachment(client, {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      }),
    ).rejects.toThrow('inspect Drafts before retrying')
    expect(records).toHaveLength(1)
  })

  it('rejects attachment responses with mismatched raw-file size', async () => {
    const { records, client } = mockClient({
      draftId: 'draft-1',
      attachmentId: 'attachment-1',
      name: 'note.txt',
      size: 4,
    })

    await expect(
      addDraftAttachment(client, {
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      }),
    ).rejects.toThrow('size does not match')
    expect(records).toHaveLength(1)
  })
})

it.each([
  { isDraft: true },
  { id: 'draft', isDraft: false },
  { id: 'draft', isDraft: true, unexpected: 'PRIVATE-SENTINEL' },
])('validates draft state and drops raw fields', async (data) => {
  const { client } = mockClient(data)
  const result = createDraft(client, { to: ['to@example.edu'], subject: 'S', body: 'B' })
  if (data.id !== undefined && data.isDraft === true) {
    await expect(result).resolves.toEqual({ draftId: 'draft', isDraft: true })
  } else {
    await expect(result).rejects.toThrow('inspect Drafts before retrying')
  }
})
