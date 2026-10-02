import { describe, expect, it } from 'vitest'

import {
  MAX_DRAFT_ATTACHMENT_BYTES,
  addDraftAttachmentInputSchema,
  createDraftInputSchema,
  createReplyDraftInputSchema,
  draftAttachmentSize,
} from './schema.js'

function base64Of(size: number): string {
  let binary = ''
  const chunk = '\0'.repeat(8192)
  let remaining = size
  while (remaining > 0) {
    const next = Math.min(remaining, chunk.length)
    binary += chunk.slice(0, next)
    remaining -= next
  }
  return btoa(binary)
}

describe('draft schemas', () => {
  it('accepts a bounded create draft request', () => {
    expect(
      createDraftInputSchema.parse({
        to: ['to@example.edu'],
        cc: ['cc@example.edu'],
        bcc: ['bcc@example.edu'],
        subject: 'Subject',
        body: 'Plain text body',
      }),
    ).toEqual({
      to: ['to@example.edu'],
      cc: ['cc@example.edu'],
      bcc: ['bcc@example.edu'],
      subject: 'Subject',
      body: 'Plain text body',
    })
  })

  it('rejects more than 50 total recipients', () => {
    const recipients = Array.from({ length: 49 }, (_, index) => `u${index}@example.edu`)
    expect(() =>
      createDraftInputSchema.parse({
        to: recipients,
        cc: ['cc@example.edu'],
        bcc: ['bcc@example.edu'],
        subject: 'Subject',
        body: 'Body',
      }),
    ).toThrow()
  })

  it('rejects unknown passthrough properties and invalid recipient addresses', () => {
    expect(() =>
      createDraftInputSchema.parse({
        to: ['not an email'],
        subject: 'Subject',
        body: 'Body',
        graphBody: { internetMessageHeaders: [] },
      }),
    ).toThrow()
  })

  it('uses the same safe message id boundary for reply drafts', () => {
    expect(() =>
      createReplyDraftInputSchema.parse({ messageId: 'bad\nID', body: 'Body' }),
    ).toThrow()
    expect(() =>
      createReplyDraftInputSchema.parse({ messageId: 'bad\u0085ID', body: 'Body' }),
    ).toThrow()
    expect(createReplyDraftInputSchema.parse({ messageId: 'AQMk-safe_id', body: 'Body' })).toEqual({
      messageId: 'AQMk-safe_id',
      body: 'Body',
    })
  })

  it('accepts strict canonical base64 attachment content and computes raw size', () => {
    const input = {
      draftId: 'draft-1',
      name: 'note.txt',
      contentType: 'text/plain',
      contentBytes: 'aGVsbG8=',
    }
    expect(addDraftAttachmentInputSchema.parse(input)).toEqual(input)
    expect(draftAttachmentSize(input.contentBytes)).toBe(5)
  })

  it.each([
    '',
    'aGVsbG8',
    'aGVsbG8===',
    'aGVsbG8-',
    ' aGVsbG8=',
    'AAAA=AAA',
    'AB==',
    'AAB=',
    'AAAA\n',
  ])('rejects non-canonical base64 %s', (contentBytes) => {
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes,
      }),
    ).toThrow()
  })

  it('rejects empty bytes and raw bytes above 2 MiB', () => {
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: btoa(''),
      }),
    ).toThrow()
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: base64Of(MAX_DRAFT_ATTACHMENT_BYTES + 1),
      }),
    ).toThrow()
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: base64Of(MAX_DRAFT_ATTACHMENT_BYTES),
      }),
    ).not.toThrow()
  })

  it('rejects unsafe filenames, MIME parameters, and arbitrary URL/body passthrough', () => {
    for (const name of [
      'bad/name.txt',
      'bad\\name.txt',
      'bad\nname.txt',
      'bad\u007fname.txt',
      'bad\u0085name.txt',
      'bad\u009fname.txt',
    ]) {
      expect(() =>
        addDraftAttachmentInputSchema.parse({
          draftId: 'draft-1',
          name,
          contentType: 'text/plain',
          contentBytes: 'aGVsbG8=',
        }),
      ).toThrow()
    }
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain; charset=utf-8',
        contentBytes: 'aGVsbG8=',
      }),
    ).toThrow()
    expect(() =>
      addDraftAttachmentInputSchema.parse({
        draftId: 'draft-1',
        name: 'note.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
        uploadUrl: 'https://example.test/file',
      }),
    ).toThrow()
  })
})

it.each(['text/plain\n', 'text/plain\r', 'text/plain; charset=utf8'])(
  'rejects trailing MIME control or parameters %j',
  (contentType) => {
    expect(
      addDraftAttachmentInputSchema.safeParse({
        draftId: 'draft',
        name: 'a.txt',
        contentType,
        contentBytes: 'YQ==',
      }).success,
    ).toBe(false)
  },
)
