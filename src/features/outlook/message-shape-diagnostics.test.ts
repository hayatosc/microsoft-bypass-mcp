import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { PowerAutomateError } from '../../lib/power-automate.js'
import { messageShapeDiagnostics } from './message-shape-diagnostics.js'
import { normalizeConversation, normalizeMessage, normalizeMessageList } from './normalize.js'

const CANARY = 'SYNTHETIC_VALUE_SENTINEL'
const KEY_CANARY = 'SYNTHETIC_KEY_SENTINEL'
const message = {
  id: 'synthetic-message',
  subject: 'Synthetic subject',
  from: { emailAddress: { name: 'Synthetic sender', address: 'sender@example.test' } },
  receivedDateTime: '2026-01-01T12:00:00Z',
  sentDateTime: '2026-01-01T11:59:00Z',
  parentFolderId: 'synthetic-folder',
  conversationId: 'synthetic-conversation',
  hasAttachments: false,
  importance: 'normal',
  isRead: false,
  bodyPreview: 'Synthetic preview',
  toRecipients: [{ emailAddress: { name: 'Synthetic recipient', address: 'to@example.test' } }],
  ccRecipients: [{ emailAddress: { name: 'Synthetic copy', address: 'cc@example.test' } }],
  body: { contentType: 'text', content: 'Synthetic body' },
}

function changed(path: readonly PropertyKey[], value: unknown, remove = false): unknown {
  const copy = structuredClone(message)
  let parent: unknown = copy
  for (const key of path.slice(0, -1)) {
    if (typeof parent !== 'object' || parent === null) throw new Error('Invalid synthetic path')
    const child: unknown = Reflect.get(parent, key)
    parent = child
  }
  const key = path.at(-1)
  if (typeof parent !== 'object' || parent === null || key === undefined)
    throw new Error('Invalid synthetic path')
  if (remove) Reflect.deleteProperty(parent, key)
  else Reflect.set(parent, key, value)
  return copy
}
function errorText(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(PowerAutomateError)
    if (error instanceof PowerAutomateError) return error.message
    throw error
  }
  throw new Error('Expected unchanged rejection')
}
function assertDiagnostic(body: unknown, code: string, list = false): void {
  const text = errorText(() =>
    list ? normalizeMessageList({ value: [body] }) : normalizeMessage(body),
  )
  expect(text).toBe(
    `malformed response: expected ${list ? 'a list of messages' : 'a message'} (${code})`,
  )
  for (const forbidden of [
    CANARY,
    KEY_CANARY,
    ...Object.values(message).filter((value) => typeof value === 'string'),
  ])
    expect(text).not.toContain(forbidden)
}

const fields = [
  {
    path: ['id'],
    code: 'ID',
    wrong: 123,
    invalid: 'bad id',
    optional: false,
    nullable: false,
    list: true,
  },
  { path: ['subject'], code: 'SUBJECT', wrong: {}, optional: false, nullable: true, list: true },
  { path: ['from'], code: 'FROM', wrong: CANARY, optional: false, nullable: true, list: true },
  {
    path: ['from', 'emailAddress'],
    code: 'FROM_EMAIL_ADDRESS',
    wrong: [],
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['from', 'emailAddress', 'name'],
    code: 'FROM_NAME',
    wrong: 123,
    optional: false,
    nullable: true,
    list: true,
  },
  {
    path: ['from', 'emailAddress', 'address'],
    code: 'FROM_ADDRESS',
    wrong: {},
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['sentDateTime'],
    code: 'SENT_DATE',
    wrong: false,
    invalid: CANARY,
    optional: true,
    nullable: false,
    list: true,
  },
  {
    path: ['receivedDateTime'],
    code: 'RECEIVED_DATE',
    wrong: [],
    invalid: CANARY,
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['parentFolderId'],
    code: 'PARENT_FOLDER_ID',
    wrong: {},
    invalid: 'bad id',
    optional: true,
    nullable: false,
    list: true,
  },
  {
    path: ['conversationId'],
    code: 'CONVERSATION_ID',
    wrong: true,
    invalid: 'bad id',
    optional: true,
    nullable: false,
    list: true,
  },
  {
    path: ['hasAttachments'],
    code: 'HAS_ATTACHMENTS',
    wrong: CANARY,
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['importance'],
    code: 'IMPORTANCE',
    wrong: 123,
    invalid: CANARY,
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['isRead'],
    code: 'IS_READ',
    wrong: CANARY,
    optional: false,
    nullable: false,
    list: true,
  },
  {
    path: ['bodyPreview'],
    code: 'BODY_PREVIEW',
    wrong: [],
    optional: true,
    nullable: true,
    list: true,
  },
  { path: ['body'], code: 'BODY', wrong: CANARY, optional: false, nullable: false, list: false },
  {
    path: ['body', 'contentType'],
    code: 'BODY_CONTENT_TYPE',
    wrong: {},
    invalid: CANARY,
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['body', 'content'],
    code: 'BODY_CONTENT',
    wrong: [],
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['toRecipients'],
    code: 'TO_RECIPIENTS',
    wrong: CANARY,
    optional: true,
    nullable: false,
    list: false,
  },
  {
    path: ['toRecipients', 0],
    code: 'TO_RECIPIENT',
    wrong: CANARY,
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['toRecipients', 0, 'emailAddress'],
    code: 'TO_EMAIL_ADDRESS',
    wrong: [],
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['toRecipients', 0, 'emailAddress', 'name'],
    code: 'TO_NAME',
    wrong: 123,
    optional: false,
    nullable: true,
    list: false,
  },
  {
    path: ['toRecipients', 0, 'emailAddress', 'address'],
    code: 'TO_ADDRESS',
    wrong: {},
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['ccRecipients'],
    code: 'CC_RECIPIENTS',
    wrong: CANARY,
    optional: true,
    nullable: false,
    list: false,
  },
  {
    path: ['ccRecipients', 0],
    code: 'CC_RECIPIENT',
    wrong: CANARY,
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['ccRecipients', 0, 'emailAddress'],
    code: 'CC_EMAIL_ADDRESS',
    wrong: [],
    optional: false,
    nullable: false,
    list: false,
  },
  {
    path: ['ccRecipients', 0, 'emailAddress', 'name'],
    code: 'CC_NAME',
    wrong: 123,
    optional: false,
    nullable: true,
    list: false,
  },
  {
    path: ['ccRecipients', 0, 'emailAddress', 'address'],
    code: 'CC_ADDRESS',
    wrong: {},
    optional: false,
    nullable: false,
    list: false,
  },
] as const

const rejectedVariants = fields.flatMap((field) => [
  {
    path: field.path,
    code: `${field.code}_TYPE`,
    value: field.wrong,
    list: field.list,
    remove: false,
  },
  ...(!field.optional
    ? [
        {
          path: field.path,
          code: `${field.code}_MISSING`,
          value: undefined,
          list: field.list,
          remove: true,
        },
      ]
    : []),
  ...(!field.nullable
    ? [
        {
          path: field.path,
          code: `${field.code}_NULL`,
          value: null,
          list: field.list,
          remove: false,
        },
      ]
    : []),
  ...('invalid' in field
    ? [
        {
          path: field.path,
          code: `${field.code}_INVALID`,
          value: field.invalid,
          list: field.list,
          remove: false,
        },
      ]
    : []),
])

describe('failure-only sanitized message shape diagnostics', () => {
  it.each(rejectedVariants)('rejects with exactly $code', ({ path, code, value, list, remove }) => {
    const body = changed(path, value, remove)
    assertDiagnostic(body, code)
    if (list) assertDiagnostic(body, code, true)
  })

  it.each(fields.filter((field) => field.optional || field.nullable))(
    'preserves already accepted omissions/nulls for $code',
    (field) => {
      if (field.optional) {
        expect(() => normalizeMessage(changed(field.path, undefined, true))).not.toThrow()
        if (field.list)
          expect(() =>
            normalizeMessageList({ value: [changed(field.path, undefined, true)] }),
          ).not.toThrow()
      }
      if (field.nullable) {
        expect(() => normalizeMessage(changed(field.path, null))).not.toThrow()
        if (field.list)
          expect(() => normalizeMessageList({ value: [changed(field.path, null)] })).not.toThrow()
      }
    },
  )

  it('returns unchanged valid outputs and strips unknown malicious keys', () => {
    const expected = {
      id: message.id,
      subject: message.subject,
      from: message.from.emailAddress,
      receivedDateTime: message.receivedDateTime,
      sentDateTime: message.sentDateTime,
      parentFolderId: message.parentFolderId,
      conversationId: message.conversationId,
      hasAttachments: false,
      importance: 'normal',
      isRead: false,
      bodyPreview: message.bodyPreview,
    }
    const body = { ...message, [KEY_CANARY]: CANARY }
    expect(normalizeMessage(body, message.id)).toEqual({
      ...expected,
      to: [message.toRecipients[0]?.emailAddress],
      cc: [message.ccRecipients[0]?.emailAddress],
      body: message.body,
    })
    expect(normalizeMessageList({ value: [body], [KEY_CANARY]: CANARY })).toEqual({
      messages: [expected],
      hasMore: false,
      page: { nextLink: null, count: 1 },
    })
    expect(errorText(() => normalizeMessage(body, 'different-id'))).toBe(
      'malformed response: message ID does not match',
    )
  })

  it('combines sorted codes without retaining malicious keys, values or nested paths', () => {
    const body = {
      ...message,
      from: { emailAddress: { address: CANARY, [KEY_CANARY]: CANARY } },
      sentDateTime: null,
      receivedDateTime: CANARY,
      body: { contentType: CANARY, content: CANARY, [KEY_CANARY]: CANARY },
      [KEY_CANARY]: { [CANARY]: CANARY },
    }
    assertDiagnostic(
      body,
      'BODY_CONTENT_TYPE_INVALID,FROM_NAME_MISSING,RECEIVED_DATE_INVALID,SENT_DATE_NULL',
    )
    assertDiagnostic(body, 'FROM_NAME_MISSING,RECEIVED_DATE_INVALID,SENT_DATE_NULL', true)
    assertDiagnostic(
      Object.fromEntries(Object.entries(body).reverse()),
      'BODY_CONTENT_TYPE_INVALID,FROM_NAME_MISSING,RECEIVED_DATE_INVALID,SENT_DATE_NULL',
    )
  })

  it('suppresses recipient and list indices/counts and deduplicates repeated failures', () => {
    const recipient = { emailAddress: { name: true, address: null, [KEY_CANARY]: CANARY } }
    const body = {
      ...message,
      toRecipients: Array.from({ length: 200 }, () => recipient),
      ccRecipients: [recipient],
    }
    assertDiagnostic(body, 'CC_ADDRESS_NULL,CC_NAME_TYPE,TO_ADDRESS_NULL,TO_NAME_TYPE')
    const bad = { ...message, receivedDateTime: CANARY }
    const text = errorText(() =>
      normalizeMessageList({ value: Array.from({ length: 50 }, () => bad) }),
    )
    expect(text).toBe('malformed response: expected a list of messages (RECEIVED_DATE_INVALID)')
  })

  it.each([
    [null, 'MESSAGE_NULL'],
    [undefined, 'MESSAGE_MISSING'],
    [[], 'MESSAGE_TYPE'],
    [CANARY, 'MESSAGE_TYPE'],
  ])('classifies root rejection without showing the input %j', (body, code) => {
    assertDiagnostic(body, code)
    assertDiagnostic(body, code, true)
  })

  it.each([
    null,
    [],
    {},
    { value: null },
    { value: CANARY },
    { value: [], '@odata.nextLink': null },
    { value: [], '@odata.nextLink': 'x'.repeat(8193) + CANARY },
  ])('uses a fixed list envelope code for %j', (body) =>
    expect(errorText(() => normalizeMessageList(body))).toBe(
      'malformed response: expected a list of messages (MESSAGE_LIST_ENVELOPE)',
    ),
  )

  it('combines fixed list-envelope and entry codes without returning hostile nextLink data', () => {
    const text = errorText(() =>
      normalizeMessageList({
        value: [{ ...message, sentDateTime: null }],
        '@odata.nextLink': { [KEY_CANARY]: CANARY },
      }),
    )
    expect(text).toBe(
      'malformed response: expected a list of messages (MESSAGE_LIST_ENVELOPE,SENT_DATE_NULL)',
    )
  })

  it('preserves both list size bounds with a fixed limit code', () => {
    const fifty = Array.from({ length: 50 }, () => message)
    expect(normalizeMessageList({ value: fifty }, 50).messages).toHaveLength(50)
    expect(normalizeMessageList({ value: [] }, 1).messages).toEqual([])
    for (const [value, limit] of [
      [fifty, 49],
      [[...fifty, message], 50],
    ] as const)
      expect(errorText(() => normalizeMessageList({ value }, limit))).toBe(
        'malformed response: expected a list of messages (MESSAGE_LIST_LIMIT)',
      )
    expect(
      errorText(() =>
        normalizeMessageList({
          value: Array.from({ length: 51 }, () => ({ ...message, sentDateTime: null })),
        }),
      ),
    ).toBe('malformed response: expected a list of messages (MESSAGE_LIST_LIMIT,SENT_DATE_NULL)')
  })

  it('bounds combined list diagnostics deterministically without values or counts', () => {
    const listFields = fields.filter((field) => field.list)
    const value = listFields.flatMap((field) => [
      changed(field.path, field.wrong),
      changed(field.path, null),
      changed(field.path, undefined, true),
    ])
    const forward = errorText(() => normalizeMessageList({ value }))
    const reverse = errorText(() => normalizeMessageList({ value: [...value].reverse() }))
    expect(forward).toBe(reverse)
    expect(forward).toContain('MESSAGE_SHAPE_TRUNCATED)')
    const codes = forward.split('(')[1]?.replace(')', '').split(',') ?? []
    expect(codes).toHaveLength(32)
    expect(new Set(codes).size).toBe(32)
    expect(forward.length).toBeLessThanOrEqual(1024)
    expect(forward).not.toMatch(/\d/)
    expect(forward).not.toContain(CANARY)
  })

  it('collapses unknown/malicious issue paths and unrecognized failures to one fixed code', () => {
    const unknown = z
      .object({ [KEY_CANARY]: z.string() })
      .safeParse({ [KEY_CANARY]: { [CANARY]: CANARY } })
    if (unknown.success) throw new Error('Expected synthetic schema failure')
    expect(messageShapeDiagnostics({}, unknown.error.issues)).toBe('MESSAGE_SHAPE_OTHER')
    expect(messageShapeDiagnostics({}, unknown.error.issues, true)).toBe('MESSAGE_SHAPE_OTHER')
    const nestedSchema = z.object({
      from: z.object({ emailAddress: z.object({ [KEY_CANARY]: z.string() }) }),
      toRecipients: z.array(z.object({ emailAddress: z.object({ [KEY_CANARY]: z.string() }) })),
    })
    const nestedBody = {
      from: { emailAddress: { [KEY_CANARY]: { [CANARY]: CANARY } } },
      toRecipients: Array.from({ length: 200 }, () => ({ emailAddress: { [KEY_CANARY]: null } })),
    }
    const nested = nestedSchema.safeParse(nestedBody)
    if (nested.success) throw new Error('Expected synthetic nested failure')
    expect(messageShapeDiagnostics(nestedBody, nested.error.issues)).toBe('MESSAGE_SHAPE_OTHER')
    const nestedList = z.object({ value: z.array(nestedSchema) }).safeParse({ value: [nestedBody] })
    if (nestedList.success) throw new Error('Expected synthetic nested list failure')
    expect(messageShapeDiagnostics({ value: [nestedBody] }, nestedList.error.issues, true)).toBe(
      'MESSAGE_SHAPE_OTHER',
    )
    const custom = z
      .object({ id: z.string().refine(() => false, CANARY) })
      .safeParse({ id: CANARY })
    if (custom.success) throw new Error('Expected synthetic custom failure')
    expect(messageShapeDiagnostics({ id: CANARY }, custom.error.issues)).toBe('MESSAGE_SHAPE_OTHER')
    expect(messageShapeDiagnostics({}, [])).toBe('MESSAGE_SHAPE_OTHER')
  })
})

describe('conversation reuse of sanitized list-mode shape diagnostics', () => {
  it('collapses combined message and recipient failures without values, IDs, paths or indices', () => {
    const recipient = { emailAddress: { name: true, address: null, [KEY_CANARY]: CANARY } }
    const bad = {
      ...message,
      id: `${CANARY} bad id`,
      conversationId: CANARY,
      from: { emailAddress: { address: CANARY, [KEY_CANARY]: CANARY } },
      sentDateTime: null,
      receivedDateTime: CANARY,
      body: { contentType: CANARY, content: CANARY, [KEY_CANARY]: CANARY },
      toRecipients: Array.from({ length: 200 }, () => recipient),
      ccRecipients: [recipient],
      [KEY_CANARY]: { [CANARY]: CANARY },
    }
    const value = [bad, Object.fromEntries(Object.entries(bad).reverse())]
    const text = errorText(() => normalizeConversation({ value }, CANARY))
    expect(text).toBe(
      'malformed response: expected conversation messages (BODY_CONTENT_TYPE_INVALID,CC_ADDRESS_NULL,CC_NAME_TYPE,FROM_NAME_MISSING,ID_INVALID,RECEIVED_DATE_INVALID,SENT_DATE_NULL,TO_ADDRESS_NULL,TO_NAME_TYPE)',
    )
    expect(errorText(() => normalizeConversation({ value: [...value].reverse() }, CANARY))).toBe(
      text,
    )
    for (const forbidden of [
      CANARY,
      KEY_CANARY,
      ...Object.values(message).filter((value) => typeof value === 'string'),
      'value',
      'emailAddress',
      'toRecipients',
      'ccRecipients',
      '[',
      ']',
    ])
      expect(text).not.toContain(forbidden)
    expect(text).not.toMatch(/\d/)
  })

  it.each([
    [null, 'MESSAGE_LIST_ENVELOPE'],
    [[], 'MESSAGE_LIST_ENVELOPE'],
    [{}, 'MESSAGE_LIST_ENVELOPE'],
    [{ value: null }, 'MESSAGE_LIST_ENVELOPE'],
    [{ value: CANARY }, 'MESSAGE_LIST_ENVELOPE'],
    [{ value: [null] }, 'MESSAGE_NULL'],
    [{ value: [CANARY] }, 'MESSAGE_TYPE'],
    [{ value: [changed(['body'], undefined, true)] }, 'BODY_MISSING'],
    [{ value: [message], '@odata.nextLink': null }, 'MESSAGE_LIST_ENVELOPE'],
    [{ value: [message], '@odata.nextLink': CANARY.repeat(500) }, 'MESSAGE_LIST_ENVELOPE'],
    [
      { value: [{ ...message, sentDateTime: null }], '@odata.nextLink': { [KEY_CANARY]: CANARY } },
      'MESSAGE_LIST_ENVELOPE,SENT_DATE_NULL',
    ],
  ])('continues rejecting malformed conversation responses with fixed codes: %j', (body, code) => {
    expect(errorText(() => normalizeConversation(body, message.conversationId))).toBe(
      `malformed response: expected conversation messages (${code})`,
    )
  })

  it('preserves both the schema maximum and requested conversation limit', () => {
    const fifty = Array.from({ length: 50 }, (_, index) => ({
      ...message,
      id: `synthetic-message-${index}`,
    }))
    expect(
      normalizeConversation({ value: fifty }, message.conversationId, 50).messages,
    ).toHaveLength(50)
    expect(normalizeConversation({ value: [] }, message.conversationId, 1).messages).toEqual([])
    for (const [value, limit] of [
      [fifty, 49],
      [[...fifty, { ...message, id: 'synthetic-message-50' }], 51],
    ] as const)
      expect(errorText(() => normalizeConversation({ value }, message.conversationId, limit))).toBe(
        'malformed response: expected conversation messages (MESSAGE_LIST_LIMIT)',
      )
    expect(
      errorText(() =>
        normalizeConversation(
          { value: Array.from({ length: 51 }, () => ({ ...message, sentDateTime: null })) },
          message.conversationId,
        ),
      ),
    ).toBe('malformed response: expected conversation messages (MESSAGE_LIST_LIMIT,SENT_DATE_NULL)')
  })

  it('keeps conversation identity and duplicate errors unchanged and free of diagnostics', () => {
    for (const conversationId of [undefined, 'different-conversation'])
      expect(
        errorText(() =>
          normalizeConversation(
            { value: [{ ...message, conversationId }] },
            message.conversationId,
          ),
        ),
      ).toBe('malformed response: conversation message ID does not match')
    expect(
      errorText(() => normalizeConversation({ value: [message, message] }, message.conversationId)),
    ).toBe('malformed response: duplicate conversation message')
  })

  it('returns unchanged detail outputs, paging and timestamp/ID sorting without diagnostics', () => {
    const expected = {
      id: message.id,
      subject: message.subject,
      from: message.from.emailAddress,
      receivedDateTime: message.receivedDateTime,
      sentDateTime: message.sentDateTime,
      parentFolderId: message.parentFolderId,
      conversationId: message.conversationId,
      hasAttachments: false,
      importance: 'normal',
      isRead: false,
      bodyPreview: message.bodyPreview,
      to: [message.toRecipients[0]?.emailAddress],
      cc: [message.ccRecipients[0]?.emailAddress],
      body: message.body,
    }
    const tie = { ...message, id: `${message.id}-tie` }
    const newer = { ...message, id: 'newer', sentDateTime: '2026-01-01T12:00:00.0000001Z' }
    const { sentDateTime: _unused, ...withoutSent } = message
    const newest = {
      ...withoutSent,
      id: 'newest',
      receivedDateTime: '2026-01-01T12:00:00.0000002Z',
    }
    const { sentDateTime: _unusedExpected, ...expectedWithoutSent } = expected
    const value = [tie, { ...message, [KEY_CANARY]: CANARY }, newer, newest]
    const messages = [
      { ...expectedWithoutSent, id: newest.id, receivedDateTime: newest.receivedDateTime },
      { ...expected, id: newer.id, sentDateTime: newer.sentDateTime },
      expected,
      { ...expected, id: tie.id },
    ]
    const nextLink = 'https://graph.microsoft.com/v1.0/me/messages?$skip=4'
    expect(
      normalizeConversation({ value, '@odata.nextLink': nextLink }, message.conversationId, 4),
    ).toEqual({ messages, hasMore: true, page: { nextLink, count: 4 } })
    expect(
      normalizeConversation({ value, '@odata.nextLink': '' }, message.conversationId, 4),
    ).toEqual({
      messages,
      hasMore: false,
      page: { nextLink: null, count: 4 },
    })
  })

  it('reuses the deterministic 32-code and 1,024-character bounds for conversation failures', () => {
    const value = rejectedVariants.map(({ path, value, remove }) => changed(path, value, remove))
    const forward = errorText(() => normalizeConversation({ value }, message.conversationId))
    const reverse = errorText(() =>
      normalizeConversation({ value: [...value].reverse() }, message.conversationId),
    )
    expect(forward).toBe(reverse)
    expect(forward).toContain('MESSAGE_SHAPE_TRUNCATED)')
    const codes = forward.split('(')[1]?.replace(')', '').split(',') ?? []
    expect(codes).toHaveLength(32)
    expect(new Set(codes).size).toBe(32)
    expect(forward.length).toBeLessThanOrEqual(1024)
    expect(forward).not.toMatch(/\d/)
    expect(forward).not.toContain(CANARY)
    expect(forward).not.toContain(KEY_CANARY)
  })
})
