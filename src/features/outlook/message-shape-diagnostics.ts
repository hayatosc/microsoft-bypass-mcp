import type { z } from 'zod'

// Only these schema-owned paths may produce field diagnostics. Numeric slots
// match recipient indices structurally; neither paths nor indices are emitted.
const messageFields = [
  [[], 'MESSAGE'],
  [['id'], 'ID'],
  [['subject'], 'SUBJECT'],
  [['from'], 'FROM'],
  [['from', 'emailAddress'], 'FROM_EMAIL_ADDRESS'],
  [['from', 'emailAddress', 'name'], 'FROM_NAME'],
  [['from', 'emailAddress', 'address'], 'FROM_ADDRESS'],
  [['sentDateTime'], 'SENT_DATE'],
  [['receivedDateTime'], 'RECEIVED_DATE'],
  [['parentFolderId'], 'PARENT_FOLDER_ID'],
  [['conversationId'], 'CONVERSATION_ID'],
  [['hasAttachments'], 'HAS_ATTACHMENTS'],
  [['importance'], 'IMPORTANCE'],
  [['isRead'], 'IS_READ'],
  [['bodyPreview'], 'BODY_PREVIEW'],
  [['body'], 'BODY'],
  [['body', 'contentType'], 'BODY_CONTENT_TYPE'],
  [['body', 'content'], 'BODY_CONTENT'],
  [['toRecipients'], 'TO_RECIPIENTS'],
  [['toRecipients', 0], 'TO_RECIPIENT'],
  [['toRecipients', 0, 'emailAddress'], 'TO_EMAIL_ADDRESS'],
  [['toRecipients', 0, 'emailAddress', 'name'], 'TO_NAME'],
  [['toRecipients', 0, 'emailAddress', 'address'], 'TO_ADDRESS'],
  [['ccRecipients'], 'CC_RECIPIENTS'],
  [['ccRecipients', 0], 'CC_RECIPIENT'],
  [['ccRecipients', 0, 'emailAddress'], 'CC_EMAIL_ADDRESS'],
  [['ccRecipients', 0, 'emailAddress', 'name'], 'CC_NAME'],
  [['ccRecipients', 0, 'emailAddress', 'address'], 'CC_ADDRESS'],
] as const

type FieldCode = (typeof messageFields)[number][1]
type FailureKind = 'MISSING' | 'NULL' | 'TYPE' | 'INVALID'
type DiagnosticCode =
  | `${FieldCode}_${FailureKind}`
  | 'MESSAGE_LIST_ENVELOPE'
  | 'MESSAGE_LIST_LIMIT'
  | 'MESSAGE_SHAPE_OTHER'
  | 'MESSAGE_SHAPE_TRUNCATED'
const MAX_DIAGNOSTIC_CODES = 32

function isIndex(value: PropertyKey | undefined): boolean {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}
function valueAtPath(body: unknown, path: readonly PropertyKey[]): unknown {
  let value = body
  for (const key of path) {
    if (typeof value !== 'object' || value === null || !Object.hasOwn(value, key)) return undefined
    const child: unknown = Reflect.get(value, key)
    value = child
  }
  return value
}

/** Failure-only, finite allowlist. Never emit Zod messages, input keys or values. */
export function messageShapeDiagnostics(
  body: unknown,
  issues: readonly z.core.$ZodIssue[],
  list = false,
): string {
  const codes = new Set<DiagnosticCode>()
  for (const issue of issues) {
    let path = issue.path
    if (list) {
      if (
        path.length === 0 ||
        (path.length === 1 && (path[0] === 'value' || path[0] === '@odata.nextLink'))
      ) {
        codes.add(
          path[0] === 'value' && issue.code === 'too_big'
            ? 'MESSAGE_LIST_LIMIT'
            : 'MESSAGE_LIST_ENVELOPE',
        )
        continue
      }
      if (path[0] !== 'value' || !isIndex(path[1])) {
        codes.add('MESSAGE_SHAPE_OTHER')
        continue
      }
      path = path.slice(2)
    }
    const field = messageFields.find(
      ([allowed]) =>
        allowed.length === path.length &&
        allowed.every((key, index) =>
          typeof key === 'number' ? isIndex(path[index]) : key === path[index],
        ),
    )?.[1]
    if (field === undefined) {
      codes.add('MESSAGE_SHAPE_OTHER')
      continue
    }
    if (
      issue.code !== 'invalid_type' &&
      issue.code !== 'invalid_value' &&
      issue.code !== 'invalid_format' &&
      issue.code !== 'too_small' &&
      issue.code !== 'too_big'
    ) {
      codes.add('MESSAGE_SHAPE_OTHER')
      continue
    }
    const value = valueAtPath(body, issue.path)
    const kind: FailureKind =
      value === undefined
        ? 'MISSING'
        : value === null
          ? 'NULL'
          : issue.code === 'invalid_type' ||
              (issue.code === 'invalid_value' && typeof value !== 'string')
            ? 'TYPE'
            : 'INVALID'
    codes.add(`${field}_${kind}`)
  }
  if (codes.size === 0) codes.add('MESSAGE_SHAPE_OTHER')
  // Stable order independent of input property/issue order and entry indices.
  const sorted = [...codes].sort()
  if (sorted.length > MAX_DIAGNOSTIC_CODES)
    return [...sorted.slice(0, MAX_DIAGNOSTIC_CODES - 1), 'MESSAGE_SHAPE_TRUNCATED'].join(',')
  return sorted.join(',')
}
