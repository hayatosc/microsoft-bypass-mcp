import { z } from 'zod'

import { messageIdSchema } from '../schema.js'

export const MAX_DRAFT_RECIPIENTS = 50
export const MAX_DRAFT_SUBJECT_CHARACTERS = 512
export const MAX_DRAFT_BODY_CHARACTERS = 20000
export const MAX_DRAFT_ATTACHMENT_BYTES = 2 * 1024 * 1024
export const MAX_DRAFT_ATTACHMENT_NAME_CHARACTERS = 255
export const MAX_DRAFT_ATTACHMENT_CONTENT_TYPE_CHARACTERS = 127

// oxlint-disable-next-line no-control-regex
const safeFileNamePattern = /^[^/\\\u0000-\u001f\u007f-\u009f]+$(?![\s\S])/u
const simpleMimeTypePattern = /^[A-Za-z0-9!#$&^_.+-]+\/[A-Za-z0-9!#$&^_.+-]+$(?![\s\S])/

export const draftEmailAddressSchema = z.string().trim().min(1).max(254).email()

function recipientListSchema(required: boolean) {
  const schema = z.array(draftEmailAddressSchema).max(MAX_DRAFT_RECIPIENTS)
  return required ? schema.min(1) : schema
}

function totalRecipients(value: {
  to: string[]
  cc?: string[] | undefined
  bcc?: string[] | undefined
}): number {
  return value.to.length + (value.cc?.length ?? 0) + (value.bcc?.length ?? 0)
}

function canonicalBase64DecodedLength(value: string): number | null {
  if (
    value.length === 0 ||
    value.length > 4 * Math.ceil(MAX_DRAFT_ATTACHMENT_BYTES / 3) ||
    value.length % 4 !== 0
  )
    return null
  if (!/^[A-Za-z0-9+/]*={0,2}$(?![\s\S])/.test(value)) return null
  let decoded: string
  try {
    decoded = atob(value)
  } catch {
    return null
  }
  if (btoa(decoded) !== value) return null
  return decoded.length
}

export function draftAttachmentSize(contentBytes: string): number {
  const size = canonicalBase64DecodedLength(contentBytes)
  if (size === null) throw new DraftSchemaError('Invalid canonical base64 attachment content')
  return size
}

export class DraftSchemaError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DraftSchemaError'
  }
}

export const createDraftInputSchema = z
  .object({
    to: recipientListSchema(true),
    cc: recipientListSchema(false).optional(),
    bcc: recipientListSchema(false).optional(),
    subject: z.string().max(MAX_DRAFT_SUBJECT_CHARACTERS),
    body: z.string().max(MAX_DRAFT_BODY_CHARACTERS),
  })
  .strict()
  .refine((value) => totalRecipients(value) <= MAX_DRAFT_RECIPIENTS, {
    message: `total recipients must be at most ${MAX_DRAFT_RECIPIENTS}`,
  })

export const createReplyDraftInputSchema = z
  .object({
    messageId: messageIdSchema,
    body: z.string().max(MAX_DRAFT_BODY_CHARACTERS),
  })
  .strict()

export const addDraftAttachmentInputSchema = z
  .object({
    draftId: messageIdSchema,
    name: z
      .string()
      .min(1)
      .max(MAX_DRAFT_ATTACHMENT_NAME_CHARACTERS)
      .regex(
        safeFileNamePattern,
        'filename must not contain slash, backslash, or control characters',
      ),
    contentType: z
      .string()
      .min(1)
      .max(MAX_DRAFT_ATTACHMENT_CONTENT_TYPE_CHARACTERS)
      .regex(simpleMimeTypePattern, 'contentType must be a simple MIME type without parameters'),
    contentBytes: z
      .string()
      .max(4 * Math.ceil(MAX_DRAFT_ATTACHMENT_BYTES / 3))
      .refine(
        (value) => {
          const size = canonicalBase64DecodedLength(value)
          return size !== null && size > 0 && size <= MAX_DRAFT_ATTACHMENT_BYTES
        },
        {
          message: `contentBytes must be strict canonical base64 for 1-${MAX_DRAFT_ATTACHMENT_BYTES} raw bytes`,
        },
      ),
  })
  .strict()

export const draftSummaryOutputSchema = z
  .object({ draftId: messageIdSchema, isDraft: z.literal(true) })
  .strict()

export const draftAttachmentOutputSchema = z
  .object({
    draftId: messageIdSchema,
    attachmentId: messageIdSchema,
    name: z.string().min(1).max(MAX_DRAFT_ATTACHMENT_NAME_CHARACTERS),
    size: z
      .number()
      .int()
      .min(1)
      .max(MAX_DRAFT_ATTACHMENT_BYTES)
      .describe('Verified raw-file size in bytes, not Microsoft Graph attachment metadata size.'),
  })
  .strict()

export type CreateDraftInput = z.infer<typeof createDraftInputSchema>
export type CreateReplyDraftInput = z.infer<typeof createReplyDraftInputSchema>
export type AddDraftAttachmentInput = z.infer<typeof addDraftAttachmentInputSchema>
export type DraftSummaryOutput = z.infer<typeof draftSummaryOutputSchema>
export type DraftAttachmentOutput = z.infer<typeof draftAttachmentOutputSchema>
