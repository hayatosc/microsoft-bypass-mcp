import { z } from 'zod'

import type { PowerAutomateClient } from '../../../lib/power-automate.js'
import { boundedToolResult } from '../../documents/source.js'
import { messageIdSchema } from '../schema.js'
import { draftAttachmentOutputSchema, draftAttachmentSize } from './schema.js'
import type {
  AddDraftAttachmentInput,
  CreateDraftInput,
  CreateReplyDraftInput,
  DraftAttachmentOutput,
  DraftSummaryOutput,
} from './schema.js'

export class DraftError extends Error {
  constructor(message: string) {
    super(`${message}. Write outcome may be ambiguous; inspect Drafts before retrying.`)
    this.name = 'DraftError'
  }
}

function normalizeDraftSummary(data: unknown): DraftSummaryOutput {
  const parsed = z.object({ id: messageIdSchema, isDraft: z.literal(true) }).safeParse(data)
  if (!parsed.success) throw new DraftError('Draft operation returned an unexpected response')
  return { draftId: parsed.data.id, isDraft: true }
}

function normalizeDraftAttachment(
  data: unknown,
  input: AddDraftAttachmentInput,
): DraftAttachmentOutput {
  const parsed = draftAttachmentOutputSchema.strip().safeParse(data)
  if (!parsed.success)
    throw new DraftError('Draft attachment operation returned an unexpected response')
  if (parsed.data.draftId !== input.draftId)
    throw new DraftError('Draft attachment response draftId does not match')
  if (parsed.data.name !== input.name)
    throw new DraftError('Draft attachment response name does not match')
  const expectedSize = draftAttachmentSize(input.contentBytes)
  if (parsed.data.size !== expectedSize)
    throw new DraftError('Draft attachment response size does not match')
  return parsed.data
}

export async function createDraft(
  client: PowerAutomateClient,
  input: CreateDraftInput,
): Promise<DraftSummaryOutput> {
  return normalizeDraftSummary(await client.call('create_draft', input))
}

export async function createReplyDraft(
  client: PowerAutomateClient,
  input: CreateReplyDraftInput,
): Promise<DraftSummaryOutput> {
  return normalizeDraftSummary(await client.call('create_reply_draft', input))
}

export async function addDraftAttachment(
  client: PowerAutomateClient,
  input: AddDraftAttachmentInput,
): Promise<DraftAttachmentOutput> {
  return normalizeDraftAttachment(await client.call('add_draft_attachment', input), input)
}

export function draftResult<T extends Record<string, unknown>>(value: T) {
  return boundedToolResult(value)
}
