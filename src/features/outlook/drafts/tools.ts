import type { McpServer } from '@modelcontextprotocol/server'

import type { PowerAutomateClient } from '../../../lib/power-automate.js'
import { PowerAutomateError } from '../../../lib/power-automate.js'
import { DocumentSourceError } from '../../documents/source.js'
import {
  addDraftAttachmentInputSchema,
  createDraftInputSchema,
  createReplyDraftInputSchema,
  draftAttachmentOutputSchema,
  draftSummaryOutputSchema,
} from './schema.js'
import {
  DraftError,
  addDraftAttachment,
  createDraft,
  createReplyDraft,
  draftResult,
} from './service.js'

const draftDescriptionSuffix =
  'Requires explicit user approval. Creates a draft only and never sends mail. Do not blindly retry after an ambiguous failure; inspect Drafts first.'

const attachmentDescriptionSuffix =
  'Requires explicit user approval and only attaches bytes after the host has materialized an approved source. Do not pass chat upload URLs. Creates or changes draft content only and never sends mail. Do not blindly retry after an ambiguous failure; inspect Drafts first.'

const annotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: true,
}

type SafeErrorResult = {
  isError: true
  content: [{ type: 'text'; text: string }]
  structuredContent: undefined
}

function errorText(error: unknown): string {
  if (
    error instanceof DraftError ||
    error instanceof PowerAutomateError ||
    error instanceof DocumentSourceError
  )
    return error.message
  return 'Draft operation failed safely'
}

async function guarded<T>(run: () => Promise<T>): Promise<T | SafeErrorResult> {
  try {
    return await run()
  } catch (error) {
    return {
      isError: true,
      content: [{ type: 'text', text: errorText(error) }],
      structuredContent: undefined,
    }
  }
}

export function registerDraftTools(server: McpServer, client: PowerAutomateClient) {
  server.registerTool(
    'outlook_create_draft',
    {
      title: 'Create Outlook draft',
      description:
        'Create a university Outlook draft with plain-text body and up to 50 total recipients. ' +
        draftDescriptionSuffix,
      inputSchema: createDraftInputSchema,
      outputSchema: draftSummaryOutputSchema,
      annotations,
    },
    async (input) => guarded(async () => draftResult(await createDraft(client, input))),
  )

  server.registerTool(
    'outlook_create_reply_draft',
    {
      title: 'Create Outlook reply draft',
      description:
        'Create a single-sender reply draft for one existing university Outlook message with a plain-text body. ' +
        draftDescriptionSuffix,
      inputSchema: createReplyDraftInputSchema,
      outputSchema: draftSummaryOutputSchema,
      annotations,
    },
    async (input) => guarded(async () => draftResult(await createReplyDraft(client, input))),
  )

  server.registerTool(
    'outlook_add_draft_attachment',
    {
      title: 'Add Outlook draft attachment',
      description:
        'Add one bounded attachment to an existing Outlook draft using strict canonical base64 bytes and verified draft metadata. ' +
        attachmentDescriptionSuffix,
      inputSchema: addDraftAttachmentInputSchema,
      outputSchema: draftAttachmentOutputSchema,
      annotations,
    },
    async (input) => guarded(async () => draftResult(await addDraftAttachment(client, input))),
  )
}
