/**
 * MCP server factory exposing the three fixed Outlook read-only tools.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'

import type { PowerAutomateClient } from '../../lib/power-automate.js'
import { normalizeMessage, normalizeMessageList } from './normalize.js'
import {
  getMessageInputSchema,
  getMessageOutputSchema,
  listMessagesInputSchema,
  listMessagesOutputSchema,
  searchMessagesInputSchema,
  searchMessagesOutputSchema,
} from './schema.js'

export const TOOL_NAMES = [
  'outlook_list_messages',
  'outlook_search_messages',
  'outlook_get_message',
] as const

/**
 * Creates a fresh MCP server wired to the given Power Automate client.
 * Each tool returns both text content and structuredContent validated against
 * its output schema.
 */
export function createOutlookMcpServer(client: PowerAutomateClient): McpServer {
  const server = new McpServer({ name: 'university-outlook', version: '0.1.0' })

  server.registerTool(
    'outlook_list_messages',
    {
      title: 'List messages',
      description: 'List the most recent mailbox messages (metadata only).',
      inputSchema: listMessagesInputSchema,
      outputSchema: listMessagesOutputSchema,
    },
    async ({ limit }) => {
      const messages = normalizeMessageList(await client.call('list_messages', { top: limit }))
      const result = { messages }
      return {
        content: [{ type: 'text', text: JSON.stringify(result) }],
        structuredContent: result,
      }
    },
  )

  server.registerTool(
    'outlook_search_messages',
    {
      title: 'Search messages',
      description: 'Search mailbox messages by a free-text query (metadata only).',
      inputSchema: searchMessagesInputSchema,
      outputSchema: searchMessagesOutputSchema,
    },
    async ({ query, limit }) => {
      const messages = normalizeMessageList(
        await client.call('search_messages', { query, top: limit }),
      )
      const result = { messages }
      return {
        content: [{ type: 'text', text: JSON.stringify(result) }],
        structuredContent: result,
      }
    },
  )

  server.registerTool(
    'outlook_get_message',
    {
      title: 'Get message',
      description: 'Fetch a full message, including its body, by ID.',
      inputSchema: getMessageInputSchema,
      outputSchema: getMessageOutputSchema,
    },
    async ({ messageId }) => {
      const message = normalizeMessage(await client.call('get_message', { messageId }))
      return {
        content: [{ type: 'text', text: JSON.stringify(message) }],
        structuredContent: message,
      }
    },
  )

  return server
}
