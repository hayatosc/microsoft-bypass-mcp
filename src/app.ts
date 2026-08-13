/**
 * Hono app exposing the public info page and the MCP Streamable HTTP endpoint
 * at /mcp. Authentication is delegated to Cloudflare Access (OAuth) placed in
 * front of the Worker; this app performs no access control itself.
 */
import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js'
import { Hono } from 'hono'

import { createOutlookMcpServer, TOOL_NAMES } from './features/outlook/server.js'
import { getPowerAutomateUrl } from './lib/env.js'
import type { Bindings } from './lib/env.js'
import { PowerAutomateClient } from './lib/power-automate.js'

const app = new Hono<{ Bindings: Bindings }>()

app.get('/', (c) =>
  c.json({
    name: 'university-outlook',
    version: '0.1.0',
    tools: TOOL_NAMES,
  }),
)

app.all('/mcp', async (c) => {
  // Fresh, stateless server + transport per request; no state survives.
  const client = new PowerAutomateClient({ baseUrl: getPowerAutomateUrl(c.env) })
  const transport = new WebStandardStreamableHTTPServerTransport()
  const server = createOutlookMcpServer(client)
  await server.connect(transport)
  return transport.handleRequest(c.req.raw)
})

export default app
