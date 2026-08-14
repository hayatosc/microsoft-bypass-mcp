/**
 * Hono app exposing the public info page and the MCP Streamable HTTP endpoint
 * at /mcp. The /mcp endpoint is guarded by Cloudflare Access JWT validation
 * (defense in depth behind the Access policy itself); the MCP handler uses a
 * fresh, stateless server per request.
 */
import { createMcpHandler } from '@modelcontextprotocol/server'
import { Hono } from 'hono'

import { createOutlookMcpServer, TOOL_NAMES } from './features/outlook/server.js'
import { createAccessAuth } from './lib/access-auth.js'
import { getPowerAutomateUrl } from './lib/env.js'
import type { Bindings } from './lib/env.js'
import { PowerAutomateClient } from './lib/power-automate.js'

const app = new Hono<{ Bindings: Bindings }>()

app.get('/', (c) =>
  c.json({
    name: 'university-m365',
    version: '0.1.0',
    tools: TOOL_NAMES,
  }),
)

app.use('/mcp', createAccessAuth())

app.all('/mcp', (c) => {
  // Fresh, stateless server + handler per request; no state survives.
  const client = new PowerAutomateClient({ baseUrl: getPowerAutomateUrl(c.env) })
  const handler = createMcpHandler(() => createOutlookMcpServer(client))
  return handler.fetch(c.req.raw)
})

export default app
