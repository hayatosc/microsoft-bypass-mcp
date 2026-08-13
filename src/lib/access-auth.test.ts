import { Hono } from 'hono'
import { describe, expect, it } from 'vitest'

import { createAccessAuth } from './access-auth.js'
import type { Bindings } from './env.js'

function makeApp() {
  const app = new Hono<{ Bindings: Bindings }>()
  app.use('/mcp', createAccessAuth())
  app.all('/mcp', (c) => c.json({ ok: true }))
  return app
}

describe('createAccessAuth', () => {
  it('passes through when Access is not configured (local dev)', async () => {
    const res = await makeApp().request('/mcp', { method: 'POST' }, {})
    expect(res.status).toBe(200)
  })

  it('rejects a request without a JWT assertion when Access is configured', async () => {
    const res = await makeApp().request(
      '/mcp',
      { method: 'POST' },
      {
        TEAM_DOMAIN: 'https://team.cloudflareaccess.com',
        POLICY_AUD: 'aud',
      },
    )
    expect(res.status).toBe(401)
  })

  it('rejects an invalid JWT assertion when Access is configured', async () => {
    const res = await makeApp().request(
      '/mcp',
      {
        method: 'POST',
        headers: { 'Cf-Access-Jwt-Assertion': 'garbage' },
      },
      {
        TEAM_DOMAIN: 'https://team.cloudflareaccess.com',
        POLICY_AUD: 'aud',
      },
    )
    expect(res.status).toBe(401)
  })
})
