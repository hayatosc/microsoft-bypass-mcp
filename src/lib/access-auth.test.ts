import { Hono } from 'hono'
import * as jose from 'jose'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { createAccessAuth } from './access-auth.js'
import type { Bindings } from './env.js'

vi.mock('jose', async (importOriginal) => {
  const actual = await importOriginal<typeof jose>()
  return { ...actual, createRemoteJWKSet: vi.fn(actual.createRemoteJWKSet) }
})

function makeApp() {
  const app = new Hono<{ Bindings: Bindings }>()
  app.use('/mcp', createAccessAuth())
  app.all('/mcp', (c) => c.json({ ok: true }))
  return app
}

beforeEach(() => {
  vi.mocked(jose.createRemoteJWKSet).mockClear()
})

afterEach(() => {
  vi.unstubAllGlobals()
})

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

  it('reuses the JWKS resolver across requests for the same team domain', async () => {
    const app = makeApp()
    const env = {
      TEAM_DOMAIN: `https://${crypto.randomUUID()}.cloudflareaccess.com`,
      POLICY_AUD: 'aud',
    }
    const send = () =>
      app.request(
        '/mcp',
        { method: 'POST', headers: { 'Cf-Access-Jwt-Assertion': 'garbage' } },
        env,
      )

    await send()
    await send()

    expect(vi.mocked(jose.createRemoteJWKSet)).toHaveBeenCalledTimes(1)
  })

  it('creates a separate resolver per team domain', async () => {
    const app = makeApp()
    const send = (domain: string) =>
      app.request(
        '/mcp',
        { method: 'POST', headers: { 'Cf-Access-Jwt-Assertion': 'garbage' } },
        { TEAM_DOMAIN: domain, POLICY_AUD: 'aud' },
      )

    await send(`https://${crypto.randomUUID()}.cloudflareaccess.com`)
    await send(`https://${crypto.randomUUID()}.cloudflareaccess.com`)

    expect(vi.mocked(jose.createRemoteJWKSet)).toHaveBeenCalledTimes(2)
  })

  it('accepts a validly signed Access assertion', async () => {
    const domain = `https://${crypto.randomUUID()}.cloudflareaccess.com`
    const aud = 'test-aud'
    const { publicKey, privateKey } = await jose.generateKeyPair('RS256')
    const publicJwk = await jose.exportJWK(publicKey)

    vi.stubGlobal(
      'fetch',
      vi.fn<typeof fetch>(async (input) => {
        const url =
          typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
        if (url.includes('/cdn-cgi/access/certs')) {
          return new Response(
            JSON.stringify({
              keys: [{ ...publicJwk, kid: 'test-kid', alg: 'RS256', use: 'sig' }],
            }),
            { status: 200 },
          )
        }
        return new Response('not found', { status: 404 })
      }),
    )

    const token = await new jose.SignJWT({ email: 'user@example.com' })
      .setProtectedHeader({ alg: 'RS256', kid: 'test-kid' })
      .setIssuer(domain)
      .setAudience(aud)
      .setIssuedAt()
      .setExpirationTime('5m')
      .sign(privateKey)

    const res = await makeApp().request(
      '/mcp',
      { method: 'POST', headers: { 'Cf-Access-Jwt-Assertion': token } },
      { TEAM_DOMAIN: domain, POLICY_AUD: aud },
    )

    expect(res.status).toBe(200)
  })
})
