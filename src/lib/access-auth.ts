/**
 * Cloudflare Access JWT validation middleware. When `TEAM_DOMAIN` and
 * `POLICY_AUD` are configured, requests to the guarded route must carry a
 * valid `Cf-Access-Jwt-Assertion` header signed by Cloudflare Access; otherwise
 * they are rejected with 401. When neither is configured (local development),
 * the middleware passes requests through, since Access only fronts the
 * deployed Worker.
 */
import { createMiddleware } from 'hono/factory'
import { createRemoteJWKSet, jwtVerify } from 'jose'

import { getAccessConfig } from './env.js'
import type { Bindings } from './env.js'

/** Middleware enforcing Cloudflare Access JWT assertions. */
export function createAccessAuth() {
  return createMiddleware<{ Bindings: Bindings }>(async (c, next) => {
    const config = getAccessConfig(c.env)
    if (!config) {
      await next()
      return
    }

    const token = c.req.header('Cf-Access-Jwt-Assertion')
    if (!token) {
      return c.json({ error: 'Unauthorized' }, 401)
    }

    try {
      const jwks = createRemoteJWKSet(new URL(`${config.domain}/cdn-cgi/access/certs`))
      await jwtVerify(token, jwks, {
        issuer: config.domain,
        audience: config.aud,
      })
    } catch {
      return c.json({ error: 'Unauthorized' }, 401)
    }

    await next()
  })
}
