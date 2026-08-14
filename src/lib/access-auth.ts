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
import type { RemoteJWKSet } from 'jose'

import { getAccessConfig } from './env.js'
import type { Bindings } from './env.js'

/**
 * JWKS resolvers memoized per team domain. Module scope is shared across
 * requests within a Worker isolate, so the resolver — and the JWKS cache held
 * inside `jose` — survives between requests instead of re-fetching the signing
 * keys on every call.
 */
const remoteJwkSets = new Map<string, RemoteJWKSet>()

function getRemoteJwkSet(domain: string): RemoteJWKSet {
  // Strip a trailing slash so `TEAM_DOMAIN` values ending in "/" do not produce
  // a double-slash certs URL.
  const normalized = domain.replace(/\/+$/, '')
  const cached = remoteJwkSets.get(normalized)
  if (cached) {
    return cached
  }
  const jwks = createRemoteJWKSet(new URL(`${normalized}/cdn-cgi/access/certs`))
  remoteJwkSets.set(normalized, jwks)
  return jwks
}

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
      await jwtVerify(token, getRemoteJwkSet(config.domain), {
        issuer: config.domain,
        audience: config.aud,
      })
    } catch {
      return c.json({ error: 'Unauthorized' }, 401)
    }

    await next()
  })
}
