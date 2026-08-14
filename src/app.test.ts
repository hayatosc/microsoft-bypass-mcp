import { describe, expect, it } from 'vitest'

import app from './app.js'
import { getAccessConfig, getPowerAutomateUrl } from './lib/env.js'

const env = {
  POWER_AUTOMATE_URL: 'https://example.test/flow',
}

describe('mcp endpoint', () => {
  it('handles an initialize request with no app-level auth', async () => {
    const res = await app.request(
      '/mcp',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json, text/event-stream',
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: {
            protocolVersion: '2025-06-18',
            capabilities: {},
            clientInfo: { name: 'test', version: '0.0.0' },
          },
        }),
      },
      env,
    )
    expect(res.status).toBe(200)
  })
})

describe('public info', () => {
  it('GET / exposes the server name and tools', async () => {
    const res = await app.request('/')
    expect(res.status).toBe(200)
    const text = await res.text()
    expect(text).toContain('outlook_list_messages')
    expect(text).toContain('outlook_search_messages')
    expect(text).toContain('outlook_get_message')
  })
})

describe('env accessors', () => {
  it('fail fast on missing required values', () => {
    expect(() => getPowerAutomateUrl({})).toThrow('POWER_AUTOMATE_URL')
  })

  it('returns null when Access is not configured', () => {
    expect(getAccessConfig({})).toBeNull()
  })

  it('returns the config when both Access values are set', () => {
    expect(
      getAccessConfig({
        TEAM_DOMAIN: 'https://team.cloudflareaccess.com',
        POLICY_AUD: 'aud',
      }),
    ).toEqual({ domain: 'https://team.cloudflareaccess.com', aud: 'aud' })
  })

  it('fails fast on a partially configured Access pair', () => {
    expect(() => getAccessConfig({ TEAM_DOMAIN: 'https://team.cloudflareaccess.com' })).toThrow(
      'TEAM_DOMAIN',
    )
  })
})
