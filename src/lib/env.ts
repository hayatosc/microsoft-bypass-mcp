/**
 * Worker bindings and validated accessors. Missing required values fail fast
 * instead of silently defaulting.
 */
export interface Bindings {
  POWER_AUTOMATE_URL?: string
  /** Cloudflare Access team domain, e.g. `https://<team>.cloudflareaccess.com`. */
  TEAM_DOMAIN?: string
  /** Cloudflare Access Application Audience (AUD) tag for the MCP application. */
  POLICY_AUD?: string
}

/** Returns the Power Automate trigger URL or throws if it is not configured. */
export function getPowerAutomateUrl(bindings: Bindings): string {
  const url = bindings.POWER_AUTOMATE_URL
  if (!url) {
    throw new Error('Missing required binding: POWER_AUTOMATE_URL')
  }
  return url
}

/** Access JWT verification configuration, derived from the team domain + AUD. */
export interface AccessConfig {
  domain: string
  aud: string
}

/**
 * Returns the Cloudflare Access verification config, or `null` when neither
 * `TEAM_DOMAIN` nor `POLICY_AUD` is set (local development has no Access in
 * front). A partially configured pair is a misconfiguration and fails fast.
 */
export function getAccessConfig(bindings: Bindings): AccessConfig | null {
  const domain = bindings.TEAM_DOMAIN
  const aud = bindings.POLICY_AUD
  if (domain && aud) {
    return { domain, aud }
  }
  if (domain || aud) {
    throw new Error('Incomplete Cloudflare Access config: set both TEAM_DOMAIN and POLICY_AUD')
  }
  return null
}
