/**
 * Worker bindings and validated accessors. Missing required values fail fast
 * instead of silently defaulting.
 */
export interface Bindings {
  POWER_AUTOMATE_URL?: string
}

/** Returns the Power Automate trigger URL or throws if it is not configured. */
export function getPowerAutomateUrl(bindings: Bindings): string {
  const url = bindings.POWER_AUTOMATE_URL
  if (!url) {
    throw new Error('Missing required binding: POWER_AUTOMATE_URL')
  }
  return url
}
