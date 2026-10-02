import type { MiddlewareHandler } from 'hono'

/** Count actual bytes after Access authentication, even if Content-Length lies. */
export function boundedRequestBody(maxBytes: number): MiddlewareHandler {
  return async (c, next) => {
    const request = c.req.raw
    if (request.body === null) return next()
    const reader = request.body.getReader()
    let bytes = new Uint8Array(Math.min(64 * 1024, maxBytes))
    let size = 0
    try {
      while (true) {
        const chunk = await reader.read()
        if (chunk.done) break
        if (chunk.value.byteLength === 0) continue
        const nextSize = size + chunk.value.byteLength
        if (nextSize > maxBytes) {
          await reader.cancel().catch(() => {})
          return c.json({ error: 'MCP request body exceeds the 4 MiB limit' }, 413)
        }
        if (nextSize > bytes.byteLength) {
          const grown = new Uint8Array(Math.min(maxBytes, Math.max(nextSize, bytes.byteLength * 2)))
          grown.set(bytes.subarray(0, size))
          bytes = grown
        }
        bytes.set(chunk.value, size)
        size = nextSize
      }
    } catch {
      return c.json({ error: 'MCP request body could not be read' }, 400)
    } finally {
      reader.releaseLock()
    }
    const headers = new Headers(request.headers)
    headers.delete('transfer-encoding')
    headers.set('content-length', String(size))
    c.req.raw = new Request(request, {
      method: request.method,
      headers,
      body: bytes.subarray(0, size),
    })
    return next()
  }
}
