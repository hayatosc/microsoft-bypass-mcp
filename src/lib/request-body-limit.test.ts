import { Hono } from 'hono'
import { describe, expect, it, vi } from 'vitest'

import { boundedRequestBody } from './request-body-limit.js'

function app() {
  const instance = new Hono()
  instance.use('*', boundedRequestBody(16))
  instance.post('/', async (c) => c.text(await c.req.text()))
  instance.get('/', (c) => c.text('ready'))
  return instance
}

describe('actual MCP request byte limit', () => {
  it('allows a bodyless request', async () => {
    expect(await (await app().request('/')).text()).toBe('ready')
  })
  it('preserves exact bytes at the boundary', async () => {
    const res = await app().request('/', { method: 'POST', body: 'x'.repeat(16) })
    expect(res.status).toBe(200)
    expect(await res.text()).toBe('x'.repeat(16))
  })
  it.each([undefined, '1', '999'])(
    'counts real bytes independently of Content-Length %s',
    async (declared) => {
      const headers = declared === undefined ? {} : { 'content-length': declared }
      const res = await app().request('/', { method: 'POST', headers, body: 'x'.repeat(17) })
      expect(res.status).toBe(413)
      expect(await res.text()).not.toContain('x'.repeat(17))
    },
  )
  it('counts UTF-8 bytes rather than characters', async () => {
    expect((await app().request('/', { method: 'POST', body: 'あ'.repeat(6) })).status).toBe(413)
  })
  it('cancels a stream once over the bound', async () => {
    const cancel = vi.fn()
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(17))
      },
      cancel,
    })
    const res = await app().request('/', {
      method: 'POST',
      body: stream,
      headers: { 'content-length': '1' },
    })
    expect(res.status).toBe(413)
    expect(cancel).toHaveBeenCalledOnce()
  })
  it('sanitizes stream exceptions', async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.error(new Error('private-data'))
      },
    })
    const res = await app().request('/', { method: 'POST', body: stream })
    expect(res.status).toBe(400)
    expect(await res.text()).not.toContain('private-data')
  })
})

it('coalesces many small chunks and ignores empty chunks', async () => {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (let i = 0; i < 10000; i++) controller.enqueue(new Uint8Array(0))
      for (let i = 0; i < 16; i++) controller.enqueue(new Uint8Array([97]))
      controller.close()
    },
  })
  const res = await app().request('/', { method: 'POST', body: stream })
  expect(res.status).toBe(200)
  expect(await res.text()).toBe('a'.repeat(16))
})
