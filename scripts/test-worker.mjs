/** Offline smoke test: bundled production app in real workerd, synthetic files only. */
import assert from 'node:assert/strict'
import { build } from 'esbuild'
import { Miniflare, convertV4MiniflareOptions } from 'miniflare'
import { syntheticAttachment } from '../src/features/outlook/attachments/synthetic-fixtures.ts'
import { makePdf } from '../src/features/outlook/attachments/fixtures-pdf.ts'

const bundle = await build({entryPoints: ['src/index.ts'], bundle: true, write: false,
  format: 'esm', platform: 'browser', target: 'es2022', minify: true})
let format = 'pdf'
let calls = 0
const runtime = new Miniflare(convertV4MiniflareOptions({
  modules: true, script: bundle.outputFiles[0].text,
  compatibilityDate: '2025-08-03', cf: false,
  bindings: {POWER_AUTOMATE_URL: 'https://flow.example.test/', POWER_AUTOMATE_GATEWAY_KEY: 'synthetic'},
  outboundService: async (request) => {
    assert.equal(new URL(request.url).hostname, 'flow.example.test', 'Unexpected parser network request')
    assert.equal(request.headers.get('X-MCP-Gateway-Key'), 'synthetic')
    const body = await request.json()
    calls++
    let file = syntheticAttachment(format)
    if (format === 'pdf') {
      const bytes = makePdf({pages: 2, text: 'Synthetic PDF page', compressed: true})
      file = {...file, size: bytes.length, contentBytes: Buffer.from(bytes).toString('base64')}
    }
    const data = body.operation === 'list_attachments' ? {value: [file], '@odata.nextLink': ''} : file
    return Response.json({ok: true, requestId: body.requestId, operation: body.operation, data})
  },
}))
async function rpc(method, params) {
  const response = await runtime.dispatchFetch('http://localhost/mcp', {
    method: 'POST', headers: {'Content-Type': 'application/json', Accept: 'application/json, text/event-stream'},
    body: JSON.stringify({jsonrpc: '2.0', id: 1, method, params}),
  })
  assert.equal(response.status, 200)
  const text = await response.text()
  const envelope = JSON.parse(text.startsWith('event:') || text.startsWith('data:')
    ? text.split('\n').find((line) => line.startsWith('data:')).slice(5).trim() : text)
  assert.equal(envelope.error, undefined)
  assert.equal(envelope.result.isError, undefined, JSON.stringify(envelope.result))
  return envelope.result
}
try {
  const tools = await rpc('tools/list', {})
  assert.equal(tools.tools.length, 6)
  for (format of ['pdf', 'docx', 'xlsx']) {
    const target = {messageId: 'msg-1', attachmentId: 'att-1'}
    const listing = await rpc('tools/call', {name: 'outlook_list_attachments', arguments: {messageId: 'msg-1'}})
    assert.equal(listing.structuredContent.attachments[0].supportedFormat, format)
    assert.equal(listing.structuredContent.hasMore, false)
    assert.ok(!JSON.stringify(listing).includes('contentBytes'))
    const inspection = await rpc('tools/call', {name: 'outlook_inspect_attachment', arguments: target})
    assert.equal(inspection.structuredContent.structure.format, format)
    const selection = format === 'pdf' ? {format, pageStart: 1, pageEnd: 1}
      : format === 'docx' ? {format, offset: 0, length: 100}
        : {format, sheet: 'Budget', range: 'A1:B2'}
    const read = await rpc('tools/call', {name: 'outlook_read_attachment', arguments: {...target, selection}})
    assert.equal(read.structuredContent.data.format, format)
    assert.ok(JSON.stringify(read).includes(format === 'pdf' ? 'Synthetic PDF page' : format === 'docx' ? 'Synthetic document body' : 'Synthetic'))
    assert.ok(!JSON.stringify(read).includes('contentBytes'))
  }
  assert.equal(calls, 9)
  console.log('workerd smoke passed: six tools, PDF/DOCX/XLSX list + inspect + read, no parser network access')
} finally {
  await runtime.dispose()
}
