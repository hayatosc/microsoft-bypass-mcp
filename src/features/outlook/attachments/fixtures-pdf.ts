import { zlibSync } from 'fflate'

const encoder = new TextEncoder()

function join(chunks: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(chunks.reduce((length, part) => length + part.length, 0))
  let offset = 0
  for (const part of chunks) {
    result.set(part, offset)
    offset += part.length
  }
  return result
}

function stream(dictionary: string, bytes: Uint8Array): Uint8Array {
  return join([
    encoder.encode(`<< ${dictionary} /Length ${bytes.length} >>\nstream\n`),
    bytes,
    encoder.encode('\nendstream'),
  ])
}

/** Real classic-xref PDFs, assembled in memory with correct byte offsets. */
export function makePdf(
  options: {
    pages?: number
    text?: string
    compressed?: boolean
    content?: string
    streamDictionary?: string
    trailer?: string
    contentRepeats?: number
  } = {},
): Uint8Array {
  const count = options.pages ?? 2
  const text = options.text ?? 'Hello attachment'
  const escape = (value: string) =>
    value.replaceAll('\\', '\\\\').replaceAll('(', '\\(').replaceAll(')', '\\)')
  const objects: Uint8Array[] = [
    encoder.encode('<< /Type /Catalog /Pages 2 0 R >>'),
    encoder.encode(
      `<< /Type /Pages /Count ${count} /Kids [${Array.from({ length: count }, (_, index) => `${4 + index * 2} 0 R`).join(' ')}] >>`,
    ),
    encoder.encode('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'),
  ]
  for (let index = 0; index < count; index++) {
    objects.push(
      encoder.encode(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${options.contentRepeats ? `[${Array.from({ length: options.contentRepeats }, () => `${5 + index * 2} 0 R`).join(' ')}]` : `${5 + index * 2} 0 R`} >>`,
      ),
    )
    const raw = encoder.encode(
      options.content ?? `BT /F1 12 Tf 72 720 Td (${escape(text)} page ${index + 1}) Tj ET`,
    )
    const data = options.compressed ? zlibSync(raw) : raw
    objects.push(
      stream(
        `${options.compressed ? '/Filter /FlateDecode' : ''}${options.streamDictionary ? ` ${options.streamDictionary}` : ''}`,
        data,
      ),
    )
  }
  const chunks: Uint8Array[] = [encoder.encode('%PDF-1.7\n')]
  let length = chunks[0]?.length ?? 0
  const offsets = [0]
  for (let index = 0; index < objects.length; index++) {
    const object = objects[index]
    if (!object) throw new Error('Invalid fixture')
    offsets.push(length)
    const header = encoder.encode(`${index + 1} 0 obj\n`)
    const footer = encoder.encode('\nendobj\n')
    chunks.push(header, object, footer)
    length += header.length + object.length + footer.length
  }
  const xref = length
  const table = encoder.encode(
    `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets
      .slice(1)
      .map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`)
      .join(
        '',
      )}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R ${options.trailer ?? ''} >>\nstartxref\n${xref}\n%%EOF\n`,
  )
  chunks.push(table)
  return join(chunks)
}

export function makeObjectStreamPdf(): Uint8Array {
  const values = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let body = ''
  let header = ''
  for (const [index, value] of values.entries()) {
    header += `${index + 1} ${body.length} `
    body += value + '\n'
  }
  const objectBytes = zlibSync(encoder.encode(header + body))
  const chunks = [encoder.encode('%PDF-1.5\n')]
  const offsets = new Map<number, number>()
  const add = (id: number, data: Uint8Array) => {
    offsets.set(
      id,
      chunks.reduce((length, part) => length + part.length, 0),
    )
    chunks.push(encoder.encode(`${id} 0 obj\n`), data, encoder.encode('\nendobj\n'))
  }
  add(5, stream(`/Type /ObjStm /N 4 /First ${header.length} /Filter /FlateDecode`, objectBytes))
  add(6, stream('', encoder.encode('BT /F1 12 Tf 72 720 Td (Object stream text) Tj ET')))
  const xrefOffset = chunks.reduce((length, part) => length + part.length, 0)
  const xref = new Uint8Array(8 * 7)
  const view = new DataView(xref.buffer)
  const entry = (id: number, type: number, first: number, second: number) => {
    const at = id * 7
    view.setUint8(at, type)
    view.setUint32(at + 1, first)
    view.setUint16(at + 5, second)
  }
  entry(0, 0, 0, 65535)
  for (let id = 1; id <= 4; id++) entry(id, 2, 5, id - 1)
  entry(5, 1, offsets.get(5) ?? 0, 0)
  entry(6, 1, offsets.get(6) ?? 0, 0)
  entry(7, 1, xrefOffset, 0)
  add(7, stream('/Type /XRef /Size 8 /Root 1 0 R /W [1 4 2]', xref))
  chunks.push(encoder.encode(`startxref\n${xrefOffset}\n%%EOF\n`))
  return join(chunks)
}

export function makeJpegStreamPdf(): Uint8Array {
  // Locally generated single white pixel; no external or private image data.
  const jpeg = Uint8Array.from(
    atob(
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3+iiigD//2Q==',
    ),
    (char) => char.charCodeAt(0),
  )
  const objects = [
    encoder.encode('<< /Type /Catalog /Pages 2 0 R >>'),
    encoder.encode('<< /Type /Pages /Count 1 /Kids [3 0 R] >>'),
    encoder.encode(
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 6 0 R >> >> /Contents 5 0 R >>',
    ),
    encoder.encode('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'),
    stream('', encoder.encode('q /Im1 Do Q BT /F1 12 Tf 72 720 Td (JPEG-stream text) Tj ET')),
    stream(
      '/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode',
      jpeg,
    ),
  ]
  const chunks = [encoder.encode('%PDF-1.7\n')]
  const offsets = [0]
  for (const [index, object] of objects.entries()) {
    offsets.push(chunks.reduce((length, part) => length + part.length, 0))
    chunks.push(encoder.encode(`${index + 1} 0 obj\n`), object, encoder.encode('\nendobj\n'))
  }
  const xref = chunks.reduce((length, part) => length + part.length, 0)
  chunks.push(
    encoder.encode(
      `xref\n0 7\n0000000000 65535 f \n${offsets
        .slice(1)
        .map((offset) => String(offset).padStart(10, '0') + ' 00000 n \n')
        .join('')}trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`,
    ),
  )
  return join(chunks)
}

/** One blank page, synthetic RC4-128 password fixture generated locally with pypdf. */
export function makeEncryptedPdf(): Uint8Array {
  return Uint8Array.from(
    atob(
      'JVBERi0xLjMKJeLjz9MKMSAwIG9iago8PAovUHJvZHVjZXIgPDMwYTZlYmJiZmE+Cj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9UeXBlIC9QYWdlcwovQ291bnQgMQovS2lkcyBbIDQgMCBSIF0KPj4KZW5kb2JqCjMgMCBvYmoKPDwKL1R5cGUgL0NhdGFsb2cKL1BhZ2VzIDIgMCBSCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9UeXBlIC9QYWdlCi9SZXNvdXJjZXMgPDwKPj4KL01lZGlhQm94IFsgMC4wIDAuMCA3MiA3MiBdCi9QYXJlbnQgMiAwIFIKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL1YgMgovUiAzCi9MZW5ndGggMTI4Ci9QIDQyOTQ5NjcyOTIKL0ZpbHRlciAvU3RhbmRhcmQKL08gPDQyMzAzNmVhODk0MDNlOWZhY2ZlMTY2MzM4ZDYzMGI0YTFlYjVlODc5YzY2ZDM5MmVhMWYyOGZlY2IyNjFkMDg+Ci9VIDxmNWNmNGZiMThiYjk1NDA1NTUxZWI0MGQzMzUxZDc5ZDI4YmY0ZTVlNGU3NThhNDE2NDAwNGU1NmZmZmEwMTA4Pgo+PgplbmRvYmoKeHJlZgowIDYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDE1IDAwMDAwIG4gCjAwMDAwMDAwNTkgMDAwMDAgbiAKMDAwMDAwMDExOCAwMDAwMCBuIAowMDAwMDAwMTY3IDAwMDAwIG4gCjAwMDAwMDAyNTkgMDAwMDAgbiAKdHJhaWxlcgo8PAovU2l6ZSA2Ci9Sb290IDMgMCBSCi9JbmZvIDEgMCBSCi9JRCBbIDw2NDY2NjM2MTY2MzUzNDMyMzczOTMwMzMzMjMwMzY2NDY0MzEzMTM0MzQzMTY0MzA2MjM1NjI2MTM5MzYzMTYyPiA8NjQ2NjYzNjE2NjM1MzQzMjM3MzkzMDMzMzIzMDM2NjQ2NDMxMzEzNDM0MzE2NDMwNjIzNTYyNjEzOTM2MzE2Mj4gXQovRW5jcnlwdCA1IDAgUgo+PgpzdGFydHhyZWYKNDc0CiUlRU9GCg==',
    ),
    (char) => char.charCodeAt(0),
  )
}
