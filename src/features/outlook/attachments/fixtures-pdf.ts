import { zlibSync } from 'fflate'

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
  const encoder = new TextEncoder()
  const objects: Uint8Array[] = [
    encoder.encode('<< /Type /Catalog /Pages 2 0 R >>'),
    encoder.encode(
      `<< /Type /Pages /Count ${count} /Kids [${Array.from({ length: count }, (_, i) => `${4 + i * 2} 0 R`).join(' ')}] >>`,
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
    const header = encoder.encode(
      `<< /Length ${data.length}${options.compressed ? ' /Filter /FlateDecode' : ''}${options.streamDictionary ? ` ${options.streamDictionary}` : ''} >>\nstream\n`,
    )
    const footer = encoder.encode('\nendstream')
    const stream = new Uint8Array(header.length + data.length + footer.length)
    stream.set(header)
    stream.set(data, header.length)
    stream.set(footer, header.length + data.length)
    objects.push(stream)
  }
  const chunks: Uint8Array[] = [encoder.encode('%PDF-1.7\n')]
  let length = chunks[0]?.length ?? 0
  const offsets = [0]
  for (let i = 0; i < objects.length; i++) {
    const object = objects[i]
    if (!object) throw new Error('Invalid fixture')
    offsets.push(length)
    const header = encoder.encode(`${i + 1} 0 obj\n`)
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
  length += table.length
  const result = new Uint8Array(length)
  let cursor = 0
  for (const chunk of chunks) {
    result.set(chunk, cursor)
    cursor += chunk.length
  }
  return result
}
