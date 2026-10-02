/** Public, synthetic documents only. Never use real mailbox data in tests. */
import { strToU8, zipSync } from 'fflate'

export function syntheticDocx(text = 'Synthetic document body') {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p><w:p><w:r><w:t>' +
        text +
        '</w:t></w:r></w:p></w:body></w:document>',
    ),
  })
}
export function syntheticXlsx() {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Synthetic</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="B2"><f>SUM(B1)</f><v>42</v></c></row></sheetData></worksheet>',
    ),
  })
}
export function syntheticPdf() {
  const stream = 'BT /F1 12 Tf 72 720 Td (Synthetic PDF page) Tj ET\n'
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}endstream`,
  ]
  let pdf = '%PDF-1.4\n'
  const offsets = [0]
  objects.forEach((object, index) => {
    offsets.push(pdf.length)
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xref = pdf.length
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const offset of offsets.slice(1)) pdf += `${offset.toString().padStart(10, '0')} 00000 n \n`
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return strToU8(pdf)
}
/** Valid 889-byte DOCX with synthetic Graph metadata size 1223, not raw size. */
export function syntheticGraphSizeAttachment() {
  const file = syntheticAttachment('docx')
  const binary = atob(file.contentBytes)
  // fflate emits an empty ZIP comment. Replace its EOCD length and append a
  // comment to reach the reproduction size without changing document content.
  const commentLength = 889 - binary.length
  const padded =
    binary.slice(0, -2) +
    String.fromCharCode(commentLength & 0xff, commentLength >> 8) +
    ' '.repeat(commentLength)
  return { ...file, size: 1223, contentBytes: btoa(padded) }
}

export function syntheticAttachment(format: 'pdf' | 'docx' | 'xlsx') {
  const bytes =
    format === 'pdf' ? syntheticPdf() : format === 'docx' ? syntheticDocx() : syntheticXlsx()
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    id: 'att-1',
    name: `synthetic.${format}`,
    contentType: 'application/octet-stream',
    size: bytes.length,
    isInline: false,
    contentBytes: btoa(binary),
  }
}
