import { describe, expect, it } from 'vitest'

import { inspectDocx, inspectXlsx, OfficeParseError, readDocx, readXlsx } from './office.js'

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const SHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const encode = (text: string) => new TextEncoder().encode(text)
const entries = (parts: Record<string, string>) =>
  new Map(Object.entries(parts).map(([name, xml]) => [name, encode(xml)]))
const paragraph = (text: string, properties = '') =>
  `<w:p>${properties}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
function docx(body: string, styles?: string) {
  const result = entries({
    'word/document.xml': `<w:document xmlns:w="${WORD}"><w:body>${body}</w:body></w:document>`,
  })
  if (styles !== undefined)
    result.set('word/styles.xml', encode(`<w:styles xmlns:w="${WORD}">${styles}</w:styles>`))
  return result
}
function xlsx(
  data: string,
  options: {
    shared?: string
    date1904?: boolean
    dimension?: string
    extraRelations?: string
  } = {},
) {
  const result = entries({
    'xl/workbook.xml': `<workbook xmlns="${SHEET}" xmlns:r="${REL}"><workbookPr date1904="${options.date1904 ? '1' : '0'}"/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${PACKAGE_REL}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>${options.shared === undefined ? '' : `<Relationship Id="rId2" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`}${options.extraRelations ?? ''}</Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${SHEET}"><dimension ref="${options.dimension ?? 'A1'}"/><sheetData>${data}</sheetData></worksheet>`,
  })
  if (options.shared !== undefined)
    result.set('xl/sharedStrings.xml', encode(`<sst xmlns="${SHEET}">${options.shared}</sst>`))
  return result
}
function withXml(part: Map<string, Uint8Array>, path: string, change: (xml: string) => string) {
  const source = part.get(path)
  if (!source) throw new Error('Missing test fixture part')
  part.set(path, encode(change(new TextDecoder().decode(source))))
  return part
}
function expectCode(operation: () => unknown, code: string) {
  try {
    operation()
    throw new Error('Expected the operation to fail')
  } catch (error) {
    expect(error).toBeInstanceOf(OfficeParseError)
    expect(error).toHaveProperty('code', code)
  }
}

describe('DOCX parsing', () => {
  it('extracts entities, preserved spaces, tabs, line breaks, and table paragraphs', () => {
    const source = docx(
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve"> A &amp; B </w:t><w:tab/><w:t>C</w:t><w:br/><w:t>D</w:t></w:r></w:p><w:tbl><w:tr><w:tc>' +
        paragraph('Table') +
        '</w:tc></w:tr></w:tbl>',
    )
    const result = readDocx(source, {})
    expect(result.text).toBe(' A & B \tC\nD\nTable')
    expect(result.sourceEnd).toBe(result.text.length)
    expect(inspectDocx(source).paragraphCount).toBe(2)
  })

  it('delineates heading and preface sections with deterministic provenance', () => {
    const source = docx(
      paragraph('Intro') +
        paragraph('First', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>') +
        paragraph('Body') +
        paragraph('Second', '<w:pPr><w:outlineLvl w:val="1"/></w:pPr>'),
    )
    const inspection = inspectDocx(source)
    expect(inspection.textCharacters).toBe(23)
    expect(inspection.sections).toEqual([
      {
        id: 'section-1',
        title: 'Document',
        headingLevel: null,
        paragraphStart: 1,
        paragraphEnd: 1,
        start: 0,
        end: 6,
      },
      {
        id: 'section-2',
        title: 'First',
        headingLevel: 1,
        paragraphStart: 2,
        paragraphEnd: 3,
        start: 6,
        end: 17,
      },
      {
        id: 'section-3',
        title: 'Second',
        headingLevel: 2,
        paragraphStart: 4,
        paragraphEnd: 4,
        start: 17,
        end: 23,
      },
    ])
    expect(readDocx(source, { sectionId: 'section-2', offset: 2, length: 4 })).toMatchObject({
      text: 'rst\n',
      sourceStart: 8,
      sourceEnd: 12,
      totalCharacters: 11,
      offset: 2,
      length: 4,
      nextOffset: 6,
    })
    expect(
      inspection.sections
        .map((section) => readDocx(source, { sectionId: section.id }).text)
        .join(''),
    ).toBe(readDocx(source, {}).text)
  })

  it('recognizes custom inherited paragraph styles and explicit body outline overrides', () => {
    const source = docx(
      paragraph('Custom', '<w:pPr><w:pStyle w:val="MyHeading"/></w:pPr>') +
        paragraph('Body', '<w:pPr><w:pStyle w:val="MyHeading"/><w:outlineLvl w:val="9"/></w:pPr>'),
      '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Localized name"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="MyHeading"><w:basedOn w:val="Base"/></w:style>',
    )
    expect(inspectDocx(source).sections).toMatchObject([
      { title: 'Custom', headingLevel: 3, paragraphStart: 1, paragraphEnd: 2 },
    ])
  })

  it('supports arbitrary namespace prefixes, strict OOXML, CDATA, and UTF-16 XML', () => {
    const source = entries({
      'word/document.xml':
        '<?xml version="1.0" encoding="UTF-16"?><q:document xmlns:q="http://purl.oclc.org/ooxml/wordprocessingml/main"><q:body><q:p><q:r><q:t><![CDATA[A < B]]></q:t></q:r></q:p></q:body></q:document>',
    })
    const raw = new TextDecoder().decode(source.get('word/document.xml'))
    const utf16 = new Uint8Array(raw.length * 2 + 2)
    utf16[0] = 0xff
    utf16[1] = 0xfe
    for (let i = 0; i < raw.length; i++) {
      utf16[i * 2 + 2] = raw.charCodeAt(i) & 255
      utf16[i * 2 + 3] = raw.charCodeAt(i) >> 8
    }
    source.set('word/document.xml', utf16)
    expect(readDocx(source, {}).text).toBe('A < B')
  })

  it('omits deleted text, field instructions, scripts, and embedded text boxes', () => {
    const source = docx(
      '<w:p><w:r><w:t>Visible</w:t><w:instrText>https://do-not-fetch.invalid</w:instrText></w:r><w:del>' +
        paragraph('Deleted') +
        '</w:del><w:r><w:drawing><w:txbxContent>' +
        paragraph('Box') +
        '</w:txbxContent></w:drawing></w:r><script xmlns="https://foreign.invalid">fetch()</script></w:p>',
    )
    expect(readDocx(source, {}).text).toBe('Visible')
  })

  it('uses exact UTF-16 offsets and bounded default pagination', () => {
    const source = docx(paragraph('😀' + 'a'.repeat(20_000)))
    expect(readDocx(source, { offset: 2, length: 3 })).toMatchObject({
      text: 'aaa',
      sourceStart: 2,
      sourceEnd: 5,
    })
    expect(readDocx(source, {})).toMatchObject({ length: 10_000, nextOffset: 10_000 })
    expect(readDocx(source, { offset: 20_002 })).toMatchObject({
      text: '',
      length: 0,
      nextOffset: null,
    })
  })

  it('handles empty documents', () => {
    expect(inspectDocx(docx(''))).toMatchObject({
      paragraphCount: 0,
      textCharacters: 0,
      sections: [],
    })
    expect(readDocx(docx(''), {})).toMatchObject({
      text: '',
      nextOffset: null,
      sourceStart: 0,
      sourceEnd: 0,
    })
  })

  it.each([
    { sectionId: 'missing' },
    { offset: -1 },
    { offset: 99 },
    { length: 0 },
    { length: 20_001 },
    { offset: Number.NaN },
    { length: 1.5 },
  ])('rejects invalid selectors %j', (selection) => {
    expectCode(() => readDocx(docx(paragraph('Text')), selection), 'INVALID_RANGE')
  })

  it('rejects paragraph, section, depth, and aggregate XML overflow', () => {
    expectCode(() => inspectDocx(docx('<w:p/>'.repeat(10_001))), 'OFFICE_LIMIT')
    expectCode(
      () =>
        inspectDocx(
          docx(paragraph('H', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>').repeat(201)),
        ),
      'OFFICE_LIMIT',
    )
    expectCode(
      () => inspectDocx(docx('<w:sdt>'.repeat(70) + '<w:p/>' + '</w:sdt>'.repeat(70))),
      'OFFICE_LIMIT',
    )
    const source = docx(paragraph('small'))
    source.set('unused.xml', new Uint8Array(8 * 1024 * 1024))
    expectCode(() => inspectDocx(source), 'OFFICE_LIMIT')
  })

  it('rejects cyclic style inheritance without looping', () => {
    const source = docx(
      paragraph('X', '<w:pPr><w:pStyle w:val="A"/></w:pPr>'),
      '<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>',
    )
    expectCode(() => inspectDocx(source), 'INVALID_OFFICE')
  })
})

describe('XLSX parsing', () => {
  it('inspects actual cells instead of trusting claimed dimensions', () => {
    const source = xlsx(
      '<row r="5"><c r="B5"><v>1</v></c><c r="D5"/></row><row r="9"><c r="C9"><v>2</v></c></row>',
      { dimension: 'A1:XFD1048576' },
    )
    expect(inspectXlsx(source)).toMatchObject({
      sheets: [
        {
          name: 'Data',
          dimensions: {
            range: 'B5:D9',
            firstRow: 5,
            lastRow: 9,
            firstColumn: 2,
            lastColumn: 4,
            cellCount: 3,
          },
        },
      ],
      dateSystem: '1900',
    })
  })

  it('returns precise raw values, blank coordinates, and formula cache status', () => {
    const source = xlsx(
      '<row r="1"><c r="A1"><v>9007199254740993</v></c><c r="B1" t="b"><v>1</v></c><c r="C1"><f>WEBSERVICE("https://do-not-fetch.invalid")</f><v>42</v></c><c r="D1"><f t="shared" si="1"/></c><c r="E1" t="str"><f>""</f><v/></c><c r="F1" t="e"><v>#DIV/0!</v></c></row>',
    )
    const result = readXlsx(source, { sheet: 'Data', range: 'A1:G1' })
    expect(result.cells.map((cell) => cell.value)).toEqual([
      '9007199254740993',
      true,
      '42',
      null,
      '',
      '#DIV/0!',
      null,
    ])
    expect(result.cells[2]).toMatchObject({
      address: 'C1',
      row: 1,
      column: 3,
      hasFormula: true,
      cachedValueMissing: false,
      rawValue: '42',
    })
    expect(result.cells[3]).toMatchObject({ hasFormula: true, cachedValueMissing: true })
    expect(result.cells[4]).toMatchObject({ hasFormula: true, cachedValueMissing: false })
    expect(result.cells[6]).toMatchObject({ address: 'G1', present: false, type: 'blank' })
    expect(JSON.stringify(result)).not.toContain('WEBSERVICE')
  })

  it('joins shared and inline rich text without phonetic annotations and decodes OOXML escapes', () => {
    const source = xlsx(
      '<row><c t="s"><v>0</v></c><c t="inlineStr"><is><r><t xml:space="preserve"> inline </t></r><r><t>text_x000A_</t></r><rPh><t>omit</t></rPh></is></c></row>',
      {
        shared:
          '<si><r><t>Shared</t></r><r><t xml:space="preserve"> text &amp; _x005F_x0041_</t></r><rPh><t>omit</t></rPh></si>',
      },
    )
    const result = readXlsx(source, { sheet: 'Data', range: '$a$1:$b$1' })
    expect(result.range).toBe('A1:B1')
    expect(result.cells.map((cell) => cell.value)).toEqual([
      'Shared text & _x0041_',
      ' inline text\n',
    ])
  })

  it('preserves numeric date serials and raw ISO dates with the workbook date system', () => {
    const source = xlsx(
      '<row><c s="1"><v>45292</v></c><c t="d"><v>2024-01-01T00:00:00Z</v></c></row>',
      { date1904: true },
    )
    const result = readXlsx(source, { sheet: 'Data', range: 'A1:B1' })
    expect(result.dateSystem).toBe('1904')
    expect(result.cells.map((cell) => [cell.type, cell.value])).toEqual([
      ['number', '45292'],
      ['date', '2024-01-01T00:00:00Z'],
    ])
  })

  it('supports absolute internal relationship targets and alternate namespace prefixes', () => {
    const source = xlsx('<row><c><v>5</v></c></row>')
    withXml(source, 'xl/_rels/workbook.xml.rels', (value) =>
      value.replace('Target="worksheets/sheet1.xml"', 'Target="/xl/worksheets/sheet1.xml"'),
    )
    withXml(source, 'xl/workbook.xml', (value) =>
      value.replace('xmlns:r=', 'xmlns:q=').replace('r:id=', 'q:id='),
    )
    expect(readXlsx(source, { sheet: 'Data', range: 'A1' }).cells[0]?.value).toBe('5')
  })

  it('returns empty dimensions and sparse blank cell selections without allocation from dimension claims', () => {
    const source = xlsx('', { dimension: 'A1:XFD1048576' })
    expect(inspectXlsx(source).sheets[0]?.dimensions).toEqual({
      range: null,
      firstRow: null,
      lastRow: null,
      firstColumn: null,
      lastColumn: null,
      cellCount: 0,
    })
    expect(readXlsx(source, { sheet: 'Data', range: 'XFD1048576' }).cells).toMatchObject([
      { address: 'XFD1048576', row: 1048576, column: 16384, present: false },
    ])
    expect(readXlsx(source, { sheet: 'Data', range: 'A1:A500' }).cells).toHaveLength(500)
  })

  it.each([
    'A1:A501',
    'A1:XFD1048576',
    'XFE1',
    'A1048577',
    'A0',
    'A01',
    'B2:A1',
    'A:A',
    '1:2',
    'Data!A1',
    'A1:',
    ':A1',
    'A1:B2:C3',
    '',
    'A1#',
  ])('rejects invalid or unbounded range %s', (range) => {
    expectCode(() => readXlsx(xlsx(''), { sheet: 'Data', range }), 'INVALID_RANGE')
  })

  it('requires a known exact sheet name', () => {
    expectCode(() => readXlsx(xlsx(''), { sheet: 'Other', range: 'A1' }), 'INVALID_RANGE')
  })

  it.each([
    'https://example.invalid/sheet.xml',
    '../secret.xml',
    'worksheets/../sheet.xml',
    'worksheets/%2e%2e/secret.xml',
    '\\server\\file.xml',
    '//server/file.xml',
    'worksheets/sheet.xml#fragment',
  ])('rejects unsafe part targets %s', (target) => {
    const source = xlsx('')
    withXml(source, 'xl/_rels/workbook.xml.rels', (value) =>
      value.replace('Target="worksheets/sheet1.xml"', `Target="${target}"`),
    )
    expectCode(() => inspectXlsx(source), 'INVALID_OFFICE')
  })

  it('rejects external worksheet/shared-string relationships but ignores unrelated external hyperlinks', () => {
    const source = xlsx('')
    withXml(source, 'xl/_rels/workbook.xml.rels', (value) =>
      value.replace('Id="rId1"', 'Id="rId1" TargetMode="External"'),
    )
    expectCode(() => inspectXlsx(source), 'INVALID_OFFICE')
    const shared = xlsx('', { shared: '<si><t>X</t></si>' })
    withXml(shared, 'xl/_rels/workbook.xml.rels', (value) =>
      value.replace('Id="rId2"', 'Id="rId2" TargetMode="External"'),
    )
    expectCode(() => readXlsx(shared, { sheet: 'Data', range: 'A1' }), 'INVALID_OFFICE')
    expect(
      inspectXlsx(
        xlsx('', {
          extraRelations: `<Relationship Id="link" Type="${REL}/hyperlink" Target="https://example.invalid" TargetMode="External"/>`,
        }),
      ).sheets,
    ).toHaveLength(1)
  })

  it.each([
    '<row r="1"><c r="A1"/><c r="A1"/></row>',
    '<row r="1"><c r="B1"/><c r="A1"/></row>',
    '<row r="2"/><row r="1"/>',
    '<row r="1"><c r="A2"/></row>',
    '<row r="1048577"/>',
    '<row><c r="XFE1"/></row>',
    '<row><c t="unknown"/></row>',
    '<row><c><v>1</v><v>2</v></c></row>',
  ])('rejects invalid or duplicate cell coordinates and structures %s', (data) => {
    expectCode(() => inspectXlsx(xlsx(data)), 'INVALID_OFFICE')
  })

  it.each([
    '<row><c t="s"><v>0</v></c></row>',
    '<row><c t="b"><v>2</v></c></row>',
    '<row><c><v>NaN</v></c></row>',
    '<row><c><v>Infinity</v></c></row>',
  ])('rejects malformed stored values on reads %s', (data) => {
    expectCode(() => readXlsx(xlsx(data), { sheet: 'Data', range: 'A1' }), 'INVALID_OFFICE')
  })

  it('enforces the 50,000 parsed-cell limit', () => {
    const rows = Array.from(
      { length: 50_001 },
      (_, index) => `<row r="${index + 1}"><c/></row>`,
    ).join('')
    expectCode(() => inspectXlsx(xlsx(rows)), 'OFFICE_LIMIT')
  })

  it('enforces the 50-sheet limit before reading worksheets', () => {
    const source = xlsx('')
    const sheets = Array.from(
      { length: 51 },
      (_, index) => `<sheet name="Sheet${index}" r:id="id${index}"/>`,
    ).join('')
    const relations = Array.from(
      { length: 51 },
      (_, index) =>
        `<Relationship Id="id${index}" Type="${REL}/worksheet" Target="worksheets/sheet${index}.xml"/>`,
    ).join('')
    source.set(
      'xl/workbook.xml',
      encode(`<workbook xmlns="${SHEET}" xmlns:r="${REL}"><sheets>${sheets}</sheets></workbook>`),
    )
    source.set(
      'xl/_rels/workbook.xml.rels',
      encode(`<Relationships xmlns="${PACKAGE_REL}">${relations}</Relationships>`),
    )
    for (let i = 0; i < 51; i++)
      source.set(
        `xl/worksheets/sheet${i}.xml`,
        encode(`<worksheet xmlns="${SHEET}"><sheetData/></worksheet>`),
      )
    expectCode(() => inspectXlsx(source), 'OFFICE_LIMIT')
  })
})

describe('OOXML security and error hygiene', () => {
  it.each([
    '<!DOCTYPE w:document [<!ENTITY payload "secret">]>',
    '<!DOCTYPE w:document SYSTEM "https://do-not-fetch.invalid/entity">',
    '<!ENTITY payload SYSTEM "file:///etc/passwd">',
  ])('rejects document and entity declarations %s', (declaration) => {
    const source = docx(paragraph('secret'))
    withXml(source, 'word/document.xml', (value) => declaration + value)
    expectCode(() => readDocx(source, {}), 'INVALID_OFFICE')
  })

  it.each(['&unknown;', '<w:bad>', '</w:p>', '&illegal', '&#0;'])(
    'rejects malformed XML without leaking content: %s',
    (text) => {
      try {
        readDocx(docx(paragraph(`PRIVATE-MAIL-${text}`)), {})
      } catch (error) {
        expect(error).toBeInstanceOf(OfficeParseError)
        expect(String(error)).not.toContain('PRIVATE-MAIL')
        return
      }
      throw new Error('Expected malformed XML to fail')
    },
  )

  it('rejects wrong namespaces, malformed UTF-8, and unsupported encodings', () => {
    expectCode(
      () =>
        inspectDocx(
          entries({ 'word/document.xml': '<document><body><p>text</p></body></document>' }),
        ),
      'INVALID_OFFICE',
    )
    const source = docx('')
    source.set('word/document.xml', new Uint8Array([0xff]))
    expectCode(() => inspectDocx(source), 'INVALID_OFFICE')
    const unsupported = docx('')
    withXml(
      unsupported,
      'word/document.xml',
      (value) => '<?xml version="1.0" encoding="ISO-8859-1"?>' + value,
    )
    expectCode(() => inspectDocx(unsupported), 'INVALID_OFFICE')
  })

  it('bounds attribute work before a large start tag has been accumulated', () => {
    const attributes = Array.from({ length: 101 }, (_, index) => `a${index}="x"`).join(' ')
    expectCode(() => inspectDocx(docx(`<w:p ${attributes}/>`)), 'OFFICE_LIMIT')
  })

  it('rejects missing required parts with sanitized errors', () => {
    expectCode(() => inspectDocx(new Map()), 'INVALID_OFFICE')
    expectCode(() => inspectXlsx(new Map()), 'INVALID_OFFICE')
  })
})
