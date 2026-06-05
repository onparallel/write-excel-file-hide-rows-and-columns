import { describe, it, expect } from 'vitest'

import hideRowsAndColumns from './hideRowsAndColumns.js'
import type { HideRowsAndColumnsSheetOptions } from './types.js'

const transformSheetXml =
	hideRowsAndColumns.files!.transform!['xl/worksheets/sheet{id}.xml']!.transform!

function transform(xml: string, sheetOptions: HideRowsAndColumnsSheetOptions): string {
	return transformSheetXml(xml, sheetOptions as never, { sheetIndex: 0, sheetId: '1' })
}

const emptyWorksheet = '<worksheet><sheetData></sheetData></worksheet>'

describe('hideRowsAndColumns — pass-through', () => {
	it('returns XML unchanged when no options are provided', () => {
		expect(transform(emptyWorksheet, {})).toBe(emptyWorksheet)
	})

	it('returns XML unchanged with empty arrays', () => {
		expect(transform(emptyWorksheet, { hiddenRows: [], hiddenColumns: [] })).toBe(emptyWorksheet)
	})
})

describe('hideRowsAndColumns — columns', () => {
	it('inserts <cols> with a single hidden column', () => {
		expect(transform(emptyWorksheet, { hiddenColumns: [2] })).toBe(
			'<worksheet>' +
				'<cols><col min="3" max="3" hidden="1"/></cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})

	it('coalesces consecutive column indices into a single <col>', () => {
		expect(transform(emptyWorksheet, { hiddenColumns: [2, 3, 4] })).toBe(
			'<worksheet>' +
				'<cols><col min="3" max="5" hidden="1"/></cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})

	it('keeps non-consecutive column indices as separate <col> entries', () => {
		expect(transform(emptyWorksheet, { hiddenColumns: [0, 5] })).toBe(
			'<worksheet>' +
				'<cols><col min="1" max="1" hidden="1"/><col min="6" max="6" hidden="1"/></cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})

	it('accepts range objects for columns', () => {
		expect(transform(emptyWorksheet, { hiddenColumns: [{ from: 4, to: 6 }] })).toBe(
			'<worksheet>' +
				'<cols><col min="5" max="7" hidden="1"/></cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})

	it('mixes index numbers and ranges, deduplicating overlap', () => {
		expect(transform(emptyWorksheet, { hiddenColumns: [1, { from: 4, to: 6 }, 5] })).toBe(
			'<worksheet>' +
				'<cols>' +
				'<col min="2" max="2" hidden="1"/>' +
				'<col min="5" max="7" hidden="1"/>' +
				'</cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})

	it('appends to an existing <cols> block (custom widths preserved)', () => {
		const xml =
			'<worksheet>' +
			'<cols><col min="1" max="1" width="20" customWidth="1"/></cols>' +
			'<sheetData></sheetData>' +
			'</worksheet>'
		expect(transform(xml, { hiddenColumns: [2] })).toBe(
			'<worksheet>' +
				'<cols>' +
				'<col min="1" max="1" width="20" customWidth="1"/>' +
				'<col min="3" max="3" hidden="1"/>' +
				'</cols>' +
				'<sheetData></sheetData>' +
				'</worksheet>'
		)
	})
})

describe('hideRowsAndColumns — rows', () => {
	const xmlWithRows =
		'<worksheet>' +
		'<sheetData>' +
		'<row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c></row>' +
		'<row r="2"><c r="A2" t="inlineStr"><is><t>B</t></is></c></row>' +
		'<row r="3"><c r="A3" t="inlineStr"><is><t>C</t></is></c></row>' +
		'</sheetData>' +
		'</worksheet>'

	it('adds hidden="1" to a single targeted row', () => {
		expect(transform(xmlWithRows, { hiddenRows: [1] })).toBe(
			'<worksheet>' +
				'<sheetData>' +
				'<row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c></row>' +
				'<row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>B</t></is></c></row>' +
				'<row r="3"><c r="A3" t="inlineStr"><is><t>C</t></is></c></row>' +
				'</sheetData>' +
				'</worksheet>'
		)
	})

	it('handles multiple rows including ranges', () => {
		const result = transform(xmlWithRows, { hiddenRows: [0, { from: 2, to: 2 }] })
		expect(result).toContain('<row r="1" hidden="1">')
		expect(result).toContain('<row r="2">')
		expect(result).toContain('<row r="3" hidden="1">')
	})

	it('preserves existing row attributes like ht and customHeight', () => {
		const xml =
			'<worksheet><sheetData>' +
			'<row r="5" ht="25" customHeight="1"><c r="A5"/></row>' +
			'</sheetData></worksheet>'
		expect(transform(xml, { hiddenRows: [4] })).toBe(
			'<worksheet><sheetData>' +
				'<row r="5" hidden="1" ht="25" customHeight="1"><c r="A5"/></row>' +
				'</sheetData></worksheet>'
		)
	})

	it('does not match a row number that is a prefix of another (5 should not affect 50)', () => {
		const xml =
			'<worksheet><sheetData>' +
			'<row r="5"><c/></row>' +
			'<row r="50"><c/></row>' +
			'</sheetData></worksheet>'
		expect(transform(xml, { hiddenRows: [4] })).toBe(
			'<worksheet><sheetData>' +
				'<row r="5" hidden="1"><c/></row>' +
				'<row r="50"><c/></row>' +
				'</sheetData></worksheet>'
		)
	})

	it('silently skips rows that are not present in the XML', () => {
		const xml = '<worksheet><sheetData><row r="1"><c/></row></sheetData></worksheet>'
		expect(transform(xml, { hiddenRows: [9] })).toBe(xml)
	})
})

describe('hideRowsAndColumns — combined', () => {
	it('applies both column inserts and row attribute changes in one pass', () => {
		const xml =
			'<worksheet>' +
			'<sheetData>' +
			'<row r="1"><c/></row>' +
			'<row r="2"><c/></row>' +
			'</sheetData>' +
			'</worksheet>'
		expect(transform(xml, { hiddenRows: [0], hiddenColumns: [1] })).toBe(
			'<worksheet>' +
				'<cols><col min="2" max="2" hidden="1"/></cols>' +
				'<sheetData>' +
				'<row r="1" hidden="1"><c/></row>' +
				'<row r="2"><c/></row>' +
				'</sheetData>' +
				'</worksheet>'
		)
	})
})

describe('hideRowsAndColumns — validation', () => {
	it('throws on negative row index', () => {
		expect(() => transform(emptyWorksheet, { hiddenRows: [-1] })).toThrow(/>= 0/)
	})

	it('throws on negative column index', () => {
		expect(() => transform(emptyWorksheet, { hiddenColumns: [-3] })).toThrow(/>= 0/)
	})

	it('throws on non-integer index', () => {
		expect(() => transform(emptyWorksheet, { hiddenRows: [1.5] })).toThrow(/integer/)
	})

	it('throws on from > to', () => {
		expect(() => transform(emptyWorksheet, { hiddenRows: [{ from: 10, to: 5 }] })).toThrow(
			/from > to/
		)
	})

	it('throws when row index exceeds Excel maximum', () => {
		expect(() => transform(emptyWorksheet, { hiddenRows: [1_048_576] })).toThrow(/maximum row/)
	})

	it('throws when column index exceeds Excel maximum', () => {
		expect(() => transform(emptyWorksheet, { hiddenColumns: [16_384] })).toThrow(/maximum column/)
	})

	it('throws on malformed range object', () => {
		expect(() =>
			transform(emptyWorksheet, {
				hiddenRows: [{ from: 1 } as unknown as { from: number; to: number }]
			})
		).toThrow(/number or \{ from, to \}/)
	})
})
