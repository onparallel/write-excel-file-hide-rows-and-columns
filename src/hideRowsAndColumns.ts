import {
	insertElementMarkupAccordingToOrderOfSiblings,
	getOrderOfSiblings,
	findElement,
	appendMarkupInsideElement
} from 'write-excel-file/utility'
import type { Feature } from 'write-excel-file/node'

import normalizeRanges from './normalizeRanges.js'
import type { HideRowsAndColumnsSheetOptions } from './types.js'

// Excel hard limits (XLSX spec).
const MAX_ROWS = 1_048_576
const MAX_COLUMNS = 16_384

// `Feature<any>` rather than `Feature<unknown>` so the feature is assignable to
// `Feature<FileContent>` for any `FileContent` (Buffer/Blob/etc) chosen by the caller.
// This feature only transforms worksheet XML and never reads or writes file content,
// so the `FileContent` parameter is irrelevant here.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const hideRowsAndColumns: Feature<any> = {
	files: {
		transform: {
			'xl/worksheets/sheet{id}.xml': {
				transform: (xml, sheetOptions) => {
					const { hiddenRows, hiddenColumns } = sheetOptions as HideRowsAndColumnsSheetOptions

					const rows = normalizeRanges(hiddenRows, 'row', MAX_ROWS)
					const columns = normalizeRanges(hiddenColumns, 'column', MAX_COLUMNS)

					if (columns.intervals.length > 0) {
						xml = applyHiddenColumns(xml, columns.intervals)
					}

					if (rows.indices.length > 0) {
						xml = applyHiddenRows(xml, new Set(rows.indices))
					}

					return xml
				}
			}
		}
	}
}

export default hideRowsAndColumns

function applyHiddenColumns(xml: string, intervals: Array<[number, number]>): string {
	const colMarkup = intervals
		.map(([min, max]) => `<col min="${min}" max="${max}" hidden="1"/>`)
		.join('')

	// `write-excel-file` only emits `<cols>` when at least one column has a `width`.
	// If the user combined this feature with custom widths, append to the existing block;
	// otherwise insert a fresh `<cols>` at its canonical position.
	const existingCols = findElement(xml, 'cols')
	if (existingCols) {
		return appendMarkupInsideElement(xml, existingCols, colMarkup)
	}

	const order = getOrderOfSiblings('xl/worksheets/sheet{id}.xml', 'worksheet')
	if (!order) {
		throw new Error('write-excel-file did not return an order of siblings for `worksheet`')
	}
	return insertElementMarkupAccordingToOrderOfSiblings(
		xml,
		`<cols>${colMarkup}</cols>`,
		order,
		'worksheet'
	)
}

function applyHiddenRows(xml: string, hiddenRowSet: Set<number>): string {
	// Match the opening tag of any `<row r="N" ...>` (with or without extra attributes like
	// `ht="..."` and `customHeight="1"`). The lookahead ensures we only match at the end of
	// the `r="N"` attribute — never inside another attribute value or token.
	// Cell text content cannot collide: `write-excel-file` HTML-escapes `<` in strings.
	return xml.replace(/<row r="(\d+)"(?=[\s/>])/g, (match, rowNumStr: string) => {
		const rowNum = Number(rowNumStr)
		if (hiddenRowSet.has(rowNum)) {
			return `<row r="${rowNumStr}" hidden="1"`
		}
		return match
	})
}
