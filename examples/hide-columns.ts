import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from '../src/index.js'

// Hide columns B and D (0-based indices 1 and 3).
const data: SheetData = [
	[
		{ value: 'A (visible)', fontWeight: 'bold' },
		{ value: 'B (HIDDEN)', fontWeight: 'bold' },
		{ value: 'C (visible)', fontWeight: 'bold' },
		{ value: 'D (HIDDEN)', fontWeight: 'bold' },
		{ value: 'E (visible)', fontWeight: 'bold' }
	],
	['a1', 'b1', 'c1', 'd1', 'e1'],
	['a2', 'b2', 'c2', 'd2', 'e2']
]

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Hidden columns',
	hiddenColumns: [1, 3]
}

const output = new URL('./hide-columns.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output)

console.log(`Wrote ${output}`)
