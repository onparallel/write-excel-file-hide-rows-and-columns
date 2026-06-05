import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from '../src/index.js'

// Hide both rows and columns in the same sheet.
const data: SheetData = [
	[
		{ value: 'A', fontWeight: 'bold' },
		{ value: 'B (HIDDEN)', fontWeight: 'bold' },
		{ value: 'C', fontWeight: 'bold' }
	],
	['a1', 'b1', 'c1'],
	['a2 (HIDDEN row)', 'b2', 'c2'],
	['a3', 'b3', 'c3'],
	['a4 (HIDDEN row)', 'b4', 'c4'],
	['a5', 'b5', 'c5']
]

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Hidden rows + cols',
	hiddenRows: [2, 4],
	hiddenColumns: [1]
}

const output = new URL('./hide-rows-and-columns.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output)

console.log(`Wrote ${output}`)
