import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from '../src/index.js'

// Hide rows 2 and 4 (0-based). They remain in the file — Excel users can right-click → Unhide.
const data: SheetData = [
	[{ value: 'Row label', fontWeight: 'bold' }],
	['Row 0 (visible)'],
	['Row 1 (visible)'],
	['Row 2 (HIDDEN)'],
	['Row 3 (visible)'],
	['Row 4 (HIDDEN)'],
	['Row 5 (visible)']
]

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Hidden rows',
	hiddenRows: [3, 5] // rows at 0-based index 3 and 5
}

const output = new URL('./hide-rows.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output)

console.log(`Wrote ${output}`)
