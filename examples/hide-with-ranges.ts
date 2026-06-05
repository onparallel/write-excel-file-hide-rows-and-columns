import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from '../src/index.js'

// Use `{ from, to }` ranges to hide contiguous blocks compactly.
const data: SheetData = [
	[{ value: 'Hidden range demo', fontWeight: 'bold' }],
	...Array.from({ length: 25 }, (_, i) => [`Row ${i + 1}`])
]

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Range syntax',
	// Hide rows 10–20 (inclusive, 0-based) using a single range entry,
	// plus row 24 as a literal index. Mixing both shapes is fine.
	hiddenRows: [{ from: 10, to: 20 }, 24]
}

const output = new URL('./hide-with-ranges.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output)

console.log(`Wrote ${output}`)
