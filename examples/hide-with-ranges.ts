import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from '../src/index.js'

// Use `{ from, to }` ranges to hide contiguous blocks compactly (1-based, inclusive).
const data: SheetData = [
	[{ value: 'Hidden range demo', fontWeight: 'bold' }],
	...Array.from({ length: 25 }, (_, i) => [`Row ${i + 2}`])
]

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Range syntax',
	// Hide rows 11–21 (inclusive) using a single range entry,
	// plus row 25 as a literal index. Mixing both shapes is fine.
	hiddenRows: [{ from: 11, to: 21 }, 25]
}

const output = new URL('./hide-with-ranges.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output)

console.log(`Wrote ${output}`)
