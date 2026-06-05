# @onparallel/write-excel-file-hide-rows-and-columns

Hide rows and columns custom feature for [`write-excel-file`](https://www.npmjs.com/package/write-excel-file).

Adds support for marking individual rows and columns as **hidden** in the generated `.xlsx`, without requiring a fork of the base package. Hidden rows/columns are preserved in the file (Excel users can right-click → "Unhide" to reveal them); they are not deleted.

## Install

```sh
npm install write-excel-file @onparallel/write-excel-file-hide-rows-and-columns
```

`write-excel-file` is a peer dependency (`^4.0.0`). The package is published publicly under the `@onparallel` scope — no authentication required to install.

## Usage

Register the feature when calling `writeXlsxFile()`. `write-excel-file/node`'s built-in `SheetOptions` does not know about `hiddenRows` / `hiddenColumns`, so intersect it with the `HideRowsAndColumnsSheetOptions` type exported by this package:

```ts
import writeXlsxFile, { type SheetOptions } from 'write-excel-file/node'
import hideRowsAndColumns, {
	type HideRowsAndColumnsSheetOptions
} from '@onparallel/write-excel-file-hide-rows-and-columns'

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Sheet1',
	hiddenRows: [3, 5, { from: 10, to: 20 }], // hides rows 4, 6, and 11-21 (0-based)
	hiddenColumns: [1, { from: 4, to: 6 }] // hides columns B and E-G (0-based)
}

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile('out.xlsx')
```

If you use the intersection in many places, alias it locally: `type MySheetOptions = SheetOptions<any> & HideRowsAndColumnsSheetOptions`.

Why not module augmentation? `write-excel-file/node` re-exports its types with `export type { ... }`, which TypeScript does not allow downstream packages to augment.

## API

Both `hiddenRows` and `hiddenColumns` accept the same shape: an array of `HiddenRange`, where each entry is either a single **0-based index** or a `{ from, to }` range (inclusive, also 0-based).

| Type                           | Example                | Meaning                             |
| ------------------------------ | ---------------------- | ----------------------------------- |
| `number`                       | `5`                    | Hide a single row/column at index 5 |
| `{ from: number, to: number }` | `{ from: 10, to: 20 }` | Hide rows/columns 10 through 20     |

```ts
export type HiddenRange = number | { from: number; to: number }

export interface HideRowsAndColumnsSheetOptions {
	hiddenRows?: HiddenRange[]
	hiddenColumns?: HiddenRange[]
}
```

**Indexing:** 0-based, matching `write-excel-file`'s data array indices (row 0 = first row, column 0 = column A). The feature converts to the 1-based numbering OOXML uses internally.

**Limits:** Excel's hard maxima — 1,048,576 rows and 16,384 columns — are enforced. `from > to` and non-integer/negative indices throw.

## Combining with custom column widths

If you also pass `columns` with `width` to `writeXlsxFile()`, this feature will merge `<col hidden="1"/>` entries into the existing `<cols>` block rather than replacing it. Widths set by the user are preserved.

```ts
const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
	sheet: 'Sheet1',
	columns: [{ width: 20 }, { width: 30 }, {}],
	hiddenColumns: [2] // hides column C; A and B keep their custom widths
}
```

## How it works

The feature hooks into `xl/worksheets/sheet{id}.xml` and:

- **For hidden columns**: builds a `<cols>` block with `<col min="X" max="Y" hidden="1"/>` entries (1-based indices, consecutive indices are coalesced into ranges) and inserts it in the canonical position between `<sheetFormatPr>` and `<sheetData>` — or appends to an existing `<cols>` if the user also passed custom widths.
- **For hidden rows**: adds the `hidden="1"` attribute to existing `<row r="N">` elements that `write-excel-file` already emits. Rows with no data have no `<row>` element in the XML and are silently skipped (Excel still treats them as hidden when "unhide" is applied).

See [`src/hideRowsAndColumns.ts`](./src/hideRowsAndColumns.ts) for the implementation.

## Examples

The [`examples/`](./examples) folder contains runnable TypeScript scripts that emit `.xlsx` files next to themselves. Open the generated files in Excel/Numbers/LibreOffice to confirm rows and columns are hidden as expected.

```sh
npm install
npx tsx examples/hide-rows.ts        # run a single example
npm run examples                      # run them all
```

See [`examples/README.md`](./examples/README.md) for the full list.

## Development

```sh
npm install
npm test          # vitest
npm run typecheck # tsc --noEmit
npm run build     # emit dist/ via tsc
```

## License

[MIT](./LICENSE)
