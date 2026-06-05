import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from "../src/index.js";

// Combine this feature with `columns` (custom widths). The feature appends `<col hidden="1"/>`
// entries to the existing `<cols>` block — user-defined widths are preserved.
const data: SheetData = [
  [
    { value: "A (wide)", fontWeight: "bold" },
    { value: "B (HIDDEN)", fontWeight: "bold" },
    { value: "C (wide)", fontWeight: "bold" },
  ],
  ["a1", "b1", "c1"],
  ["a2", "b2", "c2"],
];

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
  sheet: "Widths + hidden",
  columns: [{ width: 30 }, {}, { width: 30 }],
  hiddenColumns: [2], // 1-based: column B
};

const output = new URL("./hide-with-custom-widths.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output);

console.log(`Wrote ${output}`);
