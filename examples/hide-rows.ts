import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from "../src/index.js";

// Hide rows 4 and 6 (1-based, matching Excel's visible row numbers).
// Hidden rows remain in the file — Excel users can right-click → Unhide.
const data: SheetData = [
  [{ value: "Row label", fontWeight: "bold" }],
  ["Row 2 (visible)"],
  ["Row 3 (visible)"],
  ["Row 4 (HIDDEN)"],
  ["Row 5 (visible)"],
  ["Row 6 (HIDDEN)"],
  ["Row 7 (visible)"],
];

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
  sheet: "Hidden rows",
  hiddenRows: [4, 6],
};

const output = new URL("./hide-rows.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output);

console.log(`Wrote ${output}`);
