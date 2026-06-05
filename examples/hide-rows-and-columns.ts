import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import hideRowsAndColumns, { type HideRowsAndColumnsSheetOptions } from "../src/index.js";

// Hide both rows and columns in the same sheet (1-based indices).
const data: SheetData = [
  [
    { value: "A", fontWeight: "bold" },
    { value: "B (HIDDEN)", fontWeight: "bold" },
    { value: "C", fontWeight: "bold" },
  ],
  ["a1", "b1", "c1"],
  ["a2", "b2", "c2"],
  ["a3 (HIDDEN row)", "b3", "c3"],
  ["a4", "b4", "c4"],
  ["a5 (HIDDEN row)", "b5", "c5"],
  ["a6", "b6", "c6"],
];

const sheetOptions: SheetOptions<any> & HideRowsAndColumnsSheetOptions = {
  sheet: "Hidden rows + cols",
  hiddenRows: [4, 6], // 1-based: row 4 and row 6
  hiddenColumns: [2], // 1-based: column B
};

const output = new URL("./hide-rows-and-columns.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [hideRowsAndColumns] }).toFile(output);

console.log(`Wrote ${output}`);
