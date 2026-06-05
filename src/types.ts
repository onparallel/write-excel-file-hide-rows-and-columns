export type HiddenRange = number | { from: number; to: number };

export interface HideRowsAndColumnsSheetOptions {
  hiddenRows?: HiddenRange[];
  hiddenColumns?: HiddenRange[];
}
