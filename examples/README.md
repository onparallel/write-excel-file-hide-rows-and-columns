# Examples

Runnable TypeScript scripts. Each one writes its own `.xlsx` next to the script.

| Example                                                      | Demonstrates                                                                 |
| ------------------------------------------------------------ | ---------------------------------------------------------------------------- |
| [`hide-rows.ts`](./hide-rows.ts)                             | Hide individual rows by 0-based index.                                       |
| [`hide-columns.ts`](./hide-columns.ts)                       | Hide individual columns by 0-based index.                                    |
| [`hide-rows-and-columns.ts`](./hide-rows-and-columns.ts)     | Hide both rows and columns in the same sheet.                                |
| [`hide-with-ranges.ts`](./hide-with-ranges.ts)               | Use `{ from, to }` ranges to hide contiguous blocks compactly.               |
| [`hide-with-custom-widths.ts`](./hide-with-custom-widths.ts) | Combine with `columns: [{ width }]` — user widths preserved on visible cols. |

## Run

```sh
# Single example
npx tsx examples/hide-rows.ts

# All examples
npm run examples
```
