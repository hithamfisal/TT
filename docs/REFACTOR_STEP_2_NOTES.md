# Refactor Step 2 — Workbook Parsing Extraction

This step moves only workbook parsing and ticket grouping logic out of `src/pages/Home.tsx` into:

- `src/lib/parseWorkbook.ts`

Moved functions:

- `groupTickets`
- `parseSiteOrder`
- `parseRows`

Behavior intentionally preserved:

- Excel sheet selection rules
- Site order detection rules
- Ticket field aliases
- Date/time conversion behavior
- RCA classification usage
- Unique ticket grouping behavior
- KPI formulas
- Export behavior
- Dashboard visual design

`Home.tsx` now imports:

```ts
import { parseRows } from "../lib/parseWorkbook";
```

Exports, charts, filters, and UI components were not moved in this step.

Recommended validation after extraction:

```bash
yarn install
yarn run check
yarn build
```
