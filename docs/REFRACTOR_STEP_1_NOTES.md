# Refactor Step 1 Notes

This full project ZIP was created from the original uploaded source ZIP and includes the first safe refactor only.

## Added

- `src/types/dashboard.ts`
- `src/lib/dateUtils.ts`
- `src/lib/rcaRules.ts`

## Updated

- `src/pages/Home.tsx`
  - Imports shared dashboard types from `src/types/dashboard.ts`.
  - Imports date, duration, field, and site helper functions from `src/lib/dateUtils.ts`.
  - Imports RCA classification functions/constants from `src/lib/rcaRules.ts`.
  - Removed duplicated local type/helper/RCA blocks.
- `package.json`
  - Added `fflate` as a direct dependency because template exports use it dynamically.
- `tsconfig.json`
  - Updated include/path settings for this root-level Vite project structure.

## Removed as unused/backup files

- `src/pages/Home - Copy.tsx`
- `src/pages/Home2.tsx`
- `src/pages/home-tsx-modified.txt`
- `src/assets/nascologo - Copy.png`
- `src/assets/nglogo - Copy.png`
- Root-level workbook files that are not used by the app runtime:
  - `Network Performance Report.xlsx`
  - `Raw EOA.xlsx`
  - `Raw SOA.xlsx`
- `src/const.ts` because it was unused and referenced a missing `@shared/const` module.
- `public/.gitkeep`

## Not changed

- KPI formulas
- Upload/parsing behavior
- Export behavior
- Dashboard visual design
- Chart/table JSX

## Recommended test commands

```powershell
yarn install
yarn run check
yarn build
```

If `yarn run check` only runs TypeScript, `yarn build` is still recommended to validate Vite bundling.
