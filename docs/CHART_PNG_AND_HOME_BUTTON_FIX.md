# Chart PNG + Home Button Fix

This update fixes the section/chart navigation polish requested after the default-collapsed dashboard update.

## Changes

- Added a **Home** button in the top dashboard action bar.
- Home returns the user to the welcome upload screen by clearing the active dashboard data from memory.
- The saved workbook snapshot remains available, so the user can still use **Continue previous workbook**.
- Fixed missing PNG buttons on charts that were inside sections collapsed by default.
- Chart PNG buttons are now created even while a section is collapsed, so they are visible immediately after expanding.
- The chart PNG button effect now re-runs when sections are expanded/collapsed.

## Files updated

- `src/pages/Home.tsx`
