# TypeScript Check Final Fixes

Fixed the remaining `yarn run check` issues:

- Removed unsupported `sourceRowNumber` and `sourceRegionKey` fields from manual ticket draft objects.
- Changed `MANUAL_TICKET_EXPORT_HEADERS` from readonly tuple to mutable string array.
- Spread `MANUAL_TICKET_EXPORT_HEADERS` when passing it to `XLSX.utils.aoa_to_sheet`.
- Updated manual export header mapping typing.
