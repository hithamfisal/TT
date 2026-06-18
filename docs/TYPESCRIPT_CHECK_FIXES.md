# TypeScript Check Fixes

This package fixes the `yarn run check` errors reported in V4.

Updated files:
- `src/pages/Home.tsx`
- `src/lib/ticketAnalytics.ts`
- `src/const.ts`

Fixes included:
- Removed broken `@shared/const` dependency by adding local constants.
- Added missing `TicketRecord` type import.
- Added null guard before `Number.isFinite(hours)`.
- Added explicit callback parameter types for Google Sheets header mapping.
- Guarded optional file names before regex checks.
- Forced `customSite` and `customRca` values to boolean.
- Added missing `MANUAL_TICKET_EXPORT_HEADERS` constant.
- Typed manual ticket export header mapping.
- Fixed Google Sheets load button click handler.
