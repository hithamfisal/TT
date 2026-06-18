# Security Hardening Report — Tickets Dashboard

## Scope

Project: Follow-Up Sheets / DMR Tickets Dashboard

Security hardening was applied according to the uploaded hardening prompt for the Tickets Dashboard. The dashboard business logic, KPI calculations, tabs, filters, charts, exports, dark/light theme behavior, and upload workflow were preserved.

## Completed Hardening Actions

### 1. Manus / Debug Runtime Removal

Status: Completed

Removed references and dependencies related to:

- `manus`
- `__manus__`
- `vite-plugin-manus-runtime`
- `@builder.io/vite-plugin-jsx-loc`
- Manus log/runtime/debug collector traces

Removed/cleaned:

- `vite-plugin-manus-runtime` from `package.json`
- `@builder.io/vite-plugin-jsx-loc` from `package.json`
- stale package-manager references
- unused `src/components/ManusDialog.tsx`
- unused `src/const.ts`
- debug/server-only runtime folders

Verification grep returned no Manus runtime/debug references in source files.

### 2. Clean Production Vite Configuration

Status: Completed

`vite.config.ts` was replaced with a clean React/Vite production configuration:

- `sourcemap: false`
- clean `@` and `@assets` aliases
- no Microsoft proxy middleware
- no debug/runtime plugins
- static build output to `dist`

### 3. `.htaccess` Security File

Status: Completed

Added:

- `.htaccess`
- `public/.htaccess`

The file includes:

- directory listing disabled
- SPA fallback to `index.html`
- `X-Content-Type-Options: nosniff`
- `X-Frame-Options: DENY`
- `Referrer-Policy`
- `Permissions-Policy`
- Content Security Policy
- blocked access to `.env`, `.log`, `.map`, `.db`, backup, and config-style files

### 4. CSP Meta Tag

Status: Completed

Added CSP meta tag to `index.html`.

The CSP protects the frontend while allowing the existing browser-side dashboard functions:

- React/Vite runtime
- Excel parsing
- PDF/PPT/PNG export
- local images/data blobs
- optional Google/Microsoft workbook links if manually configured

External Google font links were removed from `index.html` to reduce external requests.

### 5. Public File / Sample Data Review

Status: Completed with template retention

Removed:

- local backup workbook folder
- duplicate test workbook folders
- root-level workbook copies
- duplicate copied logo files
- Apps Script helper page
- public Google Apps Script helper text file

Sanitized:

- `public/google-sheets-config.txt`
- `public/microsoft-graph-config.txt`

Both config files now contain blank placeholders instead of live customer links or IDs.

Retained in `public/` because the dashboard export functions depend on them:

- `DMR_Monthly_Report.xlsx`
- `EOA_DMR_Monthly_Report.xlsx`
- `SOA_DMR_Monthly_Report.xlsx`
- `COA_DMR_Monthly_Report.xlsx`
- `WOA_DMR_Monthly_Report.xlsx`
- `Network_Performance_Report.xlsx`

These appear to be report templates, not uploaded raw data. They should still be reviewed by the project owner before client deployment.

### 6. Browser Storage Safety Controls

Status: Completed

Added visible button:

- `Clear Saved Dashboard Data`

It clears dashboard-specific saved data:

- saved workbook snapshot
- saved manual ticket rows
- dashboard-related IndexedDB databases if present
- selected dashboard session error key if present

It does not clear unrelated browser data.

User confirmation:

- `Saved dashboard data cleared successfully.`

### 7. Upload Privacy Note

Status: Completed

Added upload privacy note near the upload area:

> Uploaded files are processed in your browser. No files are sent to a server by this dashboard.

### 8. File Validation

Status: Completed

Added validation for:

- allowed extensions: `.xlsx`, `.xls`, `.csv`
- empty files
- max file size: 50 MB
- missing/invalid ticket data after parsing
- corrupted or unsupported files

User-friendly error examples:

- `Invalid file type. Please upload .xlsx, .xls, or .csv only.`
- `One or more selected files are empty. Please upload a valid workbook.`
- `The uploaded file could not be processed. Please check the file format and required columns.`

### 9. Safe Error Handling

Status: Completed for upload path

Upload parsing errors now show friendly user messages.

Detailed errors are logged only in development mode using `import.meta.env.DEV`.

### 10. Production Source Maps

Status: Completed

`vite.config.ts` includes:

```ts
build: {
  sourcemap: false
}
```

No `.map` files should be generated in the final production `dist` build.

## Security Notes / Owner Review Required

### Client-side login password

The existing dashboard includes a client-side login gate. Client-side passwords are visible in frontend source/builds and should not be treated as real access security.

Recommendation before production:

- replace client-side-only login with server-side authentication, or
- deploy behind cPanel/basic-auth/VPN, or
- treat the current login only as a convenience UI lock.

### Report templates

The retained Excel report templates should be manually reviewed by the owner to confirm they contain no real customer data.

## Build Verification

The sandbox environment did not have Yarn installed and `npm install` timed out before dependencies could be installed. Because of that, I could not run the final local build inside this environment.

Please run locally:

```powershell
yarn install
yarn run check
yarn build
```

Then create the deployment ZIP from `dist/`.



## Google Sheets CSP Fix

The dashboard uses the Google Visualization API (`https://docs.google.com/spreadsheets/.../gviz/tq`) through a JSONP script callback to read TT-History sheets in the browser. The CSP was updated to allow Google Sheets script sources while keeping other protections in place.

Allowed script sources added:

```text
https://docs.google.com
https://spreadsheets.google.com
https://script.google.com
https://script.googleusercontent.com
```

The Google Sheet must still be shared as `Anyone with the link can view`.
