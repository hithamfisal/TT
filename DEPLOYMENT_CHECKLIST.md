# Deployment Checklist — Tickets Dashboard

## Before Build

- [ ] Confirm no real customer data exists in `public/` templates.
- [ ] Confirm `public/google-sheets-config.txt` has no live production links unless intentionally required.
- [ ] Confirm `public/microsoft-graph-config.txt` has no client ID unless intentionally required.
- [ ] Confirm `.htaccess` exists in the project root.
- [ ] Confirm `public/.htaccess` exists so it is copied into `dist/`.
- [ ] Confirm `vite.config.ts` has `sourcemap: false`.
- [ ] Confirm no Manus/debug packages exist in `package.json`.

## Local Build Commands

Run from the project root:

```powershell
yarn install
yarn run check
yarn build
```

## After Build

Check `dist/` contains only deployment output, for example:

```text
index.html
assets/
.htaccess
*.xlsx report templates if required by export functions
*.png theme/logo assets
```

## Must Not Be In Deployment ZIP

Do not include:

```text
src/
node_modules/
package.json
vite.config.ts
tsconfig.json
.env
.git/
.gitignore
README drafts
server/
logs
*.map
real uploaded workbooks
local-excel-backup/
sample-test-workbooks/
```

## Deployment ZIP Creation

After successful build:

```powershell
cd dist
Compress-Archive -Path * -DestinationPath ..\tickets-dashboard-dist-secured.zip -Force
```

## cPanel / Namecheap Upload

- [ ] Open cPanel File Manager.
- [ ] Go to the target domain `public_html` folder.
- [ ] Backup old deployment.
- [ ] Delete old dashboard files if needed.
- [ ] Upload and extract `tickets-dashboard-dist-secured.zip`.
- [ ] Confirm `.htaccess` exists in the deployed folder.
- [ ] Open the dashboard URL.
- [ ] Upload a valid workbook.
- [ ] Test filters, tabs, exports, theme toggle, and Clear Saved Dashboard Data.

## Final Security Verification

- [ ] Directory listing is blocked.
- [ ] SPA refresh works on all routes.
- [ ] Browser console has no Manus/debug/runtime messages.
- [ ] Source maps are not present.
- [ ] Upload validation rejects unsupported files.
- [ ] Clear Saved Dashboard Data works.
- [ ] PPT/PDF/Excel/PNG exports still work.
- [ ] Dark/light theme still works.

