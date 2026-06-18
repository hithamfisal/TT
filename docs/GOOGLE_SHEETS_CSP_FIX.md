# Google Sheets CSP Fix

This update fixes the error:

```text
Could not read TT-History from Google Sheets. Share the file as 'Anyone with the link can view' and try again.
```

The security CSP previously allowed Google Sheets in `connect-src`, but the dashboard reads Google Sheets using the Google Visualization JSONP endpoint, which loads through a `<script>` tag. Therefore Google Sheets must also be allowed in `script-src`.

Updated files:

- `index.html`
- `.htaccess`
- `public/.htaccess`
- `SECURITY_HARDENING_REPORT.md`

Google Sheet sharing requirement remains:

```text
Share > General access > Anyone with the link > Viewer
```

The expected sheet/tab name is usually `TT-History`, unless configured otherwise in the dashboard link settings.
