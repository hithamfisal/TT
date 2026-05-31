# PPT Repair Fix

This update fixes the PowerPoint repair warning seen when opening exported Performance PPT files.

## Change

The Performance PPT charts were simplified to avoid PPTX XML that PowerPoint may repair:

- Removed unsupported/risky `plotAreaFill` chart option.
- Removed custom chart grid-line styling from PPT export.
- Removed rotated category axis labels from PPT export.
- Disabled chart data labels to reduce chart XML complexity on large site lists.
- Removed unsupported custom rounded-corner property from PPT shapes.

## Preserved

- Dashboard UI behavior.
- Excel/PDF/CSV exports.
- Ticket PPT export structure.
- Performance PPT export content and slides.
- Executive Insights and deep-dive sections.

After extracting, run:

```powershell
yarn install
yarn run check
yarn build
```

Then export `Performance_Report_All.pptx` again and open the newly generated file.
