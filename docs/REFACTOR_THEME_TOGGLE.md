# Theme Toggle Update

Added a dashboard theme selector in the hero topbar.

## Files added

- `public/dark.png`
- `public/light.png`
- `docs/REFACTOR_THEME_TOGGLE.md`

## Files updated

- `src/pages/Home.tsx`
- `src/index.css`

## Behavior

The toggle switches the dashboard hero, upload preview image, and report ribbon background between:

- Dark theme image: `/dark.png`
- Light theme image: `/light.png`

The light theme also adjusts the hero overlay, text color, and topbar controls for readability.
