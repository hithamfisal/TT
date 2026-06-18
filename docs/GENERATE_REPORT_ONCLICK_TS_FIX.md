# Generate Report onClick TypeScript Fix

Fixed the remaining TypeScript error by wrapping `handleGenerateManagedReport` in a mouse-event-safe callback:

```tsx
onClick={() => void handleGenerateManagedReport()}
```

This prevents React from passing the mouse event as the report type argument.
