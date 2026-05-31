# Table Width and Theme Background Focus

This update keeps the existing dashboard layout unchanged and focuses only on:

- table width / horizontal fit behavior
- dark/light theme background fitting

Updated files:
- src/pages/Home.tsx
- src/index.css

Changes:
- Tables now use the available dashboard width while keeping resizable column widths.
- Ticket and performance tables keep horizontal scrolling only when needed.
- Table cards can use the full viewport width instead of being restricted to 1320px.
- Global table minimum width is no longer forced to 2280px for all tables.
- Dark and light theme backgrounds now use consistent cover sizing, top-centered positioning, and no-repeat behavior.
- KPI ribbon/background image uses the same cover behavior.
- Upload preview image uses top-centered cover behavior.
