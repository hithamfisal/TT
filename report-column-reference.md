# Distinct TT Report Column Reference

Confirmed from the user-provided image on 2026-05-12. The distinct report table should follow this order:

1. #
2. Site ID
3. Site Name
4. Managed Resource
5. Severity
6. Issues
7. Observation Date
8. Observation Time
9. Recovery Date
10. Recovery Time
11. Escalated for L3 Support Date
12. Escalated for L3 Support Time
13. Total Duration Days/Hours
14. TT
15. Status
16. Escalated to
17. Comments-Feedback

Implementation notes:

- Site ID and Site Name should show all affected site IDs/names for the distinct TT, not only the primary row.
- Escalated for L3 Support Date and Time are new user-added sheet columns and must be parsed by header aliases.
- CSV and Excel export should use the same column order.
