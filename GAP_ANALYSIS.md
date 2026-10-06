# Grades Audit – Excel vs. Database gap analysis

Source view: `STUHCT.VW_GRADES_AUDIT_REPORT` (Oracle `BDEVL`), 147,860 rows, 2,643 students, 759 courses.
Reference: `CVE-Grades_Audit_report_for_SPRING_202520_as_of_2-Feb-26.xlsx` (39 students x 68 course columns).

## Dashboard layout (page "Grade Audit")
- Slicers: Campus, Major Code, Program, Catalog Term
- Matrix (replicates the Excel grid): rows = Campus, HCT Id, Name, Catalog Term, credit columns, Academic Standing; columns = `COURSE_NAME`; value = `GRADE`. Comments, Program Description, Advisor are in the cell tooltip (Power BI freezes all row headers, so wide fields would block horizontal scrolling). Major/Program are slicers.

## Excel columns not in the view
| Excel column | Status |
|---|---|
| Row 2 semester bands (`CVE-SEM-1..4`, `SUMMER-1`) | Not in DB – ignored as instructed |
| Row 3 (VLOOKUP to external *Live Block Scheduling Report*) | Not in DB – dropped |
| K `Prior Term GPA` | **Missing** |
| L `Block Code` | **Missing** |
| M `Class Code` | **Missing** |
| Course display order (study-plan sequence) | **Missing** – no sequence column, courses sort A→Z |
| Combined column `EGN-2712 or EGN-2713` | No equivalency logic in DB – shown as two separate columns |

## Data issues found
- **No CVE data in the DB**: only campus `ADB` exists; major `CVE` and all 39 Excel students (campus `DBA`) are absent. Connection points to the DEV database (`DEVDX-SCAN.../BDEVL`) – switch to PROD to reconcile.
- `ADVISOR` is NULL for 100% of rows (Excel has advisors).
- Grades: no `IP` / `IP-R (x)` / `N/A` / `P` values in DB (Excel has them); DB has `W`, `NG` not seen in Excel.
- 95% of `GRADE` values are NULL (planned/not-taken courses – displayed as blank cells, same as Excel).
- 169 student+course duplicates (same course flagged both `CORE` and `ELEC`, e.g. `CIS-4203`). The matrix shows one grade; flag for the view owner.
- `CORE_ELECTIVE_IND` NULL on 490 rows.
- Excel `Academic Standing` uses codes (`GS`), DB uses text (`Good Standing`).

## Rebuild
`python scripts/build_grades_audit.py <input.pbix> Grades_Audit_Dashboard.pbix` (rewrites only the report layout).
