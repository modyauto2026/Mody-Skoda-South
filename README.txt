
CHANDRA SOLD VEHICLE RETENTION DASHBOARD
========================================

WHAT THIS VERSION FIXES
-----------------------
The old Excel dashboard relies on cached/modern Excel formulas such as
XLOOKUP, FILTER, UNIQUE and MAXIFS. Some machines/Excel versions can show
old or incorrect cached values.

This app does NOT use the Excel dashboard formulas for its numbers.
It reads raw SOLD_DATA and SERVICE_DATA and recalculates the results
every time you change a filter.

It also fixes the no-service issue where an empty last service date can
behave like Excel date 0 and create incorrect overdue/critical-lost values.

FIRST TIME - EXACT STEPS
------------------------
1. Right-click the downloaded ZIP.
2. Choose "Extract All".
3. Open the extracted folder.
4. Make sure these files are together:
      app.py
      Chandra.xlsx
      requirements.txt
      RUN_CHANDRA_DASHBOARD.bat
      README.txt
5. Double-click RUN_CHANDRA_DASHBOARD.
   Windows may hide ".bat", so it may only show as RUN_CHANDRA_DASHBOARD.
6. On the first run, wait while packages install.
7. The browser should open automatically.
   If not, open Chrome and type:
      http://localhost:8501
8. KEEP THE BLACK WINDOW OPEN while using the dashboard.

NEXT TIME
---------
Just double-click RUN_CHANDRA_DASHBOARD.

HOW TO UPDATE THE DATA
----------------------
Easiest option:
- Open the dashboard.
- Use "Upload updated Chandra Excel" in the left sidebar.

OR:
- Close the dashboard.
- Replace Chandra.xlsx in this app folder with the newer workbook.
- Keep the file name Chandra.xlsx.
- Start the dashboard again.

DASHBOARD TABS
--------------
1. Executive Dashboard
   Filters for Sold Year, Sold Month, Service Year, Service Month,
   Dealer/Sold Branch, City, Brand, Model, Service Location,
   Service Type and Service Category.

2. Cohort Matrix
   Shows exactly when each sold-month cohort came back for service.

3. Action Dashboard
   Recalculates Critical Lost, High Risk, Overdue and Due Soon VINs.

4. VIN Search
   Search VIN, vehicle number, customer, model, RO, city, consultant.
   Shows exact service history.

5. Data Quality
   Shows duplicates, unmatched VINs and other accuracy checks.

IMPORTANT
---------
Do NOT run the BAT file inside the ZIP preview. Extract the ZIP first.
The BAT file automatically changes to its own folder, so requirements.txt
will be found correctly.


NEW: MONTHLY AUTO-IMPORT MODE
=============================

YOU ONLY NEED TWO FILES EACH MONTH:
1. SALES data
2. SERVICE data

MONTHLY PROCESS
---------------
1. Double-click RUN_CHANDRA_DASHBOARD.
2. Open "Monthly Data Upload".
3. Upload the new month's Sales file.
4. Upload the new month's Service file.
5. Click "Import & Update Dashboard".
6. The app automatically:
   - appends the new data to historical master data
   - keeps all old months
   - removes exact duplicate rows
   - creates a backup before every import
   - recalculates retention, cohorts and action lists
   - records the import in an import log
7. Refresh the dashboard and use it normally.

MINIMUM COLUMNS
---------------
Sales file:
VIN
Sold Date

Service file:
VIN
Service Date

Your additional normal columns will be preserved automatically.

IMPORTANT
---------
Do not upload the entire historical data every month unless you want to.
The intended process is to upload ONLY the new month's Sales and Service data.
If an exact row is uploaded again, the app removes that exact duplicate.

The first master history is seeded automatically from the included Chandra.xlsx.


V2 CHANGES
==========
- Cohort Matrix now has its own summary.
- Retention colors: 80%+ Green, 50-79% Orange, below 50% Red.
- Reported / Service Month is shown at the TOP of the Cohort Matrix.
- New Lost VIN Details tab with exact VIN details and recommended action.
- Lost VIN action list can be downloaded as CSV.


V3 CHANGES
==========
1. Cohort Matrix Summary is now displayed clearly above the matrix.
2. TOTAL RETENTION % is displayed prominently in:
   - Executive Dashboard
   - Cohort Matrix Summary
3. Cohort Matrix Summary includes a TOTAL row.
4. Dashboard filters are reordered exactly as:
   Brand
   Sold Branch
   Sold Year - Month
   Service Year - Month
   Service Type
   Service Category
5. City, Model and Service Location are moved into "More Filters".


V3.1 FIX
========
Fixed Lost VIN Details KeyError for Last Service Date caused by duplicate merge column names.


V3.2 CHANGES
============
1. Executive Dashboard month-on-month retention table now displays fully without an internal scroll box.
2. Executive Dashboard charts/cards are more compact to reduce vertical space.
3. Cohort Matrix color intensity is reversed within each band:
   - Red: lowest % = strongest red; higher % fades lighter.
   - Orange: lowest % in orange band = strongest orange; higher % fades lighter.
   - Green: lowest % in green band = strongest green; higher % fades lighter.
4. Cohort Matrix height is slightly reduced for a cleaner dashboard.


V3.3 FINAL CHANGES
==================
- Executive Summary cards are smaller and consistently sized.
- Month-on-Month Retention by Sold Month Cohort is a full non-scrollable table.
- Cohort Matrix scale stays fixed from 0 to 100%.
- Red 0-49.9%, Orange 50-79.9%, Green 80-100%.
- Within each band, the lowest value is the strongest/darkest and it fades lighter as the value rises.


V3.4 FIXES
==========
- Fixed the chart height variable issue.
- Cohort Matrix is now BLUE ONLY, matching the supplied screenshot.
- Cohort Matrix scale is fixed from 0% to 100%.
- 0% is very light blue; higher percentages progressively become darker blue.
- Executive Summary cards are smaller and equal in size.
- Month-on-Month Retention table is full-width and static, with no internal scroll box.
- Retention bar chart is below the full table and also fixed to a 0-100% y-axis.


V3.5 HOTFIX
===========
Fixed: NameError: name 'height' is not defined.
Cause: compact KPI CSS braces were interpreted by the Python f-string.
The CSS braces are now escaped correctly.
The blue 0-100 Cohort Matrix and full-width non-scroll Month-on-Month table remain unchanged.


V3.6 VISUAL UPDATE
==================
- Added colored symbols/icons to Executive Summary KPI cards.
- Icons added to Cohort Summary and Lost VIN KPIs where useful.
- KPI cards are smaller, consistent, and styled closer to the supplied screenshot.
- Existing blue 0-100 Cohort Matrix remains unchanged.
- Existing full-width non-scroll Month-on-Month table remains unchanged.


V3.7
====
Added Throughput Tracking tab using SERVICE DATA ONLY, with YoY comparison and growth arrows beside growth %.


V3.8: Throughput Tracking redesigned to one-page selected-year comparison using SERVICE DATA ONLY.


V3.9 — THROUGHPUT COMPARISON IMPROVEMENTS
=========================================
- Service Category column is grouped visually with icon/name blocks for PMS, GR, Accident Repair and Inspection.
- Total row shown after each service category.
- KPI growth is green; KPI decline is red; flat is orange.
- Service Month is now multi-select.
- Selected months are applied equally to Current Service Year and Compare With Year.
  Example: Jan-Jul compares Jan-Jul 2026 against Jan-Jul 2025.
- Month-wise table and chart show only the selected comparison months.


V3.10 — SERVICE CATEGORY ICON FIX
=================================
- Replaced Windows/Chrome emoji icons with clean inline SVG dashboard icons.
- PMS = teal wrench.
- GR = orange gear.
- Accident Repair = red vehicle/impact icon.
- Inspection = blue clipboard/magnifier.
- Colors are fixed in the app and no longer depend on the operating system's emoji style.
- Applied to Throughput KPI cards and grouped Service Category blocks.


V3.11 — REFERENCE-MATCHED THROUGHPUT UI
=======================================
- Throughput Tracking page restyled to closely match the supplied reference screenshot.
- Exact supplied screenshot icon artwork is used for PMS, GR, Accident Repair, Inspection and serviced VINs.
- Navy section bars, compact KPI cards, grouped service-category table, YoY summary, chart and Key Insights are presented in the same one-page composition.
- Current Service Year / Compare With Year remain interactive.
- Service Months remain multi-select and apply identically to both years.
- Growth is green; decline is red; flat is orange.


V3.12: Service Months remains multi-select, but selected month chips are hidden so the filter stays compact and clean.


V3.13: Service Months filter is forced to a single compact dropdown height. Selected months remain hidden inside and no longer expand downward.


V3.14 — MONTH DROPDOWN FIX
==========================
- Removed the visible Streamlit multi-select chips completely.
- Service Months is now one compact dropdown/popover.
- Jan-Dec choices appear only after clicking the dropdown.
- When the dropdown is closed, no selected month chips or month list are visible.
- Multiple months can still be selected.
- The exact same months are applied to both comparison years.


V3.15 — LOST VIN & IMMEDIATE ACTION DASHBOARD
=============================================
New separate tab added. Existing tabs are retained.

Uses BOTH SOLD_DATA + SERVICE_DATA:
- Sold VINs define the customer/VIN population.
- Service history determines whether/when each VIN returned.
- As On Date allows historical status review.
- Critical Lost: >180 days overdue.
- Overdue: >0 days overdue (Critical separated above 180).
- Due Next 30 Days: due within 30 days.
- Retained: serviced and within service cycle.
- Not Yet Due: too early for service.

Includes:
- KPI cards
- Priority Action by sold-month cohort
- VIN status distribution
- Lost/Overdue by sold location
- Retention by sold location
- Retention trend by sold-month cohort
- Detailed VIN action list with CSV export
- Recommended action guidance


V3.16 HOTFIX
============
Fixed Lost VIN & Immediate Action Dashboard error:
ValueError: invalid literal for int() with base 10: 'Dec'

Sold Month Calc can now safely contain:
- numeric months such as 1, 2, 12
- month names such as Jan, Feb, Dec

The same fix is applied to the Lost VIN retention-trend cohort dates.


V3.17 — COHORT MATRIX DISPLAY
=============================
- Chart heading simplified to REPORTED / SERVICE YEAR <selected year>.
- Removed the repeated REPORTED / SERVICE MONTH axis title.
- Month-wise Cohort Summary Retention % cells are colored:
  Green = 80%+
  Orange = 50% to 79.9%
  Red = below 50%

V3.18
- Added Missed PMS VIN-level same-month comparison in Throughput Tracking.
- Added downloadable Missed PMS detail.
- Added External / Not-in-Our-Sold-Data VIN acquisition and retention.
- Added downloadable external retention and new external VIN lists.


V3.19
- Added a dedicated dropdown/filter row to External VIN Acquisition & Retention.
- Filters: Current Service Year, Compare With Year, Service Months, Service Location, Service Type, Service Category.
- External VIN KPIs, retention table, detail view, and downloads now follow these selections.

V3.20
- Fixed KPI HTML rendering by removing Markdown code-block indentation before rendering.
- This fixes raw <div> text in Lost VIN & Immediate Action Dashboard.
- KPI calculations are unchanged.

V3.21 — ASK YOUR DATA
- Added a new Ask Your Data tab.
- Local controlled question engine; no API key or internet required.
- Supports service counts, PMS/GR/Accident/Inspection, month/year/location filters,
  Missed PMS, External VINs, sold-year retention, service-location ranking, and overdue VIN lists.
- Supporting tables can be downloaded as CSV.
- Unsupported questions are not guessed.

V3.22
- Replaced the five Lost VIN & Immediate Action KPI HTML cards with native Streamlit metric cards.
- This permanently avoids visible <div> HTML text on affected Streamlit versions.
- Lost VIN calculations and filters are unchanged.
- Ask Your Data from v3.21 remains included.

V3.23
- Fixed Ask Your Data NameError by adding the missing `re` import.

V3.24 — SMP TRACKING
- Added a dedicated SMP Tracking tab.
- Seeded with the supplied SMP workbook.
- Matches SMP VINs to actual PMS service history.
- Counts unique PMS service events within each SMP validity period.
- Filters: Due Month, SMP Validity, SMP Status, Model, Service Location.
- Summary KPIs and SMP validity summary table.
- Due/Overdue VIN-level data with CSV download.
- SMP data can be replaced from inside the tab.
- First-version due cadence = 365 days.
- Services Allowed/Balance is estimated as one PMS per validity year.

V3.25 — DATA UPLOAD CENTER
- Added two separate upload modes:
  1) Full Master Upload — validates and replaces complete Sales + Service history.
  2) Monthly Incremental Upload — appends new monthly rows and removes exact duplicates.
- Full replacement requires an explicit confirmation checkbox.
- Existing master Sales and Service files are backed up automatically before replacement.
- Import history now records the import type.

V3.26 — MODY SKODA SOUTH
- Renamed package to Mody Skoda South.
- Updated the main dashboard headline branding.
- Changed the headline palette from navy/white to dark green/mint green.
- Added extra top spacing so the headline remains fully visible below the Streamlit toolbar.

V3.27 — APPROVED HEADER/FONT STYLE
- Main Python dashboard uses Segoe UI / system sans-serif typography matching the approved screenshot.
- Main title uses dark green background with mint-green text and border.
- Panel headings use the same green/mint treatment.
- Increased top spacing so the dashboard headline remains fully visible below Streamlit's toolbar.
- Main title text simplified to SOLD VEHICLE COHORT RETENTION DASHBOARD.

V3.28 — FULL TOP-TO-BOTTOM THEME
- Applied approved dark emerald background + mint-green font/borders throughout the Python app.
- Updated main heading exactly to:
  MODY SKODA SOUTH — SOLD VEHICLE COHORT RETENTION DASHBOARD
- Rechecked header spacing so the full main heading remains visible.
- Updated cards, filters, tabs, sidebar, metrics, tables and major section panels to the same theme.

V3.29 FAST — PERFORMANCE OPTIMIZATION
- Replaced eager Streamlit tabs with a tab-style page selector.
  Only the selected dashboard section now executes.
- Cached master CSV reads based on file modified time.
- Cached full master data preparation (VIN cleanup, dates, categories, cohort fields).
- Upload/replace actions still clear cache and reload the latest master automatically.
- No dashboard feature or calculation was intentionally removed.
- Main title remains: MODY SKODA SOUTH — SOLD VEHICLE COHORT RETENTION DASHBOARD

V3.30 FAST FIX
- Fixed NameError: name 'background' is not defined.
- Cause: CSS literal braces were added inside a Python f-string.
- Escaped the full dark-green theme/navigation CSS correctly.
- Retains v3.29 performance optimization and all dashboard features.

V3.31 FONT FIX
- Rechecked the whole app top to bottom.
- Forced Segoe UI on Streamlit text, sidebar, filters, buttons, metrics, tables and navigation.
- Forced Segoe UI on Plotly SVG/chart text as well.
- Removed legacy white Plotly chart background from style_fig and aligned charts to the dark emerald/mint theme.
- Main title remains exactly:
  MODY SKODA SOUTH — SOLD VEHICLE COHORT RETENTION DASHBOARD
- Retains the v3.29/v3.30 speed optimization and CSS-brace fix.

V3.32
- Removed the Ask Your Data tab/section completely.
- SMP Tracking now follows Lost VIN & Immediate Action directly.
- Fixed Plotly compatibility error from v3.31: invalid XAxis/YAxis 'titlefont' property.
- Retains fast lazy-page execution, full dark emerald/mint theme and Segoe UI typography.

V3.33 VISIBILITY FIX
- Fixed very dark/invisible navigation labels on the dark green background.
- Navigation text is now high-contrast light mint/white.
- Selected section has a brighter mint border and white text.
- Fixed download buttons and their text so CSV/Excel download options remain clearly visible.
- Fixed file uploader/Browse files text visibility.
- Retains removal of Ask Your Data and all v3.32 functionality.

V3.34 TABLE VISIBILITY FIX
- Fixed dark/invisible dataframe and summary-table text.
- Header rows use dark green with mint text.
- Data rows use dark emerald with bright white/light text.
- Added alternating row background for easier reading.
- Preserved navigation/download visibility fixes from v3.33.

V3.36 — V3.28 FONT / COLOUR RESTORE
- Restored the v3.28 Segoe UI font treatment.
- Restored v3.28 mint, white and muted-mint text colours.
- Kept dark emerald dashboard styling.
- Kept readable v3.34 tables.
- Kept the filter dropdown arrow visibility correction.

V3.37 — V3.18 FONT + COLOUR STYLE
- Applied the font/typography appearance from the supplied v3.18 app.
- Applied the v3.18 navy/blue/green/orange/red/text colour palette.
- Kept the current dashboard calculations, pages, filters and functionality unchanged.
- Kept dropdown arrows clearly visible.

V3.38 — V3.18 STYLE ACROSS ALL AREAS
- Applied the supplied v3.18 font and colour appearance throughout the full dashboard.
- Updated page background, sidebar, navigation, KPI cards, filters, dropdowns, multiselects,
  buttons, uploads, expanders, tables and chart-area palette to the v3.18 look.
- Kept all current calculations, pages, upload logic, filters and dashboard functionality unchanged.
- Retention status colours remain the v3.18 green/orange/red scheme.

V3.39 — RE-VERIFIED V3.18 STYLE
- Re-audited all CSS blocks and Plotly chart styling.
- Removed remaining emerald/mint colours from newer dashboard sections.
- Fixed navy-background text contrast, including Plotly hover labels.
- Applied v3.18 font and colour rules inside all page-specific CSS blocks.
- Kept dashboard calculations and functionality unchanged.

V3.40 — SCREENSHOT AREA FIX
- Fixed MONTH ON MONTH RETENTION table: removed dark alternating rows and dark header text.
- Fixed Throughput Tracking and Year-on-Year Summary tables: all normal rows are white with dark text.
- Total rows use the v3.18 light-blue total style.
- Fixed Throughput chart text to dark v3.18 text colour.
- Calculations and dashboard functionality unchanged.

V3.41 — Fixed NameError caused by unescaped CSS braces in v3.40. Table fixes preserved.

V3.42 — Fixed remaining unescaped v3.39 CSS braces causing runtime NameError at line 1061. Compile and AST f-string validation passed.

V3.43 — ARROW / ICON TEXT FIX
- Fixed Streamlit/Material icons showing their text names instead of arrow/icon glyphs.
- Restored Material Symbols fonts for icon elements.
- Kept Source Sans Pro for normal dashboard text.
- Preserved v3.42 runtime fixes and v3.18 colour styling.


FULL MASTER REPLACE FIX — PYTHON / STREAMLIT VERSION
---------------------------------------------------
Full Master Upload now supports:
1) ONE combined .xlsx workbook containing SOLD_DATA + SERVICE_DATA, OR
2) Separate Sales and Service .xlsx/.csv files.

After validation the app:
- Backs up the current MASTER_SALES.csv and MASTER_SERVICE.csv.
- Replaces MASTER_SALES.csv and MASTER_SERVICE.csv.
- Backs up and rebuilds Chandra.xlsx from the same new full data.
- Clears Streamlit cached data and reruns the dashboard.

This means the Python folder itself is synchronized with the replacement full master.


V3.45 FULL REPLACE OPTIONS
--------------------------
Full Master Upload now supports three independent choices:

1. Replace FULL SALES only
   - Upload only Sales Excel/CSV
   - Existing Service master remains unchanged

2. Replace FULL SERVICE only
   - Upload only Service Excel/CSV
   - Existing Sales master remains unchanged

3. Replace BOTH Sales + Service
   - Upload one combined Excel workbook, or
   - Upload separate Sales and Service files

The selected master is backed up first.
Chandra.xlsx is then synchronized with the active Sales + Service masters.


V3.46 — SIMPLIFIED FULL MASTER REPLACE
--------------------------------------
The combined/BOTH replacement option has been removed.

Full Master Upload now has only:
1. Replace FULL SALES
   - Upload Sales Excel/CSV only
   - Existing Service master stays unchanged.

2. Replace FULL SERVICE
   - Upload Service Excel/CSV only
   - Existing Sales master stays unchanged.

Each replacement is backed up first, and Chandra.xlsx is synchronized afterward.


V3.47 — INDEPENDENT MONTHLY UPLOAD
---------------------------------
Monthly Incremental Upload now matches the Full Master Upload design.

Only two choices:
1. Append MONTHLY SALES
   - Upload Sales Excel/CSV only.
   - Service master remains unchanged.

2. Append MONTHLY SERVICE
   - Upload Service Excel/CSV only.
   - Sales master remains unchanged.

There is no requirement to upload both monthly files.
The master being changed is backed up first.
Exact duplicate rows are removed automatically.
Chandra.xlsx is synchronized after each successful monthly append.

Update — Missed PMS + Throughput multi-year comparison
- Fixed Missed PMS logic so a VIN is counted as returned when it has a PMS visit in ANY selected current-year month, rather than requiring the PMS to fall in the exact same calendar month as the comparison-year PMS.
- Verified VIN MEXBPJPB6NG018493: PMS on 21-Jun-2025 and PMS on 25-Feb-2026; with Jan-Aug 2026 selected it is no longer treated as Missed PMS.
- Throughput Tracking > Compare With Year now supports multiple year selection using a compact popover like the month selector.
- Added Multi-Year Comparison Summary for all selected comparison years. The latest selected comparison year remains the primary year for the existing detailed month-wise tables/charts and Missed PMS section.

V3.48 — SPEED OPTIMIZATION (NO LOGIC/UI CHANGES)
- SMP Tracking: PMS history is now grouped once per VIN before matching against
  each SMP contract, instead of re-scanning the entire PMS table for every VIN.
  Same results, verified to match row-for-row on test data, much faster on
  large service history.
- Lost VIN Details: status classification (Critical Lost / Lost / High Risk /
  Never Serviced) now runs as one vectorized pass instead of a Python function
  call per VIN. Verified identical output.
- Action Dashboard: same vectorization applied to Action Status. Verified
  identical output.
- SMP data file is now cached by file version, so switching filters on the SMP
  Tracking tab no longer re-reads and re-cleans the Excel file each time.
- Executive Dashboard icons (serviced/PMS/GR/accident/inspection PNGs) are now
  cached instead of being re-read and re-encoded from disk on every rerun.
- No dashboard calculation, filter, column, or visual output was changed.

UPDATE — RO CLUBBING & AMOUNT FIX
- Service records are now clubbed as one workshop visit per VIN + RO Number.
- Customer / warranty / internal invoice rows under the same RO are combined.
- Labor, Parts, Total and Invoice Amount Inc Tax are summed across the linked rows.
- The clubbed service-event data is used consistently in VIN Search, PMS history, retention, throughput and related service calculations.
- Blank RO numbers remain separate to avoid unsafe merging.
