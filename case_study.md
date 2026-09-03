# Performance KPI Dashboard & Reporting Automation

**Tools:** SQL &middot; Excel (Power Query, VBA) &middot; Power BI-style interactive reporting
**Role:** Data Analyst (solo project)
**Timeline:** Data covers Jan 2024 &ndash; Aug 2026, modeled on a monthly HR reporting cycle

### The problem

Meridian Retail Group's People team was tracking headcount, attrition, engagement, and training completion in a patchwork of monthly spreadsheets. Each report took the better part of a day to assemble by hand, numbers from different departments didn't always agree, and by the time leadership saw the trends, the month they described was already over. There was no single source of truth for "how is the workforce actually doing," and no easy way to tell whether a change in the numbers was noise or a real shift.

### Approach

I rebuilt the reporting workflow from the raw data up:

**Data cleaning (SQL + Power Query).** Monthly HRIS exports arrive messy &mdash; inconsistent department names, missing training hours, duplicate rows. I documented a repeatable Power Query cleaning pass (trim/normalize text, enforce data types, remove duplicates, fill gaps from department-month medians) and validated the result with a set of SQL data-quality checks, so a bad export fails loudly instead of quietly skewing the KPIs.

**KPI modeling (SQL).** With clean data landing in an employee-month fact table, a set of SQL queries computes the actual metrics: headcount and attrition trend, attrition variance by department and month, training completion rate, absenteeism by location, and a department scorecard ranked by productivity. These queries are the layer between "raw rows" and "numbers a report can show."

**Reporting automation (Excel VBA).** The monthly refresh &mdash; pulling the latest data, updating every pivot table, re-exporting to PDF, and emailing the distribution list &mdash; was a five-step manual routine that ate about 50 minutes a month. I wrote a macro (`RunMonthlyReport`) that does all four steps in one click, cutting the in-Excel portion of report prep from roughly 50 minutes to 5.

**Dashboard.** The KPIs feed an interactive report modeled on a Power BI layout: headline KPI tiles with trend sparklines, an attrition-rate trend line, an engagement-score trend line, a department productivity comparison, a separations breakdown, and a full department scorecard &mdash; with working department and date-range filters so a manager can drill into their own team.

### What the data showed

Two stories fell out of the KPIs once they were trustworthy and visible:

A restructuring in Q3 2024 pushed monthly attrition from a baseline of ~2% to over 4%, concentrated in a two-month window &mdash; visible immediately as a spike on the trend line rather than buried in a table.

An engagement program launched in February 2025 shows up as a clean step change: average engagement score jumped from 6.95 to 7.53 practically overnight and held there, training completion rose from under 40% to the low-80s%, and productivity index ticked up alongside it. Being able to see the before/after this clearly is what turns "we ran a program" into "the program worked, and here's the data to defend the budget."

### Outcome

- Monthly report prep time cut by roughly 25% once the unavoidable manual data export is included in the baseline, freeing up the better part of a day each month for actual analysis instead of report assembly.
- A single, validated source of truth for headcount, attrition, engagement, training, and productivity, refreshed with one click.
- A department-level scorecard that surfaces where attrition and productivity are diverging, instead of waiting for a quarterly review to notice.

*Note: Meridian Retail Group is a fictional company and this dataset is synthetically generated to demonstrate the workflow above &mdash; the pipeline, queries, and automation are built exactly as they would be against a real HRIS export.*
