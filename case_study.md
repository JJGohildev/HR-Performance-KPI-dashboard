# HR Performance and Reporting Automation: Case Study

I built this project to explore workforce metrics in Power BI and make the steps behind recurring reporting easier to follow. It combines data preparation, KPI calculations, report design, and an Excel automation routine.

**Tools:** Power BI, DAX, SQL, Excel, Power Query, and VBA  
**Project type:** Personal analytics project  
**Data period:** January 2024 to August 2026

This personal project uses simulated employee records for Meridian Retail Group, a fictional company.

## Questions I focused on

- How does workforce size change alongside hiring and exits?
- When do attrition and absenteeism increase?
- How do engagement and training completion change over time?
- Which reporting steps can be made repeatable with Excel VBA?

## Preparing the data

The project includes employee records, departments, a monthly date table, employee-month observations, and summary KPI tables. Keeping their different levels of detail clear is important: an employee count and an average of monthly rates answer different questions.

The SQL and Excel files document preparation steps for inconsistent department names, missing values, and duplicate records. The SQL queries also include data-quality checks. These should be rerun after source changes rather than assumed to pass on every refresh.

## Building the report

The native Power BI file has four pages: Overview, Workforce Trends, Department Scorecard, and Attrition & Exits.

### Overview

The first page brings together headcount, attrition, engagement, training completion, and productivity. Two trend charts add context to the summary cards.

![Power BI Overview page with all departments selected](screenshots/HR_Overview.png)

The overview combines workforce KPIs with monthly trends. Some cards use different reporting periods, so I would make those periods explicit before using the report for recurring HR reviews.

### Workforce Trends

This page compares hiring and terminations with attrition, shows absenteeism over time, and breaks down training completion by department. I used separate chart areas so the reader can follow each metric without switching report pages.

![Power BI Workforce Trends page with all departments selected](screenshots/HR_Workforce_Trends.png)

The other pages provide department comparisons and a closer look at exits. The [Power BI report](pbix/Meridian_HR_KPI.pbix) contains the full set of views.

## What the trends show

In the simulated data, attrition rises above 4% around the third quarter of 2024 before falling again. Engagement moves from roughly 7.0 to around 7.5 in early 2025. Training completion also shifts upward across departments during that period.

These patterns correspond to scenarios designed into the dataset. They provide useful examples for comparing periods and departments, but they do not prove that an engagement program caused a business improvement.

Absenteeism has several distinct peaks. Before recommending an action in a real setting, I would investigate the departments involved, the reporting definitions, and possible seasonal effects. A visible peak is a starting point for investigation, not an explanation on its own.

## Reporting automation

The repository includes a VBA procedure called RunMonthlyReport. It is designed to refresh workbook data and pivots, update the report date, and export the summary as a PDF. An optional step uses Outlook to email the file when enabled in the workbook configuration.

This routine demonstrates how recurring steps can be grouped into a repeatable process. It does not refresh the Power BI report or replace the initial source-data export. The workbook configuration and local environment still need to be checked before running it.

## Next improvements

I would clarify the reporting period on each KPI card, replace automatically generated chart subtitles with shorter labels, and record a measure-by-measure comparison with SQL under matching filters. For automation, I would test refresh failures and capture a repeatable timing baseline before reporting an efficiency gain.

[Back to the project overview](README.md)
