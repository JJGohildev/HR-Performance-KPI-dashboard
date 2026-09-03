/* =====================================================================
   Performance KPI Dashboard & Reporting Automation
   SQL layer: schema + KPI queries feeding the Power BI-style dashboard
   Dialect: ANSI SQL / SQL Server flavored (works on Snowflake/Postgres
   with minor tweaks, noted inline)
   ===================================================================== */

-- ---------------------------------------------------------------------
-- 1. SCHEMA (mirrors the cleaned tables produced by the Power Query step)
-- ---------------------------------------------------------------------
CREATE TABLE dim_employee (
    EmployeeID          VARCHAR(10)   PRIMARY KEY,
    FullName             VARCHAR(100),
    Gender               VARCHAR(20),
    Department            VARCHAR(50),
    JobLevel              VARCHAR(30),
    Location               VARCHAR(50),
    HireDate                DATE,
    TerminationDate          DATE NULL,
    TerminationReason         VARCHAR(50) NULL
);

CREATE TABLE dim_department (
    Department           VARCHAR(50) PRIMARY KEY,
    TargetHeadcount        INT,
    Division                 VARCHAR(30)
);

CREATE TABLE fact_monthly_employee (
    EmployeeID           VARCHAR(10),
    MonthKey               DATE,          -- first day of month
    Department              VARCHAR(50),
    JobLevel                  VARCHAR(30),
    Location                    VARCHAR(50),
    PerformanceScore              DECIMAL(4,2),   -- 1.00 - 5.00
    EngagementScore                 DECIMAL(4,2), -- 1.00 - 10.00
    TrainingHours                     DECIMAL(5,1),
    AbsenceDays                         INT,
    OvertimeHours                         DECIMAL(5,1),
    ProductivityIndex                       DECIMAL(6,1),
    IsNewHire                                 BIT,
    CONSTRAINT PK_fact_monthly_employee PRIMARY KEY (EmployeeID, MonthKey),
    CONSTRAINT FK_fact_employee FOREIGN KEY (EmployeeID) REFERENCES dim_employee(EmployeeID)
);


-- ---------------------------------------------------------------------
-- 2. HEADCOUNT & ATTRITION TREND (company-wide, monthly)
--    -> feeds the "Headcount" and "Attrition Rate" trend line on the dashboard
-- ---------------------------------------------------------------------
WITH monthly_headcount AS (
    SELECT
        MonthKey,
        COUNT(DISTINCT EmployeeID)                              AS HeadcountEnd,
        SUM(CASE WHEN IsNewHire = 1 THEN 1 ELSE 0 END)          AS NewHires
    FROM fact_monthly_employee
    GROUP BY MonthKey
),
monthly_terms AS (
    SELECT
        DATEFROMPARTS(YEAR(TerminationDate), MONTH(TerminationDate), 1) AS MonthKey,
        COUNT(*) AS Terminations
    FROM dim_employee
    WHERE TerminationDate IS NOT NULL
    GROUP BY DATEFROMPARTS(YEAR(TerminationDate), MONTH(TerminationDate), 1)
)
SELECT
    h.MonthKey,
    h.HeadcountEnd,
    h.NewHires,
    COALESCE(t.Terminations, 0)                                         AS Terminations,
    ROUND(100.0 * COALESCE(t.Terminations, 0) /
          NULLIF(h.HeadcountEnd + COALESCE(t.Terminations, 0), 0), 2)   AS AttritionRatePct
FROM monthly_headcount h
LEFT JOIN monthly_terms t ON t.MonthKey = h.MonthKey
ORDER BY h.MonthKey;


-- ---------------------------------------------------------------------
-- 3. ATTRITION RATE BY DEPARTMENT WITH MONTH-OVER-MONTH VARIANCE
--    -> flags the departments/months driving the Q3-2024 restructuring spike
-- ---------------------------------------------------------------------
WITH dept_month AS (
    SELECT
        Department,
        MonthKey,
        COUNT(DISTINCT EmployeeID) AS HeadcountEnd
    FROM fact_monthly_employee
    GROUP BY Department, MonthKey
),
dept_terms AS (
    SELECT
        Department,
        DATEFROMPARTS(YEAR(TerminationDate), MONTH(TerminationDate), 1) AS MonthKey,
        COUNT(*) AS Terminations
    FROM dim_employee
    WHERE TerminationDate IS NOT NULL
    GROUP BY Department, DATEFROMPARTS(YEAR(TerminationDate), MONTH(TerminationDate), 1)
),
attrition AS (
    SELECT
        d.Department,
        d.MonthKey,
        ROUND(100.0 * COALESCE(t.Terminations, 0) / NULLIF(d.HeadcountEnd, 0), 2) AS AttritionRatePct
    FROM dept_month d
    LEFT JOIN dept_terms t ON t.Department = d.Department AND t.MonthKey = d.MonthKey
)
SELECT
    Department,
    MonthKey,
    AttritionRatePct,
    AttritionRatePct - LAG(AttritionRatePct) OVER (PARTITION BY Department ORDER BY MonthKey) AS MoM_Variance_pp
FROM attrition
ORDER BY Department, MonthKey;


-- ---------------------------------------------------------------------
-- 4. AVG PERFORMANCE & ENGAGEMENT SCORE TREND (before/after training program)
--    -> supports the "trend and variance analysis" bullet; splits the
--       timeline around the Feb-2025 engagement program rollout
-- ---------------------------------------------------------------------
SELECT
    MonthKey,
    CASE WHEN MonthKey >= '2025-02-01' THEN 'Post-Program' ELSE 'Pre-Program' END AS ProgramPeriod,
    ROUND(AVG(PerformanceScore), 2)   AS AvgPerformanceScore,
    ROUND(AVG(EngagementScore), 2)    AS AvgEngagementScore,
    ROUND(AVG(ProductivityIndex), 1)  AS AvgProductivityIndex
FROM fact_monthly_employee
GROUP BY MonthKey
ORDER BY MonthKey;

-- Variance summary: pre vs post program averages (single-row comparison)
SELECT
    ROUND(AVG(CASE WHEN MonthKey <  '2025-02-01' THEN EngagementScore END), 2) AS AvgEngagement_Pre,
    ROUND(AVG(CASE WHEN MonthKey >= '2025-02-01' THEN EngagementScore END), 2) AS AvgEngagement_Post,
    ROUND(AVG(CASE WHEN MonthKey >= '2025-02-01' THEN EngagementScore END)
        - AVG(CASE WHEN MonthKey <  '2025-02-01' THEN EngagementScore END), 2) AS Engagement_Delta
FROM fact_monthly_employee;


-- ---------------------------------------------------------------------
-- 5. TRAINING COMPLETION RATE BY DEPARTMENT (completion = >= 4 hrs/month)
-- ---------------------------------------------------------------------
SELECT
    Department,
    MonthKey,
    COUNT(*)                                                          AS Headcount,
    SUM(CASE WHEN TrainingHours >= 4.0 THEN 1 ELSE 0 END)             AS CompletedTraining,
    ROUND(100.0 * SUM(CASE WHEN TrainingHours >= 4.0 THEN 1 ELSE 0 END)
          / COUNT(*), 2)                                              AS TrainingCompletionRatePct
FROM fact_monthly_employee
GROUP BY Department, MonthKey
ORDER BY Department, MonthKey;


-- ---------------------------------------------------------------------
-- 6. ABSENTEEISM RATE (% of ~21 working days/month) BY LOCATION
-- ---------------------------------------------------------------------
SELECT
    Location,
    MonthKey,
    ROUND(AVG(AbsenceDays), 2)                     AS AvgAbsenceDays,
    ROUND(100.0 * AVG(AbsenceDays) / 21.0, 2)       AS AbsenteeismRatePct
FROM fact_monthly_employee
GROUP BY Location, MonthKey
ORDER BY Location, MonthKey;


-- ---------------------------------------------------------------------
-- 7. DEPARTMENT SCORECARD (latest month) — the table behind the dashboard's
--    "department performance" matrix, ranked by a blended score
-- ---------------------------------------------------------------------
WITH latest_month AS (
    SELECT MAX(MonthKey) AS MonthKey FROM fact_monthly_employee
)
SELECT
    f.Department,
    COUNT(*)                                     AS Headcount,
    ROUND(AVG(f.PerformanceScore), 2)            AS AvgPerformanceScore,
    ROUND(AVG(f.EngagementScore), 2)             AS AvgEngagementScore,
    ROUND(AVG(f.ProductivityIndex), 1)           AS AvgProductivityIndex,
    ROUND(100.0 * SUM(CASE WHEN f.TrainingHours >= 4.0 THEN 1 ELSE 0 END) / COUNT(*), 1) AS TrainingCompletionPct,
    RANK() OVER (ORDER BY AVG(f.ProductivityIndex) DESC)  AS ProductivityRank
FROM fact_monthly_employee f
CROSS JOIN latest_month lm
WHERE f.MonthKey = lm.MonthKey
GROUP BY f.Department
ORDER BY ProductivityRank;


-- ---------------------------------------------------------------------
-- 8. TERMINATION REASON BREAKDOWN (voluntary vs involuntary split)
--    -> feeds the donut chart on the dashboard
-- ---------------------------------------------------------------------
SELECT
    TerminationReason,
    CASE WHEN TerminationReason LIKE 'Voluntary%' THEN 'Voluntary'
         WHEN TerminationReason LIKE 'Involuntary%' THEN 'Involuntary'
         ELSE 'Other' END                        AS ReasonCategory,
    COUNT(*)                                       AS Terminations,
    ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS PctOfTotal
FROM dim_employee
WHERE TerminationDate IS NOT NULL
GROUP BY TerminationReason
ORDER BY Terminations DESC;


-- ---------------------------------------------------------------------
-- 9. DATA QUALITY CHECK (used while validating the Power Query cleanup)
--    -> confirms no orphan fact rows / duplicate keys made it past cleaning
-- ---------------------------------------------------------------------
SELECT 'Orphan fact rows (no matching employee)' AS Check_Name, COUNT(*) AS Issue_Count
FROM fact_monthly_employee f
LEFT JOIN dim_employee e ON e.EmployeeID = f.EmployeeID
WHERE e.EmployeeID IS NULL

UNION ALL

SELECT 'Duplicate (EmployeeID, MonthKey) keys', COUNT(*) - COUNT(DISTINCT EmployeeID + CAST(MonthKey AS VARCHAR(10)))
FROM fact_monthly_employee

UNION ALL

SELECT 'Fact rows after termination date', COUNT(*)
FROM fact_monthly_employee f
JOIN dim_employee e ON e.EmployeeID = f.EmployeeID
WHERE e.TerminationDate IS NOT NULL AND f.MonthKey > DATEFROMPARTS(YEAR(e.TerminationDate), MONTH(e.TerminationDate), 1);
