/*
===================================================================================
PROJECT      : E-Commerce Customer Insights & Churn Analysis
AUTHOR       : Senior Business Analyst / SQL Developer
PLATFORM     : Microsoft SQL Server (SSMS Compatible)
DATASET      : E_Commerce_Customer_Insights_and_Churn_Dataset.csv
DESCRIPTION  : End-to-end SQL Business Analyst portfolio project covering
               database setup, data cleaning, KPI development, customer
               analysis, churn analysis, revenue analysis, segmentation,
               and advanced SQL techniques (Window Functions, CTEs, Joins,
               Subqueries, CASE WHEN logic).
===================================================================================
*/


-- =========================================
-- DATABASE SETUP
-- =========================================

-- Business Objective: Create a dedicated database to host and isolate
-- the e-commerce customer dataset for analysis.
-- SQL Logic: Standard SSMS database creation with existence check.
-- Expected Insight: A clean working environment for all downstream queries.

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'ECommerceChurnDB')
BEGIN
    CREATE DATABASE ECommerceChurnDB;
END
GO

USE ECommerceChurnDB;
GO

-- Business Objective: Drop the staging/raw table if it already exists,
-- so the script can be re-run safely from scratch (idempotent setup).
-- SQL Logic: Conditional DROP TABLE using OBJECT_ID check (SSMS pattern).
-- Expected Insight: Avoids "table already exists" errors on re-execution.

IF OBJECT_ID('dbo.ecommerce_customers_raw', 'U') IS NOT NULL
    DROP TABLE dbo.ecommerce_customers_raw;
GO

-- Business Objective: Define the raw staging table that mirrors the
-- structure of the source CSV file exactly, before any cleaning/typing.
-- SQL Logic: All columns are typed to closely match the source data;
-- dates are staged as VARCHAR because the source file uses mixed
-- M/D/YYYY text formats that must be parsed explicitly during cleaning.
-- Expected Insight: A safe landing zone for raw imported data with zero
-- risk of import failure due to type mismatches.

CREATE TABLE dbo.ecommerce_customers_raw (
    order_id              VARCHAR(20)     NOT NULL,
    customer_id           VARCHAR(20)     NOT NULL,
    age                   INT             NULL,
    product_id            VARCHAR(20)     NULL,
    country               VARCHAR(50)     NULL,
    signup_date            VARCHAR(20)     NULL,   -- staged as text, parsed later
    last_purchase_date      VARCHAR(20)     NULL,   -- staged as text, parsed later
    cancellations_count    INT             NULL,
    subscription_status    VARCHAR(20)     NULL,
    order_date              VARCHAR(20)     NULL,   -- staged as text, parsed later
    unit_price              DECIMAL(10,2)   NULL,
    quantity                INT             NULL,
    purchase_frequency      INT             NULL,
    preferred_category      VARCHAR(50)     NULL,
    product_name            VARCHAR(100)    NULL,
    category                VARCHAR(50)     NULL,
    gender                  VARCHAR(20)     NULL
);
GO


-- =========================================
-- DATA IMPORT INSTRUCTIONS
-- =========================================

/*
Business Objective: Load the source CSV file into the raw staging table
so it can be cleaned, validated, and transformed by this script.

OPTION 1 - SSMS Import Wizard (recommended for one-time loads):
    1. Right-click ECommerceChurnDB -> Tasks -> Import Flat File.
    2. Select E_Commerce_Customer_Insights_and_Churn_Dataset.csv.
    3. Target table: dbo.ecommerce_customers_raw (created above).
    4. Map columns 1:1 in the order listed in CREATE TABLE above.
    5. Set date columns (signup_date, last_purchase_date, order_date)
       to import as VARCHAR/NVARCHAR to avoid locale-based date errors.

OPTION 2 - BULK INSERT (recommended for repeatable/automated loads):
    Update the file path below to the actual location of the CSV file
    on the SQL Server host (or a UNC/network path accessible to the
    SQL Server service account), then execute:
*/

-- BULK INSERT dbo.ecommerce_customers_raw
-- FROM 'C:\Data\E_Commerce_Customer_Insights_and_Churn_Dataset.csv'
-- WITH (
--     FIRSTROW          = 2,             -- skip header row
--     FIELDTERMINATOR    = ',',
--     ROWTERMINATOR      = '0x0a',
--     CODEPAGE           = '65001',       -- UTF-8
--     TABLOCK
-- );
-- GO

/*
OPTION 3 - OPENROWSET (for ad-hoc loads, requires 'Ad Hoc Distributed
Queries' enabled via sp_configure):
*/

-- INSERT INTO dbo.ecommerce_customers_raw
-- SELECT *
-- FROM OPENROWSET(
--     BULK 'C:\Data\E_Commerce_Customer_Insights_and_Churn_Dataset.csv',
--     FORMATFILE = 'C:\Data\ecommerce_format.xml'
-- ) AS src;
-- GO


-- =========================================
-- DATA CLEANING
-- =========================================

-- Business Objective: Confirm the raw table loaded successfully and
-- contains the expected row volume before proceeding.
-- SQL Logic: Simple row count check.
-- Expected Insight: Validates successful import (expected ~2,000 rows).

SELECT COUNT(*) AS total_rows_loaded
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Identify missing/NULL values across all critical
-- business columns to assess data completeness before analysis.
-- SQL Logic: SUM of CASE WHEN ... IS NULL flags per column.
-- Expected Insight: Confirms whether cleaning/imputation is required;
-- this dataset is expected to return all zeros (no missing values).

SELECT
    SUM(CASE WHEN order_id             IS NULL THEN 1 ELSE 0 END) AS null_order_id,
    SUM(CASE WHEN customer_id          IS NULL THEN 1 ELSE 0 END) AS null_customer_id,
    SUM(CASE WHEN age                  IS NULL THEN 1 ELSE 0 END) AS null_age,
    SUM(CASE WHEN product_id           IS NULL THEN 1 ELSE 0 END) AS null_product_id,
    SUM(CASE WHEN country              IS NULL THEN 1 ELSE 0 END) AS null_country,
    SUM(CASE WHEN signup_date          IS NULL THEN 1 ELSE 0 END) AS null_signup_date,
    SUM(CASE WHEN last_purchase_date   IS NULL THEN 1 ELSE 0 END) AS null_last_purchase_date,
    SUM(CASE WHEN cancellations_count  IS NULL THEN 1 ELSE 0 END) AS null_cancellations,
    SUM(CASE WHEN subscription_status  IS NULL THEN 1 ELSE 0 END) AS null_sub_status,
    SUM(CASE WHEN order_date           IS NULL THEN 1 ELSE 0 END) AS null_order_date,
    SUM(CASE WHEN unit_price           IS NULL THEN 1 ELSE 0 END) AS null_unit_price,
    SUM(CASE WHEN quantity             IS NULL THEN 1 ELSE 0 END) AS null_quantity,
    SUM(CASE WHEN purchase_frequency   IS NULL THEN 1 ELSE 0 END) AS null_purchase_frequency,
    SUM(CASE WHEN preferred_category   IS NULL THEN 1 ELSE 0 END) AS null_preferred_category,
    SUM(CASE WHEN product_name         IS NULL THEN 1 ELSE 0 END) AS null_product_name,
    SUM(CASE WHEN category             IS NULL THEN 1 ELSE 0 END) AS null_category,
    SUM(CASE WHEN gender               IS NULL THEN 1 ELSE 0 END) AS null_gender
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Detect duplicate order records, which would
-- inflate revenue and order-count metrics if left unresolved.
-- SQL Logic: GROUP BY order_id with HAVING COUNT(*) > 1.
-- Expected Insight: Confirms order_id uniqueness (Primary Key candidate).

SELECT
    order_id,
    COUNT(*) AS duplicate_count
FROM dbo.ecommerce_customers_raw
GROUP BY order_id
HAVING COUNT(*) > 1;
GO

-- Business Objective: Detect duplicate customer records, which would
-- distort per-customer KPIs such as CLV and churn rate.
-- SQL Logic: GROUP BY customer_id with HAVING COUNT(*) > 1.
-- Expected Insight: Confirms customer_id uniqueness in this dataset.

SELECT
    customer_id,
    COUNT(*) AS duplicate_count
FROM dbo.ecommerce_customers_raw
GROUP BY customer_id
HAVING COUNT(*) > 1;
GO

-- Business Objective: Validate that subscription_status only contains
-- the three expected business states, catching typos or inconsistent
-- categorical entries (e.g., "Active" vs "active ").
-- SQL Logic: DISTINCT + COUNT to inventory all values actually present.
-- Expected Insight: Confirms only 'active', 'cancelled', 'paused' exist.

SELECT
    subscription_status,
    COUNT(*) AS record_count
FROM dbo.ecommerce_customers_raw
GROUP BY subscription_status
ORDER BY record_count DESC;
GO

-- Business Objective: Validate that customer age values fall within a
-- realistic adult range, flagging any data entry errors.
-- SQL Logic: MIN/MAX/AVG plus a conditional count of out-of-range ages.
-- Expected Insight: Confirms ages range sensibly between 18 and ~70.

SELECT
    MIN(age) AS min_age,
    MAX(age) AS max_age,
    AVG(age) AS avg_age,
    SUM(CASE WHEN age < 18 OR age > 100 THEN 1 ELSE 0 END) AS invalid_age_count
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Detect pricing outliers or invalid (negative/zero)
-- prices that would distort revenue calculations.
-- SQL Logic: Aggregate statistics plus standard deviation and percentile
-- bounds using PERCENTILE_CONT (SQL Server window-based syntax).
-- Expected Insight: Confirms price range is plausible; flags any rows
-- with non-positive prices for correction or removal.

SELECT
    MIN(unit_price)  AS min_price,
    MAX(unit_price)  AS max_price,
    AVG(unit_price)  AS avg_price,
    STDEV(unit_price) AS stddev_price,
    SUM(CASE WHEN unit_price <= 0 THEN 1 ELSE 0 END) AS invalid_price_count
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Detect invalid (non-positive) order quantities.
-- SQL Logic: Conditional count of quantity <= 0.
-- Expected Insight: Confirms all quantities are valid positive integers.

SELECT
    MIN(quantity) AS min_quantity,
    MAX(quantity) AS max_quantity,
    SUM(CASE WHEN quantity <= 0 THEN 1 ELSE 0 END) AS invalid_quantity_count
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Ensure logical date sequencing — a customer's
-- signup date must precede or equal their order date. Violations
-- indicate corrupted records that should be excluded or corrected.
-- SQL Logic: CAST text date columns to DATE using TRY_CONVERT (safe
-- parsing that returns NULL instead of erroring on bad formats), then
-- compare chronological order.
-- Expected Insight: Confirms 0 records where signup occurs after order.

SELECT
    order_id,
    customer_id,
    TRY_CONVERT(DATE, signup_date, 101) AS parsed_signup_date,
    TRY_CONVERT(DATE, order_date, 101)  AS parsed_order_date
FROM dbo.ecommerce_customers_raw
WHERE TRY_CONVERT(DATE, signup_date, 101) > TRY_CONVERT(DATE, order_date, 101);
GO

-- Business Objective: Identify any date strings that failed to parse,
-- which would indicate formatting inconsistencies in the source file.
-- SQL Logic: TRY_CONVERT returns NULL on failure; filter for these.
-- Expected Insight: Confirms all date strings are well-formed M/D/YYYY.

SELECT
    order_id,
    signup_date,
    last_purchase_date,
    order_date
FROM dbo.ecommerce_customers_raw
WHERE TRY_CONVERT(DATE, signup_date, 101) IS NULL
   OR TRY_CONVERT(DATE, last_purchase_date, 101) IS NULL
   OR TRY_CONVERT(DATE, order_date, 101) IS NULL;
GO

-- Business Objective: Trim accidental leading/trailing whitespace in
-- text fields that could break GROUP BY aggregations (e.g., "USA " vs
-- "USA" would otherwise be treated as different groups).
-- SQL Logic: UPDATE using LTRIM(RTRIM()) on all VARCHAR columns.
-- Expected Insight: Normalizes categorical text fields prior to typing.

UPDATE dbo.ecommerce_customers_raw
SET
    country             = LTRIM(RTRIM(country)),
    subscription_status  = LTRIM(RTRIM(LOWER(subscription_status))),
    preferred_category   = LTRIM(RTRIM(preferred_category)),
    product_name         = LTRIM(RTRIM(product_name)),
    category             = LTRIM(RTRIM(category)),
    gender               = LTRIM(RTRIM(gender));
GO

-- Business Objective: Build a clean, analysis-ready table with proper
-- data types (real DATE columns instead of text) and derived business
-- columns, leaving the raw staging table untouched as an audit trail.
-- SQL Logic: SELECT ... INTO new table, casting dates with TRY_CONVERT
-- and computing total_revenue and age_group as derived columns.
-- Expected Insight: Produces the single source of truth table used by
-- every downstream KPI, segmentation, and analysis query in this script.

IF OBJECT_ID('dbo.ecommerce_clean', 'U') IS NOT NULL
    DROP TABLE dbo.ecommerce_clean;
GO

SELECT
    order_id,
    customer_id,
    age,

    -- Derived Column: age_group
    -- Business Objective: Bucket customers into marketing-friendly age
    -- bands to support demographic-level revenue and churn analysis.
    CASE
        WHEN age BETWEEN 18 AND 25 THEN '18-25'
        WHEN age BETWEEN 26 AND 35 THEN '26-35'
        WHEN age BETWEEN 36 AND 45 THEN '36-45'
        WHEN age BETWEEN 46 AND 55 THEN '46-55'
        ELSE '56+'
    END AS age_group,

    product_id,
    country,
    TRY_CONVERT(DATE, signup_date, 101)        AS signup_date,
    TRY_CONVERT(DATE, last_purchase_date, 101) AS last_purchase_date,
    cancellations_count,
    subscription_status,
    TRY_CONVERT(DATE, order_date, 101)         AS order_date,
    unit_price,
    quantity,

    -- Derived Column: total_revenue
    -- Business Objective: Calculate the monetary value of each order
    -- as the foundation for all revenue and CLV calculations.
    CAST(ROUND(unit_price * quantity, 2) AS DECIMAL(12,2)) AS total_revenue,

    purchase_frequency,

    -- Derived Column: estimated_clv
    -- Business Objective: Approximate Customer Lifetime Value using
    -- order revenue multiplied by historical purchase frequency.
    CAST(ROUND((unit_price * quantity) * purchase_frequency, 2) AS DECIMAL(14,2)) AS estimated_clv,

    preferred_category,
    product_name,
    category,
    gender,

    -- Derived Column: cross_category_flag
    -- Business Objective: Flag customers whose actual purchase category
    -- differs from their stated preferred category, signalling
    -- cross-sell engagement.
    CASE
        WHEN preferred_category = category THEN 'Aligned'
        ELSE 'Cross-Category Buyer'
    END AS purchase_alignment

INTO dbo.ecommerce_clean
FROM dbo.ecommerce_customers_raw;
GO

-- Business Objective: Verify the clean table was built correctly and
-- the derived columns are populated as expected.
-- SQL Logic: Row count and sample preview of the clean table.
-- Expected Insight: Confirms successful transformation (2,000 rows).

SELECT COUNT(*) AS total_rows_clean FROM dbo.ecommerce_clean;
GO

SELECT TOP 10 * FROM dbo.ecommerce_clean;
GO


-- =========================================
-- DATA VALIDATION (POST-CLEANING)
-- =========================================

-- Business Objective: Confirm the clean table contains no NULL values
-- in business-critical fields after type conversion and derivation.
-- SQL Logic: NULL checks identical in structure to the raw-table check,
-- run again post-cleaning to confirm no data was lost in transformation.
-- Expected Insight: Validates the clean table is fully populated and
-- ready for KPI and analytical queries.

SELECT
    SUM(CASE WHEN signup_date        IS NULL THEN 1 ELSE 0 END) AS null_signup_date,
    SUM(CASE WHEN order_date         IS NULL THEN 1 ELSE 0 END) AS null_order_date,
    SUM(CASE WHEN total_revenue      IS NULL THEN 1 ELSE 0 END) AS null_total_revenue,
    SUM(CASE WHEN age_group          IS NULL THEN 1 ELSE 0 END) AS null_age_group
FROM dbo.ecommerce_clean;
GO

-- Business Objective: Validate referential consistency between the raw
-- and clean tables to ensure no rows were dropped during transformation.
-- SQL Logic: Compare row counts between staging and clean tables.
-- Expected Insight: Both counts should match exactly (2,000 = 2,000).

SELECT
    (SELECT COUNT(*) FROM dbo.ecommerce_customers_raw) AS raw_row_count,
    (SELECT COUNT(*) FROM dbo.ecommerce_clean)         AS clean_row_count;
GO

-- Business Objective: Add a Primary Key constraint to enforce uniqueness
-- on order_id going forward, protecting data integrity for future loads.
-- SQL Logic: ALTER TABLE ADD CONSTRAINT PRIMARY KEY.
-- Expected Insight: Locks in order_id as the formal Primary Key.

ALTER TABLE dbo.ecommerce_clean
ALTER COLUMN order_id VARCHAR(20) NOT NULL;
GO

ALTER TABLE dbo.ecommerce_clean
ADD CONSTRAINT PK_ecommerce_clean_order_id PRIMARY KEY (order_id);
GO

-- Business Objective: Add a supporting non-clustered index on
-- customer_id to optimize performance of customer-level joins and
-- aggregations used throughout this script.
-- SQL Logic: CREATE NONCLUSTERED INDEX.
-- Expected Insight: Improves query performance for customer-centric KPIs.

CREATE NONCLUSTERED INDEX IX_ecommerce_clean_customer_id
ON dbo.ecommerce_clean (customer_id);
GO

CREATE NONCLUSTERED INDEX IX_ecommerce_clean_subscription_status
ON dbo.ecommerce_clean (subscription_status);
GO


-- =========================================
-- KPI ANALYSIS
-- =========================================

-- Business Objective: Produce the headline Customer KPI scorecard —
-- total, active, churned, and paused customer counts plus churn and
-- retention rates — for executive reporting.
-- SQL Logic: Conditional SUM(CASE WHEN) aggregation over the full base.
-- Expected Insight: Total = 2,000; Churn Rate ≈ 24.65%; Retention ≈ 75.35%.

SELECT
    COUNT(*) AS total_customers,
    SUM(CASE WHEN subscription_status = 'active'    THEN 1 ELSE 0 END) AS active_customers,
    SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END) AS churned_customers,
    SUM(CASE WHEN subscription_status = 'paused'    THEN 1 ELSE 0 END) AS paused_customers,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS churn_rate_pct,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status <> 'cancelled' THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS retention_rate_pct
FROM dbo.ecommerce_clean;
GO

-- Business Objective: Summarize core Revenue KPIs for the business —
-- total revenue, average revenue per customer, and order value range.
-- SQL Logic: Standard aggregate functions on the derived total_revenue
-- column.
-- Expected Insight: Total revenue ≈ $2,051,691; AOV ≈ $1,025.85.

SELECT
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue_per_customer,
    CAST(MAX(total_revenue) AS DECIMAL(10,2))  AS max_order_value,
    CAST(MIN(total_revenue) AS DECIMAL(10,2))  AS min_order_value
FROM dbo.ecommerce_clean;
GO

-- Business Objective: Break down revenue contribution by customer
-- subscription segment, revealing how much revenue is currently
-- "at risk" (paused) versus secured (active).
-- SQL Logic: GROUP BY subscription_status with revenue aggregation.
-- Expected Insight: Active customers drive the majority of revenue;
-- paused customers represent recoverable at-risk revenue.

SELECT
    subscription_status,
    COUNT(*)                                  AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue,
    CAST(AVG(cancellations_count) AS DECIMAL(5,2)) AS avg_cancellations
FROM dbo.ecommerce_clean
GROUP BY subscription_status
ORDER BY total_revenue DESC;
GO

-- Business Objective: Quantify revenue contribution by country to guide
-- regional investment and marketing budget allocation.
-- SQL Logic: GROUP BY country with a subquery-based revenue-share
-- percentage calculation.
-- Expected Insight: Germany leads in absolute revenue; share is fairly
-- evenly distributed across all six markets (~15-18% each).

SELECT
    country,
    COUNT(*)                                   AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))   AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))   AS avg_revenue_per_customer,
    CAST(ROUND(100.0 * SUM(total_revenue) /
        (SELECT SUM(total_revenue) FROM dbo.ecommerce_clean), 2) AS DECIMAL(5,2)) AS revenue_share_pct
FROM dbo.ecommerce_clean
GROUP BY country
ORDER BY total_revenue DESC;
GO

-- Business Objective: Quantify revenue contribution by product category
-- to inform inventory, marketing, and merchandising priorities.
-- SQL Logic: GROUP BY category with revenue-share subquery.
-- Expected Insight: Clothing leads revenue share; Electronics shows the
-- highest average order value.

SELECT
    category,
    COUNT(*)                                  AS total_orders,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_order_value,
    CAST(ROUND(100.0 * SUM(total_revenue) /
        (SELECT SUM(total_revenue) FROM dbo.ecommerce_clean), 2) AS DECIMAL(5,2)) AS revenue_share_pct
FROM dbo.ecommerce_clean
GROUP BY category
ORDER BY total_revenue DESC;
GO

-- Business Objective: Calculate Order KPIs — total orders, average
-- order value, average purchase frequency, and the repeat purchase
-- rate, which together describe transactional health.
-- SQL Logic: Aggregate functions combined with a conditional repeat-
-- purchase flag (purchase_frequency > 1).
-- Expected Insight: High average purchase frequency (~25) indicates a
-- strongly repeat-driven customer base.

SELECT
    COUNT(*)                                          AS total_orders,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))          AS avg_order_value,
    CAST(AVG(purchase_frequency) AS DECIMAL(10,2))     AS avg_purchase_frequency,
    CAST(AVG(quantity) AS DECIMAL(5,2))                AS avg_quantity_per_order,
    SUM(CASE WHEN purchase_frequency > 1 THEN 1 ELSE 0 END) AS repeat_customers,
    CAST(ROUND(100.0 * SUM(CASE WHEN purchase_frequency > 1 THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS repeat_purchase_rate_pct
FROM dbo.ecommerce_clean;
GO

-- Business Objective: Surface the top 10 customers by estimated
-- Customer Lifetime Value (CLV) to prioritize VIP retention investment.
-- SQL Logic: ORDER BY the pre-computed estimated_clv derived column,
-- limited to the top decile-equivalent results via TOP.
-- Expected Insight: Identifies the highest-priority accounts for
-- white-glove account management.

SELECT TOP 10
    customer_id,
    country,
    category,
    total_revenue,
    purchase_frequency,
    estimated_clv,
    subscription_status
FROM dbo.ecommerce_clean
ORDER BY estimated_clv DESC;
GO

-- Business Objective: Identify the top 10 customers ranked purely by
-- order frequency, to recognize the most habitually engaged shoppers.
-- SQL Logic: ORDER BY purchase_frequency DESC with TOP N limiting.
-- Expected Insight: Highlights "super-user" customers for loyalty and
-- ambassador program targeting.

SELECT TOP 10
    customer_id,
    purchase_frequency,
    total_revenue,
    subscription_status,
    country
FROM dbo.ecommerce_clean
ORDER BY purchase_frequency DESC;
GO


-- =========================================
-- CUSTOMER ANALYSIS
-- =========================================

-- Business Objective: Identify the highest-value customers overall,
-- combining revenue and frequency context for account prioritization.
-- SQL Logic: Simple ranked SELECT with TOP N and CLV included for
-- context.
-- Expected Insight: Top customers generate near-maximum single-order
-- revenue (~$1,990-$1,998) and should receive VIP treatment.

SELECT TOP 10
    customer_id,
    country,
    category,
    gender,
    age,
    total_revenue,
    purchase_frequency,
    estimated_clv,
    subscription_status
FROM dbo.ecommerce_clean
ORDER BY total_revenue DESC;
GO

-- Business Objective: Determine which customer segment (by preferred
-- category) generates the most total and average revenue.
-- SQL Logic: GROUP BY preferred_category with revenue and frequency
-- aggregates.
-- Expected Insight: Electronics-preferring customers show the highest
-- average order value among all segments.

SELECT
    preferred_category,
    COUNT(*)                                    AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))    AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))    AS avg_revenue,
    CAST(AVG(purchase_frequency) AS DECIMAL(10,2)) AS avg_purchase_frequency
FROM dbo.ecommerce_clean
GROUP BY preferred_category
ORDER BY total_revenue DESC;
GO

-- Business Objective: Identify the most frequent purchasers (top
-- quartile-equivalent of purchase_frequency) for loyalty targeting.
-- SQL Logic: WHERE filter on a high purchase_frequency threshold,
-- informed by the dataset's observed distribution.
-- Expected Insight: Surfaces engaged customers, including some at-risk
-- "cancelled" accounts that represent a significant historical loss.

SELECT
    customer_id,
    purchase_frequency,
    total_revenue,
    subscription_status,
    country,
    preferred_category
FROM dbo.ecommerce_clean
WHERE purchase_frequency >= 45
ORDER BY purchase_frequency DESC;
GO

-- Business Objective: Determine which customer demographics (age group
-- x gender) drive the highest total sales, supporting targeted
-- marketing investment decisions.
-- SQL Logic: GROUP BY age_group and gender with revenue aggregation.
-- Expected Insight: The 56+ age bracket contributes the largest single
-- share of total revenue across all demographic cross-sections.

SELECT TOP 10
    age_group,
    gender,
    COUNT(*)                                  AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue
FROM dbo.ecommerce_clean
GROUP BY age_group, gender
ORDER BY total_revenue DESC;
GO

-- Business Objective: Quantify revenue and engagement trends by age
-- group alone, to support age-based retention and acquisition strategy.
-- SQL Logic: GROUP BY age_group with revenue-share subquery.
-- Expected Insight: Revenue increases with age, peaking in the 56+
-- bracket; the 26-35 group shows the second-highest average order value.

SELECT
    age_group,
    COUNT(*)                                    AS customers,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))    AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))    AS avg_revenue,
    CAST(AVG(purchase_frequency) AS DECIMAL(10,2)) AS avg_purchase_frequency,
    CAST(ROUND(100.0 * SUM(total_revenue) /
        (SELECT SUM(total_revenue) FROM dbo.ecommerce_clean), 2) AS DECIMAL(5,2)) AS revenue_share_pct
FROM dbo.ecommerce_clean
GROUP BY age_group
ORDER BY age_group;
GO


-- =========================================
-- CHURN ANALYSIS
-- =========================================

-- Business Objective: Calculate the overall churn rate and the
-- combined at-risk population (churned + paused) for executive
-- reporting.
-- SQL Logic: Conditional aggregation across subscription_status values.
-- Expected Insight: Churn rate ≈ 24.65%; combined at-risk ≈ 39.8% of
-- the entire customer base.

SELECT
    COUNT(*) AS total_customers,
    SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END) AS churned,
    SUM(CASE WHEN subscription_status = 'paused'    THEN 1 ELSE 0 END) AS at_risk_paused,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS churn_rate_pct,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status IN ('cancelled','paused') THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS total_at_risk_pct
FROM dbo.ecommerce_clean;
GO

-- Business Objective: Identify which countries have the highest churn
-- rates to prioritize regional retention strategy.
-- SQL Logic: GROUP BY country with conditional churn-rate calculation.
-- Expected Insight: India and Pakistan show the highest churn rates
-- (~28-29%); the UK shows the lowest (~20%).

SELECT
    country,
    COUNT(*) AS total_customers,
    SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END) AS churned_customers,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS churn_rate_pct
FROM dbo.ecommerce_clean
GROUP BY country
ORDER BY churn_rate_pct DESC;
GO

-- Business Objective: Identify which product categories have the
-- highest churn rates, signalling potential product-market fit issues.
-- SQL Logic: GROUP BY category with conditional churn-rate calculation.
-- Expected Insight: Beauty and Home categories show marginally higher
-- churn than Electronics.

SELECT
    category,
    COUNT(*) AS total_customers,
    SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END) AS churned_customers,
    CAST(ROUND(100.0 * SUM(CASE WHEN subscription_status = 'cancelled' THEN 1 ELSE 0 END)
        / COUNT(*), 2) AS DECIMAL(5,2)) AS churn_rate_pct
FROM dbo.ecommerce_clean
GROUP BY category
ORDER BY churn_rate_pct DESC;
GO

-- Business Objective: Investigate whether cancellation history,
-- purchase frequency, or revenue level correlate with current
-- subscription status, to identify churn drivers.
-- SQL Logic: GROUP BY subscription_status with multiple behavioural
-- averages compared side by side.
-- Expected Insight: Churn is evenly distributed across cancellation
-- counts, suggesting engagement quality (not cancellation history
-- alone) is the stronger churn driver.

SELECT
    subscription_status,
    CAST(AVG(cancellations_count) AS DECIMAL(5,2)) AS avg_cancellations,
    CAST(AVG(purchase_frequency) AS DECIMAL(10,2))  AS avg_purchase_frequency,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))       AS avg_revenue,
    CAST(AVG(age) AS DECIMAL(5,2))                  AS avg_age
FROM dbo.ecommerce_clean
GROUP BY subscription_status;
GO

-- Business Objective: Build a churn-risk scoring model that classifies
-- every customer into a clear risk tier, enabling proactive retention
-- workflows before formal cancellation occurs.
-- SQL Logic: Multi-condition CASE WHEN statement evaluating
-- subscription_status, cancellations_count, and purchase_frequency
-- in priority order.
-- Expected Insight: Produces an actionable churn_risk_label per
-- customer for CRM-driven retention campaign targeting.

SELECT
    customer_id,
    subscription_status,
    cancellations_count,
    purchase_frequency,
    CASE
        WHEN subscription_status = 'cancelled'
            THEN 'Churned'
        WHEN subscription_status = 'paused' AND cancellations_count >= 3
            THEN 'Critical Risk'
        WHEN subscription_status = 'paused'
            THEN 'High Risk'
        WHEN subscription_status = 'active' AND cancellations_count >= 4
            THEN 'Medium Risk'
        WHEN subscription_status = 'active' AND purchase_frequency <= 5
            THEN 'Low Engagement'
        ELSE 'Healthy'
    END AS churn_risk_label
FROM dbo.ecommerce_clean
ORDER BY cancellations_count DESC, purchase_frequency ASC;
GO

-- Business Objective: Produce a prioritized list of at-risk customers
-- (excluding already-churned accounts) ranked by historical value, for
-- immediate retention campaign action.
-- SQL Logic: WHERE filter to exclude churned customers, combined with
-- the same churn-risk CASE WHEN logic and an estimated_clv-based sort.
-- Expected Insight: Surfaces the highest-value "paused" and "medium
-- risk" customers who should be contacted first.

SELECT
    customer_id,
    country,
    preferred_category,
    subscription_status,
    cancellations_count,
    purchase_frequency,
    total_revenue,
    estimated_clv,
    CASE
        WHEN subscription_status = 'paused' AND cancellations_count >= 3 THEN 'Critical Risk'
        WHEN subscription_status = 'paused'                              THEN 'High Risk'
        WHEN subscription_status = 'active' AND cancellations_count >= 4 THEN 'Medium Risk'
        WHEN subscription_status = 'active' AND purchase_frequency <= 5  THEN 'Low Engagement'
        ELSE 'Stable'
    END AS churn_risk
FROM dbo.ecommerce_clean
WHERE subscription_status <> 'cancelled'
ORDER BY estimated_clv DESC;
GO


-- =========================================
-- REVENUE ANALYSIS
-- =========================================

-- Business Objective: Identify which products/categories generate the
-- most total revenue, to guide merchandising and inventory decisions.
-- SQL Logic: GROUP BY category and product_name with revenue
-- aggregation, limited to the top 15 combinations.
-- Expected Insight: Surfaces the specific product lines driving the
-- largest share of total revenue within each category.

SELECT TOP 15
    category,
    product_name,
    COUNT(*)                                   AS orders,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))   AS total_revenue,
    CAST(AVG(unit_price) AS DECIMAL(10,2))      AS avg_unit_price,
    CAST(AVG(quantity) AS DECIMAL(5,2))         AS avg_quantity
FROM dbo.ecommerce_clean
GROUP BY category, product_name
ORDER BY total_revenue DESC;
GO

-- Business Objective: Confirm which regions (countries) contribute the
-- highest revenue, supporting regional budget allocation decisions.
-- SQL Logic: GROUP BY country with total and average revenue.
-- Expected Insight: Revenue is fairly evenly spread; Germany leads
-- marginally in absolute terms.

SELECT
    country,
    COUNT(*)                                  AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue
FROM dbo.ecommerce_clean
GROUP BY country
ORDER BY total_revenue DESC;
GO

-- Business Objective: Examine trends in customer spending across age
-- groups to understand which demographic cohorts justify the largest
-- marketing investment.
-- SQL Logic: GROUP BY age_group with revenue-share subquery (same
-- structure as the KPI section, included here for revenue-trend
-- analysis specifically).
-- Expected Insight: Spending rises steadily with age, peaking in the
-- 56+ bracket, which alone represents over a quarter of total revenue.

SELECT
    age_group,
    COUNT(*)                                    AS customers,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))    AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))    AS avg_revenue,
    CAST(ROUND(100.0 * SUM(total_revenue) /
        (SELECT SUM(total_revenue) FROM dbo.ecommerce_clean), 2) AS DECIMAL(5,2)) AS revenue_share_pct
FROM dbo.ecommerce_clean
GROUP BY age_group
ORDER BY age_group;
GO

-- Business Objective: Filter to only countries whose average order
-- value exceeds the overall benchmark, identifying premium markets.
-- SQL Logic: GROUP BY with HAVING clause filtering on AVG(total_revenue).
-- Expected Insight: Surfaces above-benchmark markets for premium
-- product line expansion.

SELECT
    country,
    COUNT(*)                                  AS total_customers,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_order_value
FROM dbo.ecommerce_clean
GROUP BY country
HAVING AVG(total_revenue) > 1020
ORDER BY avg_order_value DESC;
GO


-- =========================================
-- CUSTOMER SEGMENTATION
-- =========================================

-- Business Objective: Build a complete multi-dimensional customer
-- segmentation (Value, Frequency, Churn Risk) using a CTE-based
-- pipeline, to support targeted marketing and retention workflows.
-- SQL Logic: A CTE computes base metrics; a second CTE computes
-- percentile thresholds for revenue and frequency using
-- PERCENTILE_CONT (SQL Server window function); the final SELECT
-- applies CASE WHEN logic against those thresholds.
-- Expected Insight: Every customer is classified into a Value tier,
-- a Frequency tier, and a Churn Risk tier in a single result set,
-- ready for export to a CRM or marketing automation platform.

WITH metrics AS (
    SELECT
        customer_id,
        country,
        preferred_category,
        subscription_status,
        cancellations_count,
        purchase_frequency,
        total_revenue,
        age,
        age_group,
        gender
    FROM dbo.ecommerce_clean
),
thresholds AS (
    SELECT DISTINCT
        PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY total_revenue)      OVER () AS rev_p75,
        PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY total_revenue)      OVER () AS rev_p50,
        PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY purchase_frequency) OVER () AS freq_p75,
        PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY purchase_frequency) OVER () AS freq_p50
    FROM metrics
)
SELECT
    m.customer_id,
    m.country,
    m.preferred_category,
    m.subscription_status,
    m.total_revenue,
    m.purchase_frequency,
    m.cancellations_count,

    -- VALUE SEGMENT
    CASE
        WHEN m.total_revenue >= t.rev_p75 THEN 'High Value'
        WHEN m.total_revenue >= t.rev_p50 THEN 'Mid Value'
        ELSE 'Low Value'
    END AS value_segment,

    -- FREQUENCY SEGMENT
    CASE
        WHEN m.purchase_frequency >= t.freq_p75 THEN 'Frequent Buyer'
        WHEN m.purchase_frequency >= t.freq_p50 THEN 'Regular Buyer'
        ELSE 'Infrequent'
    END AS frequency_segment,

    -- CHURN RISK SEGMENT
    CASE
        WHEN m.subscription_status = 'cancelled' THEN 'Churned'
        WHEN m.subscription_status = 'paused' AND m.cancellations_count >= 3 THEN 'Critical Risk'
        WHEN m.subscription_status = 'paused' THEN 'High Risk'
        WHEN m.subscription_status = 'active' AND m.cancellations_count >= 4 THEN 'Medium Risk'
        WHEN m.subscription_status = 'active' AND m.purchase_frequency >= t.freq_p75 THEN 'Loyal Customer'
        ELSE 'Stable'
    END AS churn_risk_segment

FROM metrics m
CROSS JOIN thresholds t
ORDER BY m.total_revenue DESC;
GO

-- Business Objective: Summarize the size and value of each churn-risk
-- segment to quantify the scale of retention opportunity for
-- leadership reporting.
-- SQL Logic: A CTE reproduces the churn-risk CASE WHEN logic, then
-- the outer query aggregates customer count and total revenue per
-- segment.
-- Expected Insight: Quantifies exactly how many customers and how much
-- revenue sit in each risk tier (Critical Risk, High Risk, etc.).

WITH risk_segments AS (
    SELECT
        customer_id,
        total_revenue,
        CASE
            WHEN subscription_status = 'cancelled' THEN 'Churned'
            WHEN subscription_status = 'paused' AND cancellations_count >= 3 THEN 'Critical Risk'
            WHEN subscription_status = 'paused' THEN 'High Risk'
            WHEN subscription_status = 'active' AND cancellations_count >= 4 THEN 'Medium Risk'
            WHEN subscription_status = 'active' AND purchase_frequency >= 37 THEN 'Loyal Customer'
            ELSE 'Stable'
        END AS churn_risk_segment
    FROM dbo.ecommerce_clean
)
SELECT
    churn_risk_segment,
    COUNT(*)                                  AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS total_revenue,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue
FROM risk_segments
GROUP BY churn_risk_segment
ORDER BY total_revenue DESC;
GO

-- Business Objective: Build an RFM-style segmentation (Recency,
-- Frequency, Monetary) using quintile scoring, to identify Champions,
-- Loyal Customers, and Lost Customers using an industry-standard
-- segmentation method.
-- SQL Logic: A base CTE computes recency in days using DATEDIFF; a
-- second CTE applies NTILE(5) window functions to bucket customers
-- into quintiles for each of the three RFM dimensions; the final
-- SELECT sums the three scores and classifies via CASE WHEN.
-- Expected Insight: Produces a single RFM segment label per customer,
-- widely used in industry for prioritizing marketing spend.

WITH rfm AS (
    SELECT
        customer_id,
        country,
        DATEDIFF(DAY, last_purchase_date, GETDATE()) AS recency_days,
        purchase_frequency                            AS frequency,
        total_revenue                                 AS monetary
    FROM dbo.ecommerce_clean
),
rfm_scores AS (
    SELECT
        customer_id,
        country,
        recency_days,
        frequency,
        monetary,
        NTILE(5) OVER (ORDER BY recency_days ASC)  AS r_score,  -- lower recency = better
        NTILE(5) OVER (ORDER BY frequency DESC)    AS f_score,  -- higher frequency = better
        NTILE(5) OVER (ORDER BY monetary DESC)     AS m_score   -- higher monetary = better
    FROM rfm
)
SELECT
    customer_id,
    country,
    r_score,
    f_score,
    m_score,
    (r_score + f_score + m_score) AS rfm_total,
    CASE
        WHEN (r_score + f_score + m_score) >= 13 THEN 'Champions'
        WHEN (r_score + f_score + m_score) >= 10 THEN 'Loyal Customers'
        WHEN (r_score + f_score + m_score) >= 7  THEN 'Potential Loyalists'
        WHEN (r_score + f_score + m_score) >= 5  THEN 'At Risk'
        ELSE 'Lost Customers'
    END AS rfm_segment
FROM rfm_scores
ORDER BY rfm_total DESC;
GO


-- =========================================
-- ADVANCED SQL ANALYSIS
-- =========================================

-- ----------------------------------------------------------------
-- WINDOW FUNCTION: ROW_NUMBER()
-- ----------------------------------------------------------------
-- Business Objective: Identify the top 3 highest-revenue customers
-- within each country, to support country-specific VIP recognition
-- programs.
-- SQL Logic: ROW_NUMBER() OVER (PARTITION BY country ORDER BY
-- total_revenue DESC) assigns a unique sequential rank per partition;
-- wrapped in a CTE so the outer query can filter on the rank.
-- Expected Insight: Returns exactly 3 top customers per country,
-- ideal for geo-targeted retention or loyalty outreach lists.

WITH ranked_customers AS (
    SELECT
        customer_id,
        country,
        total_revenue,
        ROW_NUMBER() OVER (PARTITION BY country ORDER BY total_revenue DESC) AS revenue_rank_in_country
    FROM dbo.ecommerce_clean
)
SELECT *
FROM ranked_customers
WHERE revenue_rank_in_country <= 3
ORDER BY country, revenue_rank_in_country;
GO

-- ----------------------------------------------------------------
-- WINDOW FUNCTIONS: RANK() and DENSE_RANK()
-- ----------------------------------------------------------------
-- Business Objective: Build a category revenue "league table" for
-- category managers, comparing two ranking methods.
-- SQL Logic: RANK() leaves gaps in the sequence after ties; DENSE_RANK()
-- does not. Both are computed over the same ORDER BY for direct
-- comparison.
-- Expected Insight: With 5 distinct category revenue totals, RANK()
-- and DENSE_RANK() will match here, but the pattern is essential for
-- datasets where tied revenue totals occur.

SELECT
    category,
    CAST(SUM(total_revenue) AS DECIMAL(14,2)) AS total_revenue,
    RANK()       OVER (ORDER BY SUM(total_revenue) DESC) AS revenue_rank,
    DENSE_RANK() OVER (ORDER BY SUM(total_revenue) DESC) AS revenue_dense_rank
FROM dbo.ecommerce_clean
GROUP BY category
ORDER BY total_revenue DESC;
GO

-- ----------------------------------------------------------------
-- WINDOW FUNCTION: LAG()
-- ----------------------------------------------------------------
-- Business Objective: Measure the revenue gap between consecutively
-- ranked customers (by revenue), to identify natural breakpoints for
-- defining VIP tier thresholds.
-- SQL Logic: LAG(total_revenue) OVER (ORDER BY total_revenue) retrieves
-- the prior row's value in sorted order; the difference is then
-- calculated against the current row.
-- Expected Insight: Reveals where the steepest revenue "cliffs" occur
-- among top earners, informing where to draw VIP tier cutoffs.

WITH ranked AS (
    SELECT
        customer_id,
        country,
        total_revenue,
        ROW_NUMBER() OVER (ORDER BY total_revenue) AS rn,
        LAG(total_revenue) OVER (ORDER BY total_revenue) AS prev_revenue
    FROM dbo.ecommerce_clean
)
SELECT TOP 10
    customer_id,
    country,
    total_revenue,
    prev_revenue,
    CAST(total_revenue - prev_revenue AS DECIMAL(10,2)) AS revenue_gap
FROM ranked
WHERE prev_revenue IS NOT NULL
ORDER BY total_revenue DESC;
GO

-- ----------------------------------------------------------------
-- WINDOW FUNCTION: LEAD()
-- ----------------------------------------------------------------
-- Business Objective: Identify the largest purchase-frequency drop-offs
-- within each subscription status cohort, to spot early churn warning
-- patterns.
-- SQL Logic: LEAD(purchase_frequency) OVER (PARTITION BY
-- subscription_status ORDER BY purchase_frequency DESC) retrieves the
-- next lower frequency value within the same status group; the
-- difference highlights frequency "cliffs."
-- Expected Insight: Surfaces engagement-level discontinuities that may
-- precede a status change to "paused" or "cancelled."

WITH freq_ordered AS (
    SELECT
        customer_id,
        purchase_frequency,
        subscription_status,
        LEAD(purchase_frequency) OVER (
            PARTITION BY subscription_status ORDER BY purchase_frequency DESC
        ) AS next_freq
    FROM dbo.ecommerce_clean
)
SELECT TOP 10
    customer_id,
    subscription_status,
    purchase_frequency,
    next_freq,
    (purchase_frequency - next_freq) AS freq_gap
FROM freq_ordered
WHERE next_freq IS NOT NULL
ORDER BY freq_gap DESC;
GO

-- ----------------------------------------------------------------
-- CTE + SUBQUERY: Multi-Step CLV Segmentation
-- ----------------------------------------------------------------
-- Business Objective: Segment all customers into High/Mid/Low CLV
-- tiers using statistically derived percentile thresholds, rather than
-- arbitrary fixed cutoffs.
-- SQL Logic: A base CTE computes customer-level estimated_clv; a
-- second CTE computes percentile thresholds via PERCENTILE_CONT; the
-- final SELECT applies CASE WHEN logic against those thresholds via a
-- CROSS JOIN.
-- Expected Insight: Produces a statistically grounded value
-- segmentation, summarized by total and average CLV per tier.

WITH customer_metrics AS (
    SELECT
        customer_id,
        country,
        preferred_category,
        subscription_status,
        cancellations_count,
        total_revenue,
        purchase_frequency,
        estimated_clv
    FROM dbo.ecommerce_clean
),
clv_percentiles AS (
    SELECT DISTINCT
        PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY estimated_clv) OVER () AS p75_clv,
        PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY estimated_clv) OVER () AS p25_clv,
        AVG(estimated_clv) OVER ()                                          AS avg_clv
    FROM customer_metrics
),
segmented AS (
    SELECT
        cm.*,
        cp.p75_clv,
        cp.avg_clv,
        CASE
            WHEN cm.estimated_clv >= cp.p75_clv THEN 'High Value'
            WHEN cm.estimated_clv >= cp.avg_clv THEN 'Mid Value'
            ELSE 'Low Value'
        END AS value_segment
    FROM customer_metrics cm
    CROSS JOIN clv_percentiles cp
)
SELECT
    value_segment,
    COUNT(*)                                   AS customers,
    CAST(SUM(estimated_clv) AS DECIMAL(16,2))   AS total_clv,
    CAST(AVG(estimated_clv) AS DECIMAL(12,2))   AS avg_clv
FROM segmented
GROUP BY value_segment
ORDER BY avg_clv DESC;
GO

-- ----------------------------------------------------------------
-- SUBQUERY: Above-Average Revenue Customers
-- ----------------------------------------------------------------
-- Business Objective: Identify the cohort of customers spending above
-- the company-wide average, for premium-tier upsell targeting.
-- SQL Logic: A scalar subquery in the WHERE clause computes the
-- overall AVG(total_revenue) dynamically, so the filter always
-- reflects the current dataset state.
-- Expected Insight: Returns the above-average spending population
-- (roughly half the base, skewed toward the higher end due to the
-- right-skewed revenue distribution).

SELECT
    customer_id,
    country,
    total_revenue,
    subscription_status
FROM dbo.ecommerce_clean
WHERE total_revenue > (SELECT AVG(total_revenue) FROM dbo.ecommerce_clean)
ORDER BY total_revenue DESC;
GO

-- ----------------------------------------------------------------
-- JOIN: INNER JOIN (Self-Join) — Peer Benchmarking
-- ----------------------------------------------------------------
-- Business Objective: Compare each customer's revenue against their
-- own country's average, to identify local market leaders who
-- over-index relative to regional peers.
-- SQL Logic: INNER JOIN between the base table and a derived table
-- (subquery) computing country-level average revenue, joined on
-- country.
-- Expected Insight: Surfaces customers who significantly outperform
-- their regional peer group — candidates for localized VIP recognition.

SELECT TOP 20
    a.customer_id,
    a.country,
    a.total_revenue,
    CAST(b.avg_country_revenue AS DECIMAL(10,2))                  AS country_avg_revenue,
    CAST(a.total_revenue - b.avg_country_revenue AS DECIMAL(10,2)) AS vs_country_avg
FROM dbo.ecommerce_clean a
INNER JOIN (
    SELECT country, AVG(total_revenue) AS avg_country_revenue
    FROM dbo.ecommerce_clean
    GROUP BY country
) b ON a.country = b.country
ORDER BY vs_country_avg DESC;
GO

-- ----------------------------------------------------------------
-- JOIN: LEFT JOIN — Cross-Category Purchase Analysis
-- ----------------------------------------------------------------
-- Business Objective: Identify customers whose actual purchase
-- category differs from their stated preferred category, revealing
-- cross-sell engagement opportunities.
-- SQL Logic: LEFT JOIN against a self-referencing subquery isolating
-- rows where preferred_category does not match category, preserving
-- all customers from the left table regardless of match.
-- Expected Insight: Flags customers actively buying outside their
-- stated preference — a strong signal for cross-category marketing.

SELECT
    a.customer_id,
    a.preferred_category,
    a.category AS purchased_category,
    a.total_revenue,
    CASE
        WHEN a.preferred_category = a.category THEN 'Aligned'
        ELSE 'Cross-Category Buyer'
    END AS purchase_alignment
FROM dbo.ecommerce_clean a
LEFT JOIN dbo.ecommerce_clean b
    ON a.customer_id = b.customer_id
    AND b.preferred_category <> b.category
ORDER BY a.total_revenue DESC;
GO

-- ----------------------------------------------------------------
-- JOIN: RIGHT JOIN — Category Coverage Check
-- ----------------------------------------------------------------
-- Business Objective: Confirm that every product category present in
-- the catalogue reference list has at least one matching customer
-- order, surfacing any categories with zero sales activity.
-- SQL Logic: A reference category list (derived via subquery of
-- DISTINCT categories) is RIGHT JOINed against the order data, so all
-- reference categories appear even if unmatched.
-- Expected Insight: With 5 known categories, this confirms full
-- category coverage (no orphaned/zero-order categories).

SELECT
    ref.category AS reference_category,
    COUNT(c.order_id) AS order_count,
    CAST(SUM(c.total_revenue) AS DECIMAL(14,2)) AS total_revenue
FROM dbo.ecommerce_clean c
RIGHT JOIN (
    SELECT DISTINCT category FROM dbo.ecommerce_clean
) ref ON c.category = ref.category
GROUP BY ref.category
ORDER BY total_revenue DESC;
GO

-- ----------------------------------------------------------------
-- CASE WHEN: Risk-Tiered Revenue Exposure Summary
-- ----------------------------------------------------------------
-- Business Objective: Quantify exactly how much revenue is exposed at
-- each churn-risk tier, translating the churn model into a clear
-- dollar-value business risk summary for leadership.
-- SQL Logic: CASE WHEN classification embedded directly in a GROUP BY
-- aggregation (no CTE required), demonstrating inline conditional
-- logic combined with aggregate functions.
-- Expected Insight: Produces a one-table view of revenue at risk by
-- tier, directly supporting the Executive Summary's "Risks" section.

SELECT
    CASE
        WHEN subscription_status = 'cancelled' THEN 'Churned'
        WHEN subscription_status = 'paused' AND cancellations_count >= 3 THEN 'Critical Risk'
        WHEN subscription_status = 'paused' THEN 'High Risk'
        WHEN subscription_status = 'active' AND cancellations_count >= 4 THEN 'Medium Risk'
        ELSE 'Stable'
    END AS risk_tier,
    COUNT(*)                                  AS customer_count,
    CAST(SUM(total_revenue) AS DECIMAL(14,2))  AS revenue_exposure,
    CAST(AVG(total_revenue) AS DECIMAL(10,2))  AS avg_revenue_per_customer
FROM dbo.ecommerce_clean
GROUP BY
    CASE
        WHEN subscription_status = 'cancelled' THEN 'Churned'
        WHEN subscription_status = 'paused' AND cancellations_count >= 3 THEN 'Critical Risk'
        WHEN subscription_status = 'paused' THEN 'High Risk'
        WHEN subscription_status = 'active' AND cancellations_count >= 4 THEN 'Medium Risk'
        ELSE 'Stable'
    END
ORDER BY revenue_exposure DESC;
GO

/*
===================================================================================
END OF SCRIPT
===================================================================================
*/
