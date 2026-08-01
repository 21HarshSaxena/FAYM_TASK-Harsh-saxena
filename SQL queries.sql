/* ============================================================================
   D2C Wallet Transactions — SQL Solutions
   Table: transactions (transaction_time DATE, user_id INT, transaction_amt INT,
                         narration TEXT, transaction_type TEXT, txn_id TEXT)
   narration        = payment rail (IMPS / IFT / NEFT / RTGS / UPI)
   transaction_type = direction (CREDIT = money in, DEBIT = money out)
   Dialect: SQLite (window functions); portable to Postgres/MySQL 8+ with no changes.
   ============================================================================ */


/* --------------------------------------------------------------------------
   Q1. 7th highest debit amount transacted through IMPS
   Answer on the sample data: 9,525 | 2020-01-16 | User 8 | Txn 64B08A24
   -------------------------------------------------------------------------- */
SELECT transaction_time, user_id, transaction_amt, narration, transaction_type, txn_id
FROM transactions
WHERE narration = 'IMPS' AND transaction_type = 'DEBIT'
ORDER BY transaction_amt DESC
LIMIT 1 OFFSET 6;                         -- OFFSET 6 -> the 7th row when 0-indexed

-- Context: full top-10 IMPS-debit ranking
SELECT
    ROW_NUMBER() OVER (ORDER BY transaction_amt DESC) AS rnk,
    transaction_time, user_id, transaction_amt, txn_id
FROM transactions
WHERE narration = 'IMPS' AND transaction_type = 'DEBIT'
ORDER BY transaction_amt DESC
LIMIT 10;


/* --------------------------------------------------------------------------
   Q2. Number of transactions, category-wise (UPI / IMPS / RTGS / NEFT [+ IFT])
   Answer: IMPS 254, IFT 83, UPI 79, NEFT 30, RTGS 21  (467 total)
   -------------------------------------------------------------------------- */
SELECT
    narration AS category,
    COUNT(*)  AS txn_count,
    ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM transactions), 2) AS pct_of_total
FROM transactions
GROUP BY narration
ORDER BY txn_count DESC;


/* --------------------------------------------------------------------------
   Q3. Statistical summary of "load amount" (CREDIT transaction_amt)
   Bell curve / box plot are generated separately (see sheet_workbook)
   -------------------------------------------------------------------------- */
SELECT
    COUNT(*)                                        AS n,
    AVG(transaction_amt)                             AS mean_amt,
    MIN(transaction_amt)                             AS min_amt,
    MAX(transaction_amt)                             AS max_amt,
    SQRT(AVG(transaction_amt*transaction_amt) - AVG(transaction_amt)*AVG(transaction_amt)) AS population_stddev
FROM transactions
WHERE transaction_type = 'CREDIT';

/* --------------------------------------------------------------------------
   Q4. Monthly cohort of active users (users doing DEBIT transactions)
   Cohort = calendar month of a user's FIRST DEBIT txn.
   -------------------------------------------------------------------------- */
WITH first_debit AS (
    SELECT user_id, MIN(strftime('%Y-%m', transaction_time)) AS cohort_month
    FROM transactions
    WHERE transaction_type = 'DEBIT'
    GROUP BY user_id
),
activity AS (
    SELECT user_id, strftime('%Y-%m', transaction_time) AS active_month
    FROM transactions
    WHERE transaction_type = 'DEBIT'
    GROUP BY user_id, strftime('%Y-%m', transaction_time)
)
SELECT
    f.cohort_month,
    a.active_month,
    COUNT(DISTINCT a.user_id) AS active_users
FROM first_debit f
JOIN activity a ON a.user_id = f.user_id
GROUP BY f.cohort_month, a.active_month
ORDER BY f.cohort_month, a.active_month;

/* --------------------------------------------------------------------------
   Q5. Top 10th percentile users by net amount (DEBIT − CREDIT)
   Answer on the sample data: User 3 only (net = -24,197). With just 10 unique
   users, the 10th percentile threshold resolves to a single user the query
   itself scales correctly to any user count.
   -------------------------------------------------------------------------- */
WITH user_net AS (
    SELECT
        user_id,
        SUM(CASE WHEN transaction_type = 'DEBIT'  THEN transaction_amt ELSE 0 END) AS total_debit,
        SUM(CASE WHEN transaction_type = 'CREDIT' THEN transaction_amt ELSE 0 END) AS total_credit,
        SUM(CASE WHEN transaction_type = 'DEBIT'  THEN transaction_amt ELSE 0 END)
      - SUM(CASE WHEN transaction_type = 'CREDIT' THEN transaction_amt ELSE 0 END) AS net_amount
    FROM transactions
    GROUP BY user_id
),
ranked AS (
    SELECT user_id, total_debit, total_credit, net_amount,
           PERCENT_RANK() OVER (ORDER BY net_amount DESC) AS pct_rank
    FROM user_net
)
SELECT user_id, total_debit, total_credit, net_amount,
       ROUND(pct_rank * 100, 1) AS percentile_from_top
FROM ranked
WHERE pct_rank <= 0.10
ORDER BY net_amount DESC;
