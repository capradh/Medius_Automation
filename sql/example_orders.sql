-- Demo query that works on ANY Oracle database (generates 100 rows, no tables needed).
-- Run: python scripts\run_job.py db_to_excel --db corsair --sql-file sql\example_orders.sql --group-by region_id
SELECT LEVEL                          AS order_id,
       MOD(LEVEL, 5) + 1              AS region_id,
       MOD(LEVEL * 37, 500) + 10      AS amount,
       TRUNC(SYSDATE) - MOD(LEVEL, 30) AS order_date
FROM   dual
CONNECT BY LEVEL <= 100
