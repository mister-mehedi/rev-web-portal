const express = require('express');
const router = express.Router();
const oracledb = require('oracledb');

// Dashboard
router.get('/', (req, res) => {
    res.render('index');
});

// 1️⃣ VDD F2 Day Chunk Summary
router.post('/day-chunk-summary', async (req, res) => {
    const { baseDate, chunkDays, numChunks } = req.body;
    executeQuery(`
WITH PARAMS AS (
  SELECT TO_DATE(:baseDate, 'DD-MON-RR') AS BASE_DATE,
         :chunkDays AS CHUNK_DAYS,
         :numChunks AS NUM_CHUNKS FROM DUAL
),
RANGES AS (
  SELECT LEVEL AS CHUNK_NO,
         (p.BASE_DATE - (LEVEL - 1) * p.CHUNK_DAYS) AS CHUNK_END,
         (p.BASE_DATE - LEVEL * p.CHUNK_DAYS + 1) AS CHUNK_START
  FROM PARAMS p
  CONNECT BY LEVEL <= p.NUM_CHUNKS
),
CHUNK_SUMMARY AS (
  SELECT r.CHUNK_NO,
         TO_CHAR(r.CHUNK_START,'DD-MON-RR') AS START_DATE,
         TO_CHAR(r.CHUNK_END,'DD-MON-RR') AS END_DATE,
         SUM(v.TOTAL_REVENUE_MOC_MTC) AS TOTAL_REVENUE_MOC_MTC,
         SUM(v.DATA_REV_TOT) AS DATA_REV_TOT,
         SUM(v.UNIQUE_VSUBS) AS TOTAL_UNIQUE_VSUBS,
         SUM(v.UNIQUE_DSUBS) AS TOTAL_UNIQUE_DSUBS,
         COUNT(DISTINCT v.DATE_VALUE) AS PRESENT_DAYS
  FROM RANGES r
  LEFT JOIN voice_data_details_final_2 v
  ON v.DATE_VALUE BETWEEN r.CHUNK_START AND r.CHUNK_END
  GROUP BY r.CHUNK_NO, r.CHUNK_START, r.CHUNK_END
),
FINAL_RESULT AS (
  SELECT CHUNK_NO, START_DATE, END_DATE, TOTAL_REVENUE_MOC_MTC, DATA_REV_TOT,
         CASE WHEN PRESENT_DAYS > 0 THEN ROUND(TOTAL_UNIQUE_VSUBS / PRESENT_DAYS, 2) END AS AVG_UNIQUE_VSUBS,
         CASE WHEN PRESENT_DAYS > 0 THEN ROUND(TOTAL_UNIQUE_DSUBS / PRESENT_DAYS, 2) END AS AVG_UNIQUE_DSUBS,
         PRESENT_DAYS
  FROM CHUNK_SUMMARY
)
SELECT CHUNK_NO, START_DATE, END_DATE, TOTAL_REVENUE_MOC_MTC, DATA_REV_TOT, AVG_UNIQUE_VSUBS, AVG_UNIQUE_DSUBS, PRESENT_DAYS
FROM FINAL_RESULT
ORDER BY CHUNK_NO DESC
    `, { baseDate, chunkDays: Number(chunkDays), numChunks: Number(numChunks) }, res);
});

// 2️⃣ VDD F2 Month Chunk Summary
router.post('/month-chunk-summary', async (req, res) => {
    const { baseMonth, chunkMonths } = req.body;
    executeQuery(`
WITH PARAMS AS (
  SELECT :baseMonth AS BASE_MONTH,
         :chunkMonths AS CHUNK_MONTHS FROM DUAL
),
MONTH_RANGE AS (
  SELECT TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(TO_CHAR(BASE_MONTH)||'01','YYYYMMDD'),-(LEVEL-1)),'YYYYMM')) AS MONTH_KEY,
         LEVEL AS CHUNK_NO
  FROM PARAMS
  CONNECT BY LEVEL <= (SELECT CHUNK_MONTHS FROM PARAMS)
),
AGG_DATA AS (
  SELECT V.MONTH_KEY,
         SUM(V.DATA_REV_TOT) AS TOTAL_DATA_REV_TOT,
         SUM(V.TOTAL_REVENUE_MOC_MTC) AS TOTAL_REVENUE_MOC_MTC,
         ROUND(SUM(V.UNIQUE_VSUBS)/COUNT(DISTINCT V.DATE_VALUE),2) AS AVG_DAILY_VSUBS,
         ROUND(SUM(V.UNIQUE_DSUBS)/COUNT(DISTINCT V.DATE_VALUE),2) AS AVG_DAILY_DSUBS
  FROM voice_data_details_final_2 V
  WHERE V.MONTH_KEY IN (SELECT MONTH_KEY FROM MONTH_RANGE)
  GROUP BY V.MONTH_KEY
)
SELECT M.CHUNK_NO, M.MONTH_KEY, A.TOTAL_DATA_REV_TOT, A.TOTAL_REVENUE_MOC_MTC, A.AVG_DAILY_VSUBS, A.AVG_DAILY_DSUBS
FROM MONTH_RANGE M
LEFT JOIN AGG_DATA A ON M.MONTH_KEY=A.MONTH_KEY
ORDER BY M.CHUNK_NO
    `, { baseMonth: Number(baseMonth), chunkMonths: Number(chunkMonths) }, res);
});

// 3️⃣ Zone Sales FO Day Chunk Summary
router.post('/fo-day-chunk-summary', async (req, res) => {
    const { baseDate, chunkDays, numChunks } = req.body;
    executeQuery(`
WITH PARAMS AS (
  SELECT TO_DATE(:baseDate,'DD-MON-RR') AS BASE_DATE,
         :chunkDays AS CHUNK_DAYS,
         :numChunks AS NUM_CHUNKS
  FROM DUAL
),
RANGES AS (
  SELECT LEVEL AS CHUNK_NO,
         (p.BASE_DATE-(LEVEL-1)*p.CHUNK_DAYS) AS CHUNK_END,
         (p.BASE_DATE-LEVEL*p.CHUNK_DAYS+1) AS CHUNK_START
  FROM PARAMS p
  CONNECT BY LEVEL<=p.NUM_CHUNKS
),
CHUNK_OFFICER_SUMMARY AS (
  SELECT r.CHUNK_NO, z.FIELD_OFFICER_NAME, z.INCHARGE_NAME,
         TO_CHAR(r.CHUNK_START,'DD-MON-RR') AS START_DATE,
         TO_CHAR(r.CHUNK_END,'DD-MON-RR') AS END_DATE,
         SUM(z.TOTAL_REVENUE_MOC_MTC) AS TOTAL_REVENUE_MOC_MTC,
         SUM(z.DATA_REV_TOT) AS DATA_REV_TOT,
         SUM(z.UNIQUE_VSUBS) AS TOTAL_UNIQUE_VSUBS,
         SUM(z.UNIQUE_DSUBS) AS TOTAL_UNIQUE_DSUBS,
         COUNT(DISTINCT z.DATE_VALUE) AS PRESENT_DAYS
  FROM RANGES r
  LEFT JOIN ZONE_SALES_T_FIELD_OFFICER_DAY_REV z
  ON z.DATE_VALUE BETWEEN r.CHUNK_START AND r.CHUNK_END
  GROUP BY r.CHUNK_NO,r.CHUNK_START,r.CHUNK_END,z.FIELD_OFFICER_NAME,z.INCHARGE_NAME
),
FINAL_RESULT AS (
  SELECT CHUNK_NO, START_DATE, END_DATE, FIELD_OFFICER_NAME, INCHARGE_NAME,
         TOTAL_REVENUE_MOC_MTC, DATA_REV_TOT,
         CASE WHEN PRESENT_DAYS>0 THEN ROUND(TOTAL_UNIQUE_VSUBS/PRESENT_DAYS,2) END AS AVG_UNIQUE_VSUBS,
         CASE WHEN PRESENT_DAYS>0 THEN ROUND(TOTAL_UNIQUE_DSUBS/PRESENT_DAYS,2) END AS AVG_UNIQUE_DSUBS,
         PRESENT_DAYS
  FROM CHUNK_OFFICER_SUMMARY
)
SELECT CHUNK_NO, START_DATE, END_DATE, FIELD_OFFICER_NAME, INCHARGE_NAME,
       TOTAL_REVENUE_MOC_MTC, DATA_REV_TOT, AVG_UNIQUE_VSUBS, AVG_UNIQUE_DSUBS, PRESENT_DAYS
FROM FINAL_RESULT
ORDER BY CHUNK_NO DESC, FIELD_OFFICER_NAME, INCHARGE_NAME
    `, { baseDate, chunkDays: Number(chunkDays), numChunks: Number(numChunks) }, res);
});

// 4️⃣ Zone Sales FO Month Chunk Summary
router.post('/fo-month-chunk-summary', async (req, res) => {
    const { baseMonth, chunkMonths } = req.body;
    executeQuery(`
WITH PARAMS AS (
  SELECT :baseMonth AS BASE_MONTH,
         :chunkMonths AS CHUNK_MONTHS
  FROM DUAL
),
MONTH_RANGE AS (
  SELECT P.BASE_MONTH,
         ADD_MONTHS(TO_DATE(TO_CHAR(P.BASE_MONTH)||'01','YYYYMMDD'),-(LEVEL-1)) AS MONTH_START
  FROM PARAMS P
  CONNECT BY LEVEL<=P.CHUNK_MONTHS
),
DATA AS (
  SELECT Z.MONTH_KEY, Z.FIELD_OFFICER_NAME, Z.INCHARGE_NAME,
         SUM(Z.DATA_REV_TOT) AS TOTAL_DATA_REV_TOT,
         SUM(Z.TOTAL_REVENUE_MOC_MTC) AS TOTAL_REVENUE_MOC_MTC,
         ROUND(SUM(Z.UNIQUE_VSUBS)/COUNT(DISTINCT Z.DATE_VALUE),2) AS AVG_DAILY_VSUBS,
         ROUND(SUM(Z.UNIQUE_DSUBS)/COUNT(DISTINCT Z.DATE_VALUE),2) AS AVG_DAILY_DSUBS
  FROM ZONE_SALES_T_FIELD_OFFICER_DAY_REV Z
  WHERE Z.MONTH_KEY IN (SELECT TO_NUMBER(TO_CHAR(MONTH_START,'YYYYMM')) FROM MONTH_RANGE)
  GROUP BY Z.MONTH_KEY,Z.FIELD_OFFICER_NAME,Z.INCHARGE_NAME
)
SELECT D.MONTH_KEY, D.FIELD_OFFICER_NAME, D.INCHARGE_NAME, D.TOTAL_DATA_REV_TOT, D.TOTAL_REVENUE_MOC_MTC, D.AVG_DAILY_VSUBS, D.AVG_DAILY_DSUBS
FROM DATA D
ORDER BY D.MONTH_KEY,D.FIELD_OFFICER_NAME
    `, { baseMonth: Number(baseMonth), chunkMonths: Number(chunkMonths) }, res);
});

// Helper function for query execution
async function executeQuery(query, binds, res) {
    let connection;
    try {
        connection = await oracledb.getConnection({
            user: process.env.DB_USER,
            password: process.env.DB_PASSWORD,
            connectString: process.env.DB_CONNECT_STRING
        });
        const result = await connection.execute(query, binds, { outFormat: oracledb.OUT_FORMAT_OBJECT });
        res.render('results', { data: result.rows });
    } catch (err) {
        console.error(err);
        res.send("Error executing query: " + err.message);
    } finally {
        if (connection) await connection.close();
    }
}

module.exports = router;
