using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace TKMK
{
    /// <summary>
    /// 團務佣金計算 (20261104)
    /// 由 frmGROUPSALESCOMMISSION 抽出來的共用計算，frmGROUPSALESCOMMISSION 與 frmGROUPSALESBYTA008 共用同一份
    /// 公式與原本 frmGROUPSALESBYTA008.SETMONEYS_NEW 相同，做法：
    /// 1. 只算「預約接團」的車 (只有預約接團會被回寫)
    /// 2. 先查出當天車的日期、業務員帳號，用常值 IN(...) + OPTION(RECOMPILE) 抓 POS 進 #TA/#TB/#TC
    /// 3. 每台車基本數字先存 #BASE (每台只算一次)，再算抽佣/獎金進 #TempTable
    /// 4. 交易只包最後兩個 UPDATE
    /// 注意：建立 #暫存表 的 SQL 不能帶參數 (會變成 sp_executesql，暫存表執行完就消失)
    /// SQL 語法限 SQL Server 2008
    /// </summary>
    public static class GROUPSALESCOMMISSIONCAL
    {
        private const int SQL_TIMEOUT = 300;

        /// <summary>
        /// 執行 SQL (不回傳資料)，回傳影響筆數
        /// </summary>
        public static int EXEC(SqlConnection conn, SqlTransaction trans, string sql, List<SqlParameter> pars)
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            {
                cmd.CommandTimeout = SQL_TIMEOUT;
                if (pars != null)
                {
                    foreach (SqlParameter p in pars)
                    {
                        cmd.Parameters.Add(((ICloneable)p).Clone());
                    }
                }
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 執行 SQL，回傳 DataTable
        /// </summary>
        public static DataTable QUERY(SqlConnection conn, SqlTransaction trans, string sql, List<SqlParameter> pars)
        {
            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            {
                cmd.CommandTimeout = SQL_TIMEOUT;
                if (pars != null)
                {
                    foreach (SqlParameter p in pars)
                    {
                        cmd.Parameters.Add(((ICloneable)p).Clone());
                    }
                }
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        /// <summary>
        /// 字串轉成 SQL 常值 'xxx' (單引號跳脫)
        /// 不加 N：日期、帳號都是英數字；若 POS 欄位是 varchar，N'' 會讓欄位被轉型而用不到索引
        /// </summary>
        private static string SQLSTR(string value)
        {
            return "'" + value.Replace("'", "''") + "'";
        }

        /// <summary>
        /// 產生 IN ('a','b',...) 的常值清單
        /// 注意：建立 #暫存表 的 SQL 不能帶參數！
        /// 帶參數的 SqlCommand 會用 sp_executesql 執行，裡面建立的 #暫存表 在執行完就被刪掉，
        /// 下一個命令就會出現「找不到 #TA」。所以建暫存表的 SQL 一律用常值。
        /// </summary>
        private static string INLIST(List<string> values)
        {
            return string.Join(",", values.Select(v => SQLSTR(v)).ToArray());
        }

        /// <summary>
        /// 計算佣金 (可在背景執行緒呼叫，不碰畫面)
        /// CONNSTRING：解密後的連線字串
        /// 回傳每一步的計時紀錄；失敗會丟出 Exception (回寫部分會 Rollback)
        /// </summary>
        public static List<string> CALMONEYS(string CONNSTRING, string SDATES)
        {
            List<string> LOGS = new List<string>();
            Stopwatch TOTAL = Stopwatch.StartNew();
            Stopwatch SW = Stopwatch.StartNew();

            DateTime D1 = DateTime.ParseExact(SDATES, "yyyyMMdd", null);
            DateTime D2 = D1.AddDays(1);

            using (SqlConnection conn = new SqlConnection(CONNSTRING))
            {
                conn.Open();
                LOGS.Add(string.Format("連線資料庫：{0:0.00} 秒", SW.Elapsed.TotalSeconds));

                //---------------------------------------------------------------
                //1. 當天要計算的車 (只算 預約接團；其他狀態原本也不會被更新)
                //---------------------------------------------------------------
                SW.Restart();
                List<SqlParameter> P1 = new List<SqlParameter>();
                P1.Add(new SqlParameter("@D1", SqlDbType.DateTime) { Value = D1 });
                P1.Add(new SqlParameter("@D2", SqlDbType.DateTime) { Value = D2 });

                DataTable CARS = QUERY(conn, null, @"
                                    SELECT [ID],[TA008]
                                    ,CONVERT(varchar(8),[GROUPSTARTDATES],112) AS SYMD
                                    ,CONVERT(varchar(8),[CREATEDATES],112) AS CYMD
                                    FROM [TKMK].[dbo].[GROUPSALES] WITH(NOLOCK)
                                    WHERE [CREATEDATES]>=@D1 AND [CREATEDATES]<@D2
                                    AND [STATUS] IN ('預約接團')
                                    ", P1);

                List<string> DATES = new List<string>();
                List<string> TA008S = new List<string>();
                foreach (DataRow dr in CARS.Rows)
                {
                    foreach (string col in new[] { "SYMD", "CYMD" })
                    {
                        if (dr[col] != DBNull.Value && !DATES.Contains(dr[col].ToString()))
                        {
                            DATES.Add(dr[col].ToString());
                        }
                    }
                    if (dr["TA008"] != DBNull.Value && !TA008S.Contains(dr["TA008"].ToString()))
                    {
                        TA008S.Add(dr["TA008"].ToString());
                    }
                }
                LOGS.Add(string.Format("1.查預約接團的車：{0} 台，{1:0.00} 秒", CARS.Rows.Count, SW.Elapsed.TotalSeconds));

                //---------------------------------------------------------------
                //2~5. 有預約接團的車才需要抓 POS、計算
                //---------------------------------------------------------------
                //沒有 POS 資料的車也要算 (茶水費、領券額只看車種)，所以只要有車就計算
                bool HASCARS = CARS.Rows.Count > 0;

                if (HASCARS)
                {
                    EXEC(conn, null, @"
                                    IF OBJECT_ID('tempdb..#TempTable') IS NOT NULL DROP TABLE #TempTable;
                                    IF OBJECT_ID('tempdb..#TA') IS NOT NULL DROP TABLE #TA;
                                    IF OBJECT_ID('tempdb..#TB') IS NOT NULL DROP TABLE #TB;
                                    IF OBJECT_ID('tempdb..#TC') IS NOT NULL DROP TABLE #TC;
                                    IF OBJECT_ID('tempdb..#BASE') IS NOT NULL DROP TABLE #BASE;
                                    ", null);

                    //2. POSTA：實際日期 + 實際業務員帳號 (常值)
                    SW.Restart();
                    //清單是空的就用 1=0 (抓不到資料，但後面的計算照樣跑)
                    //建 #暫存表 的 SQL 不能帶參數 (見 INLIST 說明)，所以用常值
                    string WHEREDATES_TA = DATES.Count > 0 ? "TA001 IN (" + INLIST(DATES) + ")" : "1=0";
                    string WHERETA008 = TA008S.Count > 0 ? "TA008 IN (" + INLIST(TA008S) + ")" : "1=0";
                    string WHEREDATES_TB = WHEREDATES_TA.Replace("TA001", "TB001");
                    string WHEREDATES_TC = WHEREDATES_TA.Replace("TA001", "TC001");

                    int NTA = EXEC(conn, null, string.Format(@"
                                    SELECT TA001,TA002,TA003,TA005,TA006,TA008,TA017
                                    INTO #TA
                                    FROM [TK].dbo.POSTA WITH (NOLOCK)
                                    WHERE {0}
                                    AND {1}
                                    AND TA002 IN (SELECT [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
                                    OPTION (RECOMPILE);
                                    ", WHEREDATES_TA, WHERETA008), null);
                    EXEC(conn, null, "CREATE CLUSTERED INDEX IX_TA ON #TA (TA008,TA001,TA005);", null);
                    LOGS.Add(string.Format("2.抓 POSTA：{0} 筆，{1:0.00} 秒", NTA, SW.Elapsed.TotalSeconds));

                    //3. POSTB：只抓上面那些交易的明細
                    SW.Restart();
                    int NTB = EXEC(conn, null, string.Format(@"
                                    SELECT TB001,TB002,TB003,TB006,TB010,TB019,TB033
                                    INTO #TB
                                    FROM [TK].dbo.POSTB WITH (NOLOCK)
                                    WHERE {0}
                                    AND EXISTS (SELECT 1 FROM #TA WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006)
                                    OPTION (RECOMPILE);
                                    ", WHEREDATES_TB), null);
                    EXEC(conn, null, "CREATE CLUSTERED INDEX IX_TB ON #TB (TB001,TB002,TB003,TB006);", null);
                    LOGS.Add(string.Format("3.抓 POSTB：{0} 筆，{1:0.00} 秒", NTB, SW.Elapsed.TotalSeconds));

                    //4. POSTC：付款方式 0009 (兌換券)
                    SW.Restart();
                    int NTC = EXEC(conn, null, string.Format(@"
                                    SELECT TC001,TC002,TC003,TC006,TC008
                                    INTO #TC
                                    FROM [TK].dbo.POSTC WITH (NOLOCK)
                                    WHERE {0}
                                    AND TC008='0009'
                                    AND EXISTS (SELECT 1 FROM #TA WHERE TA001=TC001 AND TA002=TC002 AND TA003=TC003 AND TA006=TC006)
                                    OPTION (RECOMPILE);
                                    ", WHEREDATES_TC), null);
                    EXEC(conn, null, "CREATE CLUSTERED INDEX IX_TC ON #TC (TC001,TC002,TC003,TC006);", null);
                    LOGS.Add(string.Format("4.抓 POSTC：{0} 筆，{1:0.00} 秒", NTC, SW.Elapsed.TotalSeconds));

                    //5. 佣金計算 (公式與 frmGROUPSALESBYTA008.SETMONEYS_NEW 相同)
                    //建 #BASE / #TempTable 的 SQL 不能帶參數 (見 INLIST 說明)，日期改成常值 'yyyyMMdd'
                    string SQL5A = SQL_CAL_BASE
                        .Replace("@D1", "'" + D1.ToString("yyyyMMdd") + "'")
                        .Replace("@D2", "'" + D2.ToString("yyyyMMdd") + "'");

                    //5a. 每台車基本數字 → #BASE
                    SW.Restart();
                    int NBASE = EXEC(conn, null, SQL5A, null);
                    LOGS.Add(string.Format("5a.計算每台車基本數字：{0} 台，{1:0.00} 秒", NBASE, SW.Elapsed.TotalSeconds));

                    //5b. 抽佣比率、獎金 → #TempTable
                    SW.Restart();
                    int NCAL = EXEC(conn, null, SQL_CAL, null);
                    LOGS.Add(string.Format("5b.計算抽佣/獎金：{0} 台，{1:0.00} 秒", NCAL, SW.Elapsed.TotalSeconds));
                }
                else
                {
                    LOGS.Add("2~5.沒有預約接團的車，略過 POS 與佣金計算");
                }

                //---------------------------------------------------------------
                //6. 回寫 GROUPSALES (只有這一段用交易)
                //---------------------------------------------------------------
                SW.Restart();
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        List<SqlParameter> P6 = new List<SqlParameter>();
                        P6.Add(new SqlParameter("@D1", SqlDbType.DateTime) { Value = D1 });
                        P6.Add(new SqlParameter("@D2", SqlDbType.DateTime) { Value = D2 });

                        int NUP1 = 0;
                        if (HASCARS)
                        {
                            NUP1 = EXEC(conn, trans, SQL_UPDATE1, P6);
                        }
                        int NUP2 = EXEC(conn, trans, SQL_UPDATE2, P6);

                        trans.Commit();
                        LOGS.Add(string.Format("6.回寫 GROUPSALES：佣金 {0} 筆、空車伴手禮 {1} 筆，{2:0.00} 秒", NUP1, NUP2, SW.Elapsed.TotalSeconds));
                    }
                    catch
                    {
                        try { trans.Rollback(); } catch { }
                        throw;
                    }
                }
            }

            LOGS.Add(string.Format("合計：{0:0.00} 秒", TOTAL.Elapsed.TotalSeconds));
            return LOGS;
        }

        #region SQL

        /// <summary>
        /// 佣金計算：與 frmGROUPSALESBYTA008.SETMONEYS_NEW 相同公式
        /// 只差在 POSTA/POSTB/POSTC 改用 #TA/#TB/#TC，且只算 預約接團
        /// 5a. 每台車的基本數字 (消費總額、特賣、券消費、領券額、茶水費、交易筆數)
        ///     20261104 先存成 #BASE：原本是巢狀子查詢，外層每引用一次 SALESMMONEYS 等欄位，
        ///     SQL Server 可能就重算一次相關子查詢；先存起來每台車只算一次
        /// </summary>
        private const string SQL_CAL_BASE = @"
                                    IF OBJECT_ID('tempdb..#BASE') IS NOT NULL DROP TABLE #BASE;

		                                    SELECT  
		                                    [ID]
		                                    ,[SERNO] AS '序號'
		                                    ,[CARNAME] AS '車名'
		                                    ,[CARNO] AS '車號'
		                                    ,[CARKIND] AS '車種'
		                                    ,[GROUPKIND]  AS '團類'
		                                    ,[ISEXCHANGE] AS '兌換券'
		                                    ,[EXCHANGETOTALMONEYS] AS '券總額'
		                                    ,[EXCHANGESALESMMONEYS] AS '券消費'
		                                    ,[SALESMMONEYS] AS '消費總額'
		                                    ,[SPECIALMNUMS] AS '特賣數'
		                                    ,[SPECIALMONEYS] AS '特賣獎金'
		                                    ,[COMMISSIONBASEMONEYS] AS '茶水費'
		                                    ,[COMMISSIONPCTMONEYS] AS '消費獎金'
		                                    ,[TOTALCOMMISSIONMONEYS] AS '總獎金'
		                                    ,[CARNUM] AS '車數'
		                                    ,[GUSETNUM] AS '交易筆數'
		                                    ,[CARCOMPANY] AS '來車公司'
		                                    ,[TA008NO] AS '業務員名'
		                                    ,[TA008] AS '業務員帳號'
		                                    ,[EXCHANNO] AS '優惠券名'
		                                    ,[EXCHANACOOUNT] AS '優惠券帳號'
		                                    ,CONVERT(varchar(100), [GROUPSTARTDATES],120) AS '實際到達時間'
		                                    ,CONVERT(varchar(100), [GROUPENDDATES],120) AS '實際離開時間'
		                                    ,[STATUS] AS '狀態'
		                                    ,CONVERT(varchar(100), [PURGROUPSTARTDATES],120) AS '預計到達時間'
		                                    ,CONVERT(varchar(100), [PURGROUPENDDATES],120) AS '預計離開時間'
		                                    ,[EXCHANGEMONEYS] AS '領券額'
		                                    ,[CREATEDATES]
		                                    ,CONVERT(varchar, [GROUPSTARTDATES], 112) AS 'YMD'
		                                    ,CONVERT(varchar,[GROUPSTARTDATES], 108) AS 'STARTHM'
		                                    ,(
		                                    SELECT CONVERT(INT,ISNULL(SUM(TB033),0),0)
		                                    FROM #TA,#TB
		                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003  AND TA006=TB006  
		                                    AND TB010  NOT IN (SELECT [ID] FROM [TKMK].[dbo].[GROUPPRODUCT] WHERE [VALID]='Y' AND [SPLITCAL]='Y')              
		                                    AND TA008=[GROUPSALES].[TA008] 
		                                    AND TA001=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112) 
		                                    AND TA005>=CONVERT(varchar, [GROUPSALES].[GROUPSTARTDATES],108 )

		                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
		                                    )  AS 'SALESMMONEYS'
		                                    ,
		                                    (
		                                    SELECT ISNULL(SUM(SUMTB019),0) 
		                                    FROM 
		                                    (
		                                    SELECT [ID],[NAME],[NUM],[MONEYS],[SPLITCAL],[VALID],[SDATES],[EDATES],TB010,SUMTB019
		                                    FROM [TKMK].[dbo].[GROUPPRODUCT]
		                                    LEFT JOIN 
		                                    (
		                                    SELECT TB010,CONVERT(INT,ISNULL(SUM(TB019),0),0) SUMTB019
		                                    FROM #TA,#TB
		                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006 
		                                    AND TA008=[GROUPSALES].[TA008]  
		                                    AND TA001=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112)  
		                                    AND TA005>=CONVERT(varchar, [GROUPSALES].[GROUPSTARTDATES],108 )

		                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
		                                    GROUP BY TB010
		                                    ) AS TEMP ON TB010=ID
		                                    WHERE [VALID]='Y' 
		                                    AND CONVERT(NVARCHAR,SDATES,112)<=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112) 
		                                    AND CONVERT(NVARCHAR,EDATES,112)>=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112) 
		                                    ) AS TEMP2
		                                    ) AS 'FINALSPECIALMNUMS'
		                                    ,
		                                    (
		                                    SELECT ISNULL(SUM(SUMTB019/[NUM]*[MONEYS]),0)
		                                    FROM 
		                                    (
		                                    SELECT [ID],[NAME],[NUM],[MONEYS],[SPLITCAL],[VALID],[SDATES],[EDATES],TB010,SUMTB019
		                                    FROM [TKMK].[dbo].[GROUPPRODUCT]
		                                    LEFT JOIN 
		                                    (
		                                    SELECT TB010,CONVERT(INT,ISNULL(SUM(TB019),0),0) SUMTB019
		                                    FROM #TA,#TB
		                                    WHERE TA001=TB001 AND TA002=TB002 AND TA003=TB003 AND TA006=TB006 
		                                    AND TA008=[GROUPSALES].[TA008]  
		                                    AND TA001=CONVERT(varchar,[GROUPSALES].[CREATEDATES], 112)  
		                                    AND TA005>=CONVERT(varchar, [GROUPSALES].[CREATEDATES],108 )

		                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
		                                    GROUP BY TB010
		                                    ) AS TEMP ON TB010=ID
		                                    WHERE [VALID]='Y' 
		                                    AND CONVERT(NVARCHAR,SDATES,112)<=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112) 
		                                    AND CONVERT(NVARCHAR,EDATES,112)>=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112) 
		                                    ) AS TEMP2
		                                    ) AS 'FINALSPECIALMONEYS' 
		                                    ,
		                                    (
		                                    SELECT CONVERT(INT,ISNULL(SUM(TA017),0))
		                                    FROM #TA,#TC
		                                    WHERE TA001=TC001 AND TA002=TC002 AND TA003=TC003  AND TA006=TC006
		                                    AND TC008='0009'
		                                    AND TA008=[GROUPSALES].[TA008] 
		                                    AND TA001=CONVERT(varchar,[GROUPSALES].[CREATEDATES], 112)
		                                    AND TA005>=CONVERT(varchar,[GROUPSALES].[CREATEDATES], 108)

		                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
		                                    ) AS 'FINALEXCHANGESALESMMONEYS'
		                                    ,
		                                    ISNULL((
		                                    SELECT  
		                                    CONVERT(INT,[EXCHANGEMONEYS],0)
		                                    FROM [TKMK].[dbo].[GROUPEXCHANGEMONEYS]
		                                    WHERE  1=1
		                                    AND [VALID]='Y'
		                                    AND [CARKIND]=[GROUPSALES].[CARKIND] 
		                                    AND CONVERT(NVARCHAR,SDATES,112)<=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112)
		                                    AND CONVERT(NVARCHAR,EDATES,112)>=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112)
		                                    ),0)  AS 'EXCHANGEMONEYS'  
		                                    ,
		                                    (
		                                    SELECT CONVERT(INT,[BASEMONEYS],0)
		                                    FROM [TKMK].[dbo].[GROUPBASE] 
		                                    WHERE 1=1
		                                    AND VALID='Y'
		                                    AND CONVERT(NVARCHAR,SDATES,112)<=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112)
		                                    AND CONVERT(NVARCHAR,EDATES,112)>=CONVERT(varchar,[GROUPSALES].[GROUPSTARTDATES], 112)
		                                    AND [NAME]=[GROUPSALES].[CARKIND] 
		                                    )  AS 'FINALBASEMONEYS' 
		                                    ,
		                                    (
		                                    SELECT COUNT(TA008) 
		                                    FROM #TA
		                                    WHERE TA008=[GROUPSALES].[TA008] 
		                                    AND TA001=CONVERT(varchar,[GROUPSALES].[CREATEDATES], 112)
		                                    AND TA005>=CONVERT(varchar,[GROUPSALES].[CREATEDATES], 108)
		
		                                    AND TA002 IN (SELECT  [TA002] FROM [TKMK].[dbo].[GROUPSTORES] WHERE KINDNAMES IN ('GROUPSTORES1'))
		                                    ) AS 'GUSETNUM'
		                                    INTO #BASE
		                                    FROM [TKMK].[dbo].[GROUPSALES]
		                                    WHERE [CREATEDATES]>=@D1 AND [CREATEDATES]<@D2
		                                    AND [STATUS] IN ('預約接團')   --只算預約接團 (只有預約接團會被回寫)

";

        /// <summary>
        /// 5b. 依 #BASE 算抽佣比率、獎金，結果放 #TempTable (公式與原本相同)
        /// </summary>
        private const string SQL_CAL = @"
                                    SELECT *
                                    ,(CASE WHEN 兌換券='是' THEN EXCHANGEMONEYS ELSE 0 END ) AS 'FINALEXCHANGEMONEYS'
                                    ,(CASE WHEN 兌換券='否' THEN FINALBASEMONEYS ELSE 0 END ) AS 'FINALCOMMISSIONBASEMONEYS'
                                    ,CONVERT(INT,FINNALSALESMMONEYS*FINALCOMMISSIONPCT,0) AS 'FINNALCOMMISSIONPCTMONEYS'
                                    
                                    --總將金=抽佣+茶水費+特賣獎金
                                    ,(CONVERT(INT,FINNALSALESMMONEYS*FINALCOMMISSIONPCT,0)+(CASE WHEN 兌換券='否' THEN FINALBASEMONEYS ELSE 0 END )+FINALSPECIALMONEYS) AS 'FINALTOTALCOMMISSIONMONEYS'

                                    INTO #TempTable
                                    FROM 
                                    (
	                                    SELECT *
	                                    FROM 
	                                    (
		                                    SELECT *
		                                    --如果司機有兌換券，而且有消費金額，可當佣金的消費金額=原消費總額-券總額(保留超過的部份金額還是可以抽佣)
		                                    ,(CASE WHEN FINALEXCHANGESALESMMONEYS>=1 AND SALESMMONEYS>=1 AND (SALESMMONEYS-EXCHANGEMONEYS)>=0 THEN (SALESMMONEYS-EXCHANGEMONEYS) ELSE SALESMMONEYS  END )  AS 'FINNALSALESMMONEYS'
		                                    FROM
		                                    #BASE AS TEMP
	                                    ) AS TEMP2
	                                    OUTER  APPLY (
	                                    SELECT TOP 1 [PCT]  AS 'FINALCOMMISSIONPCT'              
	                                    FROM [TKMK].[dbo].[GROUPPCT]
	                                    WHERE [NAME] = TEMP2.車種
	                                    AND (FINNALSALESMMONEYS - [PCTMONEYS]) >= 0
	                                    AND VALID = 'Y'
	                                    AND CONVERT(NVARCHAR, SDATES, 112) <= CONVERT(VARCHAR, TEMP2.YMD, 112)
	                                    AND CONVERT(NVARCHAR, EDATES, 112) >= CONVERT(VARCHAR, TEMP2.YMD, 112)
	                                    ORDER BY (FINNALSALESMMONEYS - [PCTMONEYS]) 
	                                    ) AS OUTERPCT
                                    ) AS TEMP3
";

        /// <summary>
        /// 回寫佣金 (只更新 預約接團)
        /// </summary>
        private const string SQL_UPDATE1 = @"
                                    UPDATE [TKMK].[dbo].[GROUPSALES]
                                    SET 
                                    [EXCHANGEMONEYS]=#TempTable.FINALEXCHANGEMONEYS
                                    ,[EXCHANGETOTALMONEYS]=#TempTable.FINALEXCHANGEMONEYS*#TempTable.車數
                                    ,[EXCHANGESALESMMONEYS]=#TempTable.FINALEXCHANGESALESMMONEYS
                                    ,[SPECIALMNUMS]=#TempTable.FINALSPECIALMNUMS
                                    ,[SPECIALMONEYS]=#TempTable.FINALSPECIALMONEYS
                                    ,[SALESMMONEYS]=#TempTable.FINNALSALESMMONEYS
                                    ,[COMMISSIONBASEMONEYS]=#TempTable.FINALCOMMISSIONBASEMONEYS
                                    ,[COMMISSIONPCT]=#TempTable.FINALCOMMISSIONPCT
                                    ,[COMMISSIONPCTMONEYS]=#TempTable.FINNALCOMMISSIONPCTMONEYS
                                    ,[TOTALCOMMISSIONMONEYS]=#TempTable.FINALTOTALCOMMISSIONMONEYS
                                    ,[GUSETNUM]=#TempTable.GUSETNUM
                                    FROM #TempTable
                                    WHERE [GROUPSALES].[CREATEDATES]>=@D1 AND [GROUPSALES].[CREATEDATES]<@D2
                                    AND [GROUPSALES].[STATUS] IN ('預約接團')
                                    AND [GROUPSALES].ID=#TempTable.ID;
";

        /// <summary>
        /// 空車伴手禮：消費 1~4999 抽 10%，不給茶水費 (與原本相同)
        /// </summary>
        private const string SQL_UPDATE2 = @"
                                    UPDATE [TKMK].[dbo].[GROUPSALES]
                                    SET [COMMISSIONBASEMONEYS]=0
                                    ,[COMMISSIONPCT]=0.1
                                    ,[COMMISSIONPCTMONEYS]=[SALESMMONEYS]*0.1
                                    ,[TOTALCOMMISSIONMONEYS]=[SALESMMONEYS]*0.1
                                    WHERE [CARKIND]='空車伴手禮' 
                                    AND [SALESMMONEYS]>=1 AND [SALESMMONEYS]<=4999 
                                    AND [GROUPSALES].[CREATEDATES]>=@D1 AND [GROUPSALES].[CREATEDATES]<@D2;
";

        #endregion
    }
}
