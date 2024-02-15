WITH Blocking
AS (SELECT [session_Id] AS [Spid],
           percent_complete,
           blocked,
           waittime,
           DB_NAME(sp.[dbid]) AS [Database],
           nt_username AS [User],
           er.[status] AS [Status],
           wait_type AS [Wait],
           SUBSTRING(qt.text, er.statement_start_offset / 2, (CASE WHEN er.statement_end_offset = -1 THEN LEN(CONVERT (NVARCHAR (MAX), qt.text)) * 2 ELSE er.statement_end_offset END - er.statement_start_offset) / 2) AS [Individual Query],
           qt.[text] AS [Parent Query],
           program_name AS Program,
           Hostname,
           cpu_time,
           reads,
           start_time,
           sp.login_time,
           sp.last_batch,
           sp.cmd
    FROM sys.dm_exec_requests AS er
         INNER JOIN sys.sysprocesses AS sp ON er.[session_id] = sp.spid 
		 CROSS APPLY sys.dm_exec_sql_text(er.[sql_handle]) AS qt
    WHERE [session_Id] > 50 AND [session_Id] NOT IN (@@SPID))
SELECT * FROM blocking WHERE blocked <> 0
UNION ALL
SELECT * FROM blocking WHERE spid IN (SELECT blocked FROM blocking WHERE blocked <> 0);

