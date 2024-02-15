SELECT ec.client_net_address,
       es.[program_name],
       es.[host_name],
       es.login_name,
       ec.local_net_address,
       ec.local_tcp_port,
       DB_NAME(es.database_id) as DatabaseName,
       COUNT(ec.session_id) AS [connection count],
       SUM(CASE WHEN encrypt_option = 'TRUE' THEN 1 ELSE 0 END) AS [connection count (enc)]
FROM sys.dm_exec_sessions AS es WITH (NOLOCK)
     INNER JOIN
     sys.dm_exec_connections AS ec WITH (NOLOCK)
     ON es.session_id = ec.session_id
WHERE 1 = 1
	-- AND es.[program_name] LIKE '%%'
	-- AND DB_NAME(es.database_id) = ''
	-- AND es.[host_name] LIKE '%%'
	-- AND es.login_name = ''
GROUP BY 
     GROUPING SETS((), --totals
          (ec.client_net_address, es.[program_name], es.[host_name], es.login_name, ec.local_net_address, ec.local_tcp_port, DB_NAME(es.database_id)))
ORDER BY 
     ec.client_net_address, es.[program_name]
     -- [connection count] DESC
OPTION (RECOMPILE);
