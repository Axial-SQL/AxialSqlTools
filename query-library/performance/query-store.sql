SELECT DATEADD(hour, DATEDIFF(hour, 0, rs.first_execution_time), 0) AS first_execution,
       sum(rs.count_executions) AS count_executions,
       sum(rs.count_executions * rs.avg_duration) AS total_duration,
       sum(rs.count_executions * rs.avg_duration) / SUM(rs.count_executions) AS avg_duration,
       sum(rs.count_executions * rs.avg_cpu_time) / SUM(rs.count_executions) AS avg_cpu_time,
       sum(rs.count_executions * rs.avg_logical_io_reads) AS total_logical_io_reads,
       sum(rs.count_executions * rs.avg_logical_io_reads) / SUM(rs.count_executions) AS avg_logical_io_reads,
       sum(rs.count_executions * rs.avg_physical_io_reads) AS total_physical_io_reads,
       sum(rs.count_executions * rs.avg_physical_io_reads) / SUM(rs.count_executions) AS avg_physical_io_reads
FROM sys.query_store_runtime_stats AS rs
WHERE rs.plan_id IN (SELECT p.plan_id
                     FROM sys.query_store_plan AS p
                     WHERE p.query_id IN (SELECT q.query_id
                                          FROM sys.query_store_query AS q
                                          WHERE q.object_id = object_id(' %% dbo.Table %% ')))
      AND rs.first_execution_time > (GETDATE() - 7)
GROUP BY DATEADD(hour, DATEDIFF(hour, 0, rs.first_execution_time), 0)
ORDER BY 1 DESC;


SELECT CAST(query_plan as XML) as queryPlan, *
FROM sys.query_store_plan AS p
WHERE p.query_id IN (SELECT q.query_id
				  FROM sys.query_store_query AS q
				  WHERE q.object_id = object_id(' %% dbo.Table %% '))
