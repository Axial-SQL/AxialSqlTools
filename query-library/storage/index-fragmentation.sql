SELECT s.[name] AS SchemaName,
       t.[name] AS TableName,
       i.[index_id] AS IndexID,
       i.[name] AS IndexName,
       ips.partition_number AS PartitionNumber,
       part.data_compression_desc AS DataComprDesc,
       FORMAT(ips.avg_fragmentation_in_percent, 'N2') AS AvgFragPercent,
       FORMAT(ips.page_count, 'N0') AS [PageCount],
       FORMAT(ips.page_count * 8 / 1024, 'N0') AS IndexSizeMb,
       CASE WHEN i.[type_desc] IN ('CLUSTERED', 'NONCLUSTERED') 
	   THEN CONCAT('ALTER INDEX [', i.[name], '] ON [', s.[name], '].[', t.[name], '] REORGANIZE PARTITION = ', 
				IIF (ps.[name] IS NULL, 'ALL', CAST (ips.partition_number AS VARCHAR (4))), ';') ELSE '' END AS CmdReorg,
       CASE WHEN i.[type_desc] IN ('HEAP') 
	   THEN CONCAT('ALTER TABLE [', s.[name], '].[', t.[name], '] REBUILD PARTITION = ', 
				IIF (ps.[name] IS NULL, 'ALL', CAST (ips.partition_number AS VARCHAR (4))), 
				' WITH (ONLINE = ON (WAIT_AT_LOW_PRIORITY (MAX_DURATION = 1 MINUTES, ABORT_AFTER_WAIT = SELF)), DATA_COMPRESSION=PAGE, MAXDOP = 4);') 
       WHEN i.[type_desc] IN ('CLUSTERED', 'NONCLUSTERED') 
	   THEN CONCAT('ALTER INDEX [', i.[name], '] ON [', s.[name], '].[', t.[name], '] REBUILD PARTITION = ', 
				IIF (ps.[name] IS NULL, 'ALL', CAST (ips.partition_number AS VARCHAR (4))), 
				' WITH (ONLINE = ON (WAIT_AT_LOW_PRIORITY (MAX_DURATION = 1 MINUTES, ABORT_AFTER_WAIT = SELF)), ', 
				'SORT_IN_TEMPDB = ON, DATA_COMPRESSION=PAGE, MAXDOP = 4);') END AS CmdRebuild
FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, -- OBJECT_ID('dbo.Table'),
								NULL, NULL, NULL) AS ips
     INNER JOIN sys.tables AS t ON t.[object_id] = ips.[object_id]
     INNER JOIN sys.schemas AS s ON t.[schema_id] = s.[schema_id]
     INNER JOIN sys.indexes AS i ON i.[object_id] = ips.[object_id] AND ips.index_id = i.index_id
     INNER JOIN sys.partitions AS part ON i.[object_id] = part.[object_id] AND i.index_id = part.index_id AND ips.partition_number = part.partition_number
     LEFT OUTER JOIN sys.partition_schemes AS ps ON i.data_space_id = ps.data_space_id
WHERE 1 = 1
	-- AND ips.page_count > 100
	-- AND i.[name] IS NOT NULL
	-- AND ips.avg_fragmentation_in_percent > 50
	-- AND part.data_compression_desc = 'NONE'
ORDER BY 
	-- ips.avg_fragmentation_in_percent DESC;
	SchemaName, TableName, PartitionNumber, IndexID;
