SELECT SCHEMA_NAME(o.schema_id) AS SchemaName,
       OBJECT_NAME(p.object_id) AS ObjectName,
       i.[name] AS IndexName,
       p.index_id AS IndexID,
       p.partition_number AS PartitionNumber,
       FORMAT(p.[rows], 'N0') AS [Rows],
       FORMAT(a.total_pages * 8 / 1024., 'N2') AS IndexSizeMb, --doesn't include LOB
       p.[data_compression_desc] AS [Compression],
       CONCAT('ALTER INDEX [', i.[name], '] ON [', SCHEMA_NAME(o.schema_id), '].[', OBJECT_NAME(p.object_id), '] REBUILD PARTITION = ', 
              CASE WHEN EXISTS (SELECT TOP 1 1 FROM sys.partition_schemes AS ps WHERE ps.data_space_id = i.data_space_id)
              THEN  CAST (p.partition_number AS VARCHAR (4)) ELSE 'ALL' END,
              ' WITH (ONLINE = ON (WAIT_AT_LOW_PRIORITY (MAX_DURATION = 15 MINUTES, ABORT_AFTER_WAIT = SELF)), ', 
              'SORT_IN_TEMPDB = ON, DATA_COMPRESSION=PAGE, MAXDOP = 8);')  + CHAR(13) + CHAR(10) + 'GO' AS RebuildCommand
FROM sys.partitions AS p
     INNER JOIN sys.indexes AS i ON i.[object_id] = p.[object_id] AND i.index_id = p.index_id
     INNER JOIN sys.tables AS o ON i.[object_id] = o.[object_id]
     INNER JOIN sys.data_spaces AS ds ON ds.data_space_id = i.data_space_id
     LEFT OUTER JOIN sys.allocation_units AS a ON p.[partition_id] = a.container_id AND a.data_pages > 0 AND a.total_pages > 0 -- B-tree only I guess
WHERE 1 = 1
      -- AND i.index_id = 1
      -- AND OBJECT_NAME(p.[object_id]) = ''
      -- AND SCHEMA_NAME(o.[schema_id]) = ''
      -- AND i.[name] = ''
      -- AND p.partition_number = 0
      -- AND p.[rows] > 0
ORDER BY ObjectName, PartitionNumber, IndexID;

