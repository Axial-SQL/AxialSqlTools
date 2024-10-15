SET NOCOUNT ON;
DECLARE @RowsStatistics AS TABLE (ObjectId INT, PartitionCount INT, RowsCount BIGINT, UnusedPagesPercent INT, INDEX IDX CLUSTERED (ObjectId));
INSERT INTO @RowsStatistics
SELECT ps.object_id AS ObjectId,
       COUNT(DISTINCT ps.partition_number) AS PartitionCount,
       SUM(ps.row_count) AS RowsCount,
       CASE WHEN SUM(ps.reserved_page_count) = 0 THEN 0 ELSE (SUM(ps.reserved_page_count) - SUM(ps.used_page_count)) * 100 / SUM(ps.reserved_page_count) END AS UnusedPagesPercent
FROM sys.dm_db_partition_stats AS ps
     INNER JOIN sys.objects AS so ON ps.object_id = so.object_id
WHERE ps.index_id < 2
GROUP BY ps.object_id;

DECLARE @TableInfo AS TABLE (SchemaName SYSNAME, TableName SYSNAME, ObjectId INT, HasPK INT, HasClusteredIndex INT, PKisClustered INT, IsReplicated INT, IndexCount INT, CreateDate DATETIME, LastIdentityValue SQL_VARIANT, INDEX IDX CLUSTERED (ObjectId));
INSERT INTO @TableInfo
SELECT 
       OBJECT_SCHEMA_NAME(t.object_id) AS SchemaName,
       OBJECT_NAME(t.object_id) AS TableName,
       t.object_id AS ObjectId,
       i.HasPK,
       i.HasClusteredIndex,
       i.PKisClustered,
       t.is_replicated as IsReplicated,
       COALESCE (i.IndexCount, 0) AS IndexCount,
       t.create_date,
       ic.last_value
FROM sys.tables AS t
     LEFT OUTER JOIN
     (SELECT si.object_id,
             COUNT(*) AS IndexCount,
             SUM(CASE WHEN si.type_desc = 'CLUSTERED' THEN 1 ELSE 0 END) AS HasClusteredIndex,
             SUM(CASE WHEN si.is_primary_key = 1 THEN 1 ELSE 0 END) AS HasPK,
             SUM(CASE WHEN si.type_desc = 'CLUSTERED'
                           AND si.is_primary_key = 1 THEN 1 ELSE 0 END) AS PKisClustered
      FROM sys.indexes AS si
      GROUP BY si.object_id) AS i
     ON t.object_id = i.object_id
     LEFT JOIN sys.identity_columns ic ON t.object_id = ic.object_id
WHERE t.is_ms_shipped = 0;

DECLARE @TableSizes AS TABLE (ObjectId INT, UsedSpaceMB NUMERIC (36, 2), UsedSpaceMB_Compressed NUMERIC (36, 2), UsedSpaceMB_LOB NUMERIC (36, 2), UsedSpaceMB_CS NUMERIC (36, 2), INDEX IDX CLUSTERED (ObjectId));
INSERT INTO @TableSizes
SELECT t.object_id AS ObjectId,
        CAST (ROUND(((SUM(a.used_pages) * 8) / 1024.00), 2) AS NUMERIC (36, 2)) AS UsedSpaceMB,
        CAST (ROUND(((SUM(CASE WHEN p.[data_compression] > 0 THEN a.used_pages ELSE 0 END) * 8) / 1024.00), 2) AS NUMERIC (36, 2)) AS UsedSpaceMB_Compressed,
        CAST (ROUND(((SUM(CASE WHEN a.[type_desc] = 'LOB_DATA' AND i.[type] in (1, 2) THEN a.used_pages ELSE 0 END) * 8) / 1024.00), 2) AS NUMERIC (36, 2)) AS UsedSpaceMB_LOB,
        CAST (ROUND(((SUM(CASE WHEN i.[type] in (5, 6) THEN a.used_pages ELSE 0 END) * 8) / 1024.00), 2) AS NUMERIC (36, 2)) AS UsedSpaceMB_CS
FROM sys.tables AS t
        INNER JOIN sys.indexes AS i ON t.object_id = i.object_id
        INNER JOIN sys.partitions AS p ON i.object_id = p.object_id AND i.index_id = p.index_id
        INNER JOIN sys.allocation_units AS a ON p.partition_id = a.container_id
GROUP BY t.object_id;

DECLARE @LastReadWrites AS TABLE (ObjectId INT, [LastWrite] DATETIME, [LastRead] DATETIME, [TotalReads] BIGINT, [TotalWrites] BIGINT, INDEX IDX CLUSTERED (ObjectId));
INSERT INTO @LastReadWrites
SELECT ObjectID AS ObjectID,
       MAX([LastUserUpdate]) AS [LastWrite],
       MAX([LastUserRead]) AS [LastRead],
       SUM([TotalReads]) AS [TotalReads],
       SUM([TotalWrites]) AS [TotalWrites]
FROM (SELECT object_id AS ObjectID,
             (last_user_update) AS [LastUserUpdate],             
             (SELECT Max(v) FROM (VALUES (last_user_seek), (last_user_scan), (last_user_lookup)) AS value(v)) AS [LastUserRead],
             (user_updates) AS [TotalWrites],
             (user_seeks + user_scans + user_lookups) AS [TotalReads]
      FROM sys.dm_db_index_usage_stats
      WHERE database_id = DB_ID()) AS a
GROUP BY ObjectID;

SELECT DB_NAME() AS DatabaseName,
       ti.SchemaName,
       ti.TableName,
       FORMAT(RS.RowsCount, 'N0') AS RowsCount,
       FORMAT(TS.UsedSpaceMB, 'N0') AS UsedSpaceMB,
       -- FORMAT(TS.UsedSpaceMB_LOB, 'N0') AS UsedSpaceMB_LOB,
       -- FORMAT(TS.UsedSpaceMB_CS, 'N0') AS UsedSpaceMB_CS,
       FORMAT(TS.UsedSpaceMB_Compressed, 'N0') AS UsedSpaceMB_ZIP,
       RS.PartitionCount,
       ti.IndexCount,
       ti.HasPK,
       ti.HasClusteredIndex,
       ti.PKisClustered,
       ti.IsReplicated,
       L.LastWrite,
       L.LastRead,
       L.TotalReads,
       L.TotalWrites,
       ti.CreateDate,
       RS.UnusedPagesPercent,
       ti.LastIdentityValue
FROM @TableInfo AS ti
	LEFT JOIN @RowsStatistics AS RS ON RS.ObjectId = ti.ObjectId
	LEFT JOIN @LastReadWrites AS L ON L.ObjectId = RS.ObjectId
	LEFT JOIN @TableSizes AS TS ON TS.ObjectId = ti.ObjectId
WHERE 1 = 1
     --AND ti.TableName IN ('','')
     --AND ti.SchemaName = ''
     --AND RS.RowsCount > 0
ORDER BY TS.UsedSpaceMB DESC;
-- ORDER BY TS.UsedSpaceMB - TS.UsedSpaceMB_LOB DESC;
