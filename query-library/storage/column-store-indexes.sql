/*
-- select all columnstore indexes
SELECT 
  OBJECT_SCHEMA_NAME(OBJECT_ID) SchemaName,
  OBJECT_NAME(OBJECT_ID) TableName,
  i.name AS IndexName, 
  i.type_desc IndexType
FROM sys.indexes AS i 
WHERE is_hypothetical = 0 
  AND i.index_id <> 0 
  AND i.type_desc IN ('CLUSTERED COLUMNSTORE','NONCLUSTERED COLUMNSTORE');
  
-- maintenance
ALTER INDEX [IndexName] ON [TableName] REORGANIZE WITH (COMPRESS_ALL_ROW_GROUPS = ON);  
ALTER INDEX [IndexName] ON [TableName] REBUILD PARTITION = ALL WITH (DATA_COMPRESSION = COLUMNSTORE_ARCHIVE, MAXDOP = 1);

-- get row group physical stats
SELECT *   
FROM sys.dm_db_column_store_row_group_physical_stats   
WHERE object_id  = object_id('TableName')  
ORDER BY row_group_id;  
*/
-- get some stats around CIX - helps to see which partitions need to be rebuilt
SELECT OBJECT_SCHEMA_NAME(ps.[object_id]) AS schemaName,
       OBJECT_NAME(ps.[object_id]) AS tableName,
       ps.partition_number,
       p.data_compression_desc,
       FORMAT(SUM(ps.row_group_id), 'N0') AS row_group_count,
       FORMAT(SUM(ps.total_rows), 'N0') AS row_count,
       FORMAT(SUM(ps.deleted_rows), 'N0') AS deleted_rows,
       FORMAT(SUM(ps.size_in_bytes) / 1024 / 1024, 'N0') AS size_in_mbytes,
       FORMAT(SUM(ps.size_in_bytes) / sum(total_rows), 'N0') AS size_per_row,
       CONCAT('ALTER INDEX [', i.[name], '] ON ', OBJECT_SCHEMA_NAME(ps.[object_id]), '.', OBJECT_NAME(ps.[object_id]), 
              ' REBUILD PARTITION = ', ps.partition_number, ' WITH (DATA_COMPRESSION = COLUMNSTORE, ONLINE = ON, SORT_IN_TEMPDB = ON /*, MAXDOP = 8 */);
GO') AS rebuiltCommand
FROM sys.dm_db_column_store_row_group_physical_stats AS ps
     INNER JOIN sys.indexes AS i ON ps.[object_id] = i.[object_id] AND ps.index_id = i.index_id
     INNER JOIN sys.partitions p ON ps.partition_number = p.partition_number AND ps.[object_id] = p.[object_id] AND ps.index_id = p.index_id
GROUP BY ps.[object_id], ps.partition_number, p.data_compression_desc, i.[name]
ORDER BY schemaName, tableName, ps.partition_number;
