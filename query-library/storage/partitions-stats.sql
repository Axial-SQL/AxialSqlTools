--partitioned table and index details
SELECT
      SCHEMA_NAME(o.schema_id) AS SchemaName,
      OBJECT_NAME(p.object_id) AS ObjectName,
      i.name                   AS IndexName,
      p.index_id               AS IndexID,
      --ds.name                  AS PartitionScheme,   
      --pf.name                  AS PartitionFunction,  
      --CONCAT('ALTER PARTITION FUNCTION ', pf.name, '() MERGE RANGE (''', cast(prv_left.value as datetime), ''')') AS PartitionFunctionMerge
      p.partition_number       AS PartitionNumber,
      --fg.name                  AS FileGroupName,
      prv_left.value           AS LowerBoundaryValue,
      prv_right.value          AS UpperBoundaryValue,
      --CASE pf.boundary_value_on_right WHEN 1 THEN 'RIGHT' ELSE 'LEFT' END AS [Range],
      FORMAT(p.rows, 'N0')	   AS [Rows],
      p.[data_compression_desc] AS [Compression]
FROM sys.partitions                  AS p
JOIN sys.indexes                     AS i ON i.object_id = p.object_id AND i.index_id = p.index_id
JOIN sys.tables				 		 AS o ON i.object_id = o.object_id
JOIN sys.data_spaces                 AS ds ON ds.data_space_id = i.data_space_id
JOIN sys.partition_schemes           AS ps ON ps.data_space_id = ds.data_space_id
JOIN sys.partition_functions         AS pf ON pf.function_id = ps.function_id
JOIN sys.destination_data_spaces     AS dds2 ON dds2.partition_scheme_id = ps.data_space_id AND dds2.destination_id = p.partition_number
JOIN sys.filegroups                  AS fg ON fg.data_space_id = dds2.data_space_id
LEFT JOIN sys.partition_range_values AS prv_left ON ps.function_id = prv_left.function_id AND prv_left.boundary_id = p.partition_number - 1
LEFT JOIN sys.partition_range_values AS prv_right ON ps.function_id = prv_right.function_id  AND prv_right.boundary_id = p.partition_number 
WHERE
      p.index_id = 1 /* only show clustered index */
      --AND OBJECT_NAME(p.object_id) = ''
      --AND SCHEMA_NAME(o.schema_id) = ''
      --AND i.[name] = ''
      --AND p.partition_number = 0
      --AND p.[rows] > 0
ORDER BY
      ObjectName, IndexID, PartitionNumber