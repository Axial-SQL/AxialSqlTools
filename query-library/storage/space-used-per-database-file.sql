IF OBJECT_ID('tempdb..#Stats') IS NOT NULL DROP TABLE #Stats;

SELECT DB_NAME() AS DatabaseName,
	   f.name AS [FileName],
       f.physical_name AS [PhysicalName],
       CAST ((f.size / 128.0) AS DECIMAL (15, 2)) AS [TotalSizeinMB],
       CAST (f.size / 128.0 - CAST (FILEPROPERTY(f.name, 'SpaceUsed') AS INT) / 128.0 AS DECIMAL (15, 2)) AS [AvailableSpaceInMB],
	   CAST(CAST (f.size / 128.0 - CAST (FILEPROPERTY(f.name, 'SpaceUsed') AS INT) / 128.0 AS DECIMAL (15, 2)) * 100 / CAST ((f.size / 128.0) AS DECIMAL (15, 2)) AS DECIMAL (15, 2)) AS [PercentFree],
       [file_id],
       fg.name AS [FilegroupName]
INTO #Stats
FROM sys.database_files AS f WITH (NOLOCK)
     LEFT OUTER JOIN sys.data_spaces AS fg WITH (NOLOCK) ON f.data_space_id = fg.data_space_id
WHERE 1 = 0
OPTION (RECOMPILE);

EXEC sp_msforeachdb N'
USE [?];
INSERT INTO #Stats
SELECT DB_NAME() AS DatabaseName,
	   f.name AS [FileName],
       f.physical_name AS [PhysicalName],
       CAST ((f.size / 128.0) AS DECIMAL (15, 2)) AS [TotalSizeinMB],
       CAST (f.size / 128.0 - CAST (FILEPROPERTY(f.name, ''SpaceUsed'') AS INT) / 128.0 AS DECIMAL (15, 2)) AS [AvailableSpaceInMB],
	   CAST(CAST (f.size / 128.0 - CAST (FILEPROPERTY(f.name, ''SpaceUsed'') AS INT) / 128.0 AS DECIMAL (15, 2)) * 100 / CAST ((f.size / 128.0) AS DECIMAL (15, 2)) AS DECIMAL (15, 2)) AS [PercentFree],
       [file_id],
       fg.name AS [FilegroupName]
FROM sys.database_files AS f WITH (NOLOCK)
     LEFT OUTER JOIN sys.data_spaces AS fg WITH (NOLOCK) ON f.data_space_id = fg.data_space_id
OPTION (RECOMPILE);
';

SELECT 
      'DBCC SHRINKFILE (''' + FileName + ''' , 0)' AS ShrinkCommand
      ,[DatabaseName]
      ,[FileName]
      ,[PhysicalName]
      ,FORMAT([TotalSizeinMB], 'N0') [TotalSizeinMB]
      ,FORMAT([AvailableSpaceInMB], 'N0') [AvailableSpaceInMB]
      ,[PercentFree]
      ,[file_id]
      ,[FilegroupName]
FROM #Stats
ORDER BY 6 DESC;

-- DROP TABLE #Stats;
