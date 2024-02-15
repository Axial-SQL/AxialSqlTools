/*
 Keep in mind that until certain point SQL servers creates 16 VLF for each growth operation.
 To make sure that VLFs are large from the beginning grow it a LOT first time.
*/

DECLARE @INT INT = 2, @Cmd VARCHAR(max);
--target size 100GB
WHILE @INT <= 100
BEGIN

	
	SET @Cmd = 'ALTER DATABASE [db] MODIFY FILE ( NAME = N''db_log'', SIZE = '+ CAST(@INT * 1024 AS VARCHAR(10))+'MB )';
	PRINT @Cmd;
	EXEC(@CMD)
	SET @Int += 1;

END


select 
  li.VLF, 
  mf.growth, 
  mf.* 
from sys.master_files mf 
cross apply (SELECT COUNT(*) AS VLF FROM sys.dm_db_log_info(mf.database_id) li) AS li 
where mf.type = 1 
	and mf.database_id > 4 
order by mf.growth
