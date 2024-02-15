CREATE TABLE #OrphanUsers (UserName nvarchar(128), UserType char(1));
--SQL Users
INSERT INTO #OrphanUsers
SELECT dp.name, dp.type
  FROM sys.database_principals dp
    LEFT OUTER JOIN sys.server_principals sp ON dp.sid = sp.sid
  WHERE dp.type = 'S'
    AND sp.sid IS NULL
    AND dp.authentication_type_desc <> 'NONE';

--Windows Users
DECLARE @sqlcmd nvarchar(4000), @username nvarchar(128);

CREATE TABLE #ADinfo (
    AccountName nvarchar(128),
    AccountType char(8), --user or group
    Privilege char(9), --admin, user, or null. 
    MappedLogin nvarchar(128), --the mapped login name by using the mapped rules    
    PermissionPath nvarchar(128));

DECLARE cur_users CURSOR FAST_FORWARD FOR
SELECT dp.name
  FROM sys.database_principals dp
  WHERE dp.type IN ('G','U');

OPEN cur_users;
FETCH NEXT FROM cur_users INTO @username;

WHILE @@FETCH_STATUS = 0
BEGIN
    TRUNCATE TABLE #ADinfo
    SET @sqlcmd = N'INSERT INTO #ADinfo EXEC xp_logininfo N''' + @username + N''',''all''';
    BEGIN TRY
        EXEC sp_executesql @sqlcmd;
    END TRY
    BEGIN CATCH
    END CATCH
    
    IF NOT EXISTS (SELECT * FROM #ADinfo WHERE MappedLogin IS NOT NULL)
        INSERT INTO #OrphanUsers VALUES (@username, 'W');
    
    FETCH NEXT FROM cur_users INTO @username;
END

CLOSE cur_users;
DEALLOCATE cur_users;

DROP TABLE #ADinfo;
SELECT 'USE [' + DB_NAME() + ']; DROP USER [' + username + '];', * FROM #OrphanUsers;
DROP TABLE #OrphanUsers;
