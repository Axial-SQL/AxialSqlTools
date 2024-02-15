/*
	Adjusted from http://www.sqlservercentral.com/Forums/Topic1560182-1550-1.aspx
*/

DROP TABLE IF EXISTS #AllPermissions;
DROP TABLE IF EXISTS #FilterByUsers;
DROP TABLE IF EXISTS #Users;


CREATE TABLE #AllPermissions (id int IDENTITY(1,1), [codeSQL] varchar(1024));
CREATE TABLE #FilterByUsers (userName varchar(64));
--INSERT INTO #FilterByUsers SELECT name FROM sys.server_principals WHERE name in (' old login to replace ');

SET NOCOUNT ON 
INSERT INTO #AllPermissions ([codeSQL]) VALUES('--=================================================================================');
INSERT INTO #AllPermissions ([codeSQL]) VALUES('USE [' + DB_NAME() + '];');
INSERT INTO #AllPermissions ([codeSQL]) VALUES('--=================================================================================');

--========================================================
--script any certificates in the database
--========================================================
IF (SELECT COUNT(*) FROM sys.certificates) = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No certificates found*/'
	PRINT '/*No certificates found*/';
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) VALUES ('/* Scripting all user certificates */');
	INSERT INTO #AllPermissions ([codeSQL])
	SELECT 'CREATE CERTIFICATE ' + name + 
		' ENCRYPTION BY PASSWORD = ''P@ssw0rd1''
		WITH SUBJECT = ''' + issuer_name + ''', 
		EXPIRY_DATE = ''' + CONVERT(NVARCHAR(25), expiry_date, 120) + '''' + CHAR(13)
	FROM sys.certificates

END

--======================================================
--Script the database users
--======================================================
SELECT principal_id INTO #users FROM sys.database_principals WHERE type IN ('U', 'G', 'S') AND principal_id > 4
	AND (name COLLATE Latin1_General_CI_AS in (SELECT F.userName FROM #FilterByUsers F) OR (SELECT COUNT(*) FROM #FilterByUsers) = 0);

IF (SELECT COUNT(*) FROM #users) = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No database users found*/'
	PRINT '/*No database users found*/';
END 
ELSE
BEGIN
	SET CONCAT_NULL_YIELDS_NULL OFF
	DECLARE @uid INT;

	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/* Scripting all database users and schemas */';

	--SELECT @sql = @sql + 'USE ' + QUOTENAME(DB_NAME(DB_ID())) + CHAR(13) + CHAR(10) + CHAR(13)
	WHILE (SELECT TOP 1 principal_id FROM #users) IS NOT NULL
		BEGIN
			SELECT TOP 1 @uid = principal_id FROM #users	
			
			INSERT INTO #AllPermissions ([codeSQL]) 
			SELECT 'IF (SELECT name FROM sys.database_principals WHERE name = ''' + dp.name + ''') IS NULL' + CHAR(13) + 'BEGIN' + CHAR(13) +
			'	CREATE USER ' + QUOTENAME(dp.name) + 
				/*CASE 
				WHEN SUSER_SID(dp.name) IS NULL THEN ''
				ELSE ' FOR LOGIN ' + QUOTENAME(dp.name)
				END +*/
				CASE
				WHEN SUSER_SNAME(dp.sid) IS NULL THEN ' WITHOUT LOGIN'
				ELSE ' FOR LOGIN ' + QUOTENAME(SUSER_SNAME(dp.sid)) 
				END + 
				CASE 			
				WHEN dp.default_schema_name IS NULL AND dp.type <> 'G' THEN ' WITH DEFAULT_SCHEMA = [dbo]'
				ELSE ' WITH DEFAULT_SCHEMA = [' + dp.default_schema_name + ']'
				END + CHAR(13) + 'END' 
				FROM sys.database_principals dp LEFT OUTER JOIN
				sys.schemas sch ON dp.principal_id = sch.principal_id
				WHERE dp.principal_id = @uid AND dp.TYPE IN ('U', 'G', 'S') AND dp.principal_id > 4			
			
			DELETE FROM #users WHERE principal_id = @uid
			
		END
		
		DROP TABLE #users
END

--========================================================
--Script any users that are protected by a cert
--========================================================
IF (SELECT count(*) FROM sys.database_principals dp INNER JOIN sys.certificates c ON dp.sid = c.sid
	WHERE dp.type = 'C' AND dp.principal_id > 4) = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No certificated users found*/'
	PRINT '/*No certificated users found*/'
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all certificated database users*/';
	INSERT INTO #AllPermissions ([codeSQL]) 
	SELECT 'CREATE USER ' + QUOTENAME(dp.name) + ' FOR CERTIFICATE ' + c.name
	FROM sys.database_principals dp INNER JOIN sys.certificates c ON dp.sid = c.sid
	WHERE dp.type = 'C' AND dp.principal_id > 4
	  		
END
--=======================================================
--script all schemas
--=======================================================
INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all user schema permissions*/';		
INSERT INTO #AllPermissions ([codeSQL]) 
SELECT dp.state_desc COLLATE latin1_general_ci_as + ' ' + 
	dp.permission_name + ' ON ' + dp.class_desc + '::' + QUOTENAME(sch.name) + 
	' TO ' + QUOTENAME(dp2.name) + ' AS ' + QUOTENAME(dp3.name) 
	FROM sys.database_permissions dp 
	INNER JOIN sys.schemas sch ON dp.grantor_principal_id = sch.principal_id
	INNER JOIN sys.database_principals dp2 ON dp.grantee_principal_id = dp2.principal_id
	INNER JOIN sys.database_principals dp3 ON dp.grantor_principal_id = dp3.principal_id
WHERE dp.class = 3  --dp.major_id BETWEEN 1 AND 8

--========================================================
--script database roles from the database
--========================================================
IF (SELECT COUNT(*) FROM sys.database_principals WHERE type = 'R' AND is_fixed_role <> 1 AND principal_id > 4) = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No database roles found*/'
	PRINT '/*No database roles found*/';
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all database roles */';
	INSERT INTO #AllPermissions ([codeSQL]) 
	SELECT 'CREATE ROLE ' + QUOTENAME(dp.name) + ' AUTHORIZATION ' + QUOTENAME(dp2.name)
	FROM sys.database_principals dp INNER JOIN sys.database_principals dp2 
		ON dp.owning_principal_id = dp2.principal_id
	WHERE dp.type = 'R' AND dp.is_fixed_role <> 1 AND dp.principal_id > 4

END

--=========================================================
--script Application roles from the database
--=========================================================
IF (SELECT COUNT(*) FROM sys.database_principals WHERE type = 'A') = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No application roles found*/'
	PRINT '/*No application roles found*/'
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all application roles */'

	INSERT INTO #AllPermissions ([codeSQL]) 
	SELECT 'CREATE APPLICATION ROLE ' + dp.name + ' WITH DEFAULT_SCHEMA = ' + 
	QUOTENAME(dp.default_schema_name) + ', PASSWORD = N''P@ssw0rd1'''
	FROM sys.database_principals dp
	WHERE dp.type = 'A' AND dp.is_fixed_role <> 1 AND dp.principal_id > 4
END

--===============================================================
--got the roles so now we need to get any nested role permissions
--===============================================================
IF (SELECT COUNT(*) from sys.database_principals dp inner join sys.database_role_members drm
		ON dp.principal_id = drm.member_principal_id inner join sys.database_principals dp2 
		ON drm.role_principal_id = dp2.principal_id WHERE dp.type = 'R') = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No nested roles found*/'
	PRINT '/*No nested roles found*/'
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all nested roles */';
	INSERT INTO #AllPermissions ([codeSQL]) SELECT 'EXEC sp_addrolemember ''' + dp2.name + ''', ''' + dp.name + ''''
	FROM sys.database_principals dp 
	INNER JOIN sys.database_role_members drm
	ON dp.principal_id = drm.member_principal_id 
	INNER JOIN sys.database_principals dp2 
	ON drm.role_principal_id = dp2.principal_id
	WHERE dp.type = 'R'
END

--================================================================
--Scripting all user connection grants
--================================================================
IF		(SELECT COUNT(*) FROM sys.database_permissions dpm INNER JOIN sys.database_principals dp 
		ON dpm.grantee_principal_id = dp.principal_id WHERE dp.principal_id > 4 AND dpm.class = 0 AND dpm.type = 'CO') = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No database connection GRANTS found*/'
	PRINT '/*No database connection GRANTS found*/'
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all database and connection GRANTS */'
	INSERT INTO #AllPermissions ([codeSQL]) 
	SELECT dpm.state_desc COLLATE Latin1_General_CI_AS + ' ' + 
		dpm.permission_name COLLATE Latin1_General_CI_AS + ' TO ' + QUOTENAME(dp.name)
	FROM sys.database_permissions dpm INNER JOIN sys.database_principals dp 
		ON dpm.grantee_principal_id = dp.principal_id
	WHERE dp.principal_id > 4 AND dpm.class = 0 --AND dpm.type = 'CO'
		AND (dp.name COLLATE Latin1_General_CI_AS in (SELECT F.userName FROM #FilterByUsers F) OR (SELECT COUNT(*) FROM #FilterByUsers) = 0)

END

--=================================================================
--Now all the object level permissions
--=================================================================
IF		(SELECT	COUNT(*) FROM sys.database_permissions dbpe INNER JOIN sys.database_principals dbpr 
		ON dbpr.principal_id = dbpe.grantee_principal_id INNER JOIN sys.objects obj 
		ON dbpe.major_id = obj.object_id WHERE obj.type NOT IN ('IT','S','X')) = 0
BEGIN
	--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No database user object GRANTS found*/'
	PRINT '/*No database user object GRANTS found*/'
END 
ELSE
BEGIN
	INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all database user object GRANTS */' 

		IF OBJECT_ID('tempdb..#objgrants') IS NOT NULL
		BEGIN
			DROP TABLE #objgrants
		END
		CREATE TABLE #objgrants(
				state_desc	VARCHAR(60)
				, perm_name		NVARCHAR(128)
				, sch_name		NVARCHAR(128)	
				, maj_ID		NVARCHAR(128)
				, name			NVARCHAR(128)	
				, pr_name		NVARCHAR(128)
				, column_name	NVARCHAR(128)
				)
		
		DECLARE @state_desc VARCHAR(60)
		DECLARE @perm_name NVARCHAR(128), @sch_name NVARCHAR(128), @maj_ID NVARCHAR(128)
		DECLARE @name NVARCHAR(128), @pr_name NVARCHAR(128), @column_name NVARCHAR(128)		
					
		INSERT INTO #objgrants
		SELECT CASE dbpe.[state] WHEN 'W' THEN 'GRANT'
		ELSE dbpe.state_desc COLLATE Latin1_General_CI_AS
		END AS [state_desc]
		, dbpe.permission_name COLLATE Latin1_General_CI_AS AS perm_name
		, sch.name AS sch_name
		, OBJECT_NAME(dbpe.major_id) AS maj_ID
		, dbpr.name AS name
		, CASE dbpe.[state] WHEN 'W' THEN '] WITH GRANT OPTION'
		ELSE ']' END AS pr_name
		, ISNULL((select al.name as ColumnName from sys.database_permissions dp
		inner join sys.all_columns al
		on dp.minor_id = al.column_id
		and al.object_id = dbpe.major_id
		where dp.major_id = dbpe.major_id and dp.minor_id > 0 and  dbpe.minor_id = dp.minor_id and dp.class_desc = 'OBJECT_OR_COLUMN'), '') as ColumnName
		FROM sys.database_permissions dbpe INNER JOIN sys.database_principals dbpr 
		ON dbpr.principal_id = dbpe.grantee_principal_id
		INNER JOIN sys.objects obj ON dbpe.major_id = obj.object_id
		INNER JOIN sys.schemas sch ON obj.schema_id = sch.schema_id
		WHERE obj.type NOT IN ('IT','S','X') 
			AND (dbpr.name COLLATE Latin1_General_CI_AS in (SELECT F.userName FROM #FilterByUsers F) OR (SELECT COUNT(*) FROM #FilterByUsers) = 0)
		ORDER BY dbpr.name, obj.name

		WHILE (SELECT COUNT(*) FROM #objgrants) > 0
		BEGIN
			
			SELECT TOP 1 @state_desc = state_desc, @perm_name = perm_name, @sch_name = sch_name, 
			@maj_ID = maj_ID, @name = name, @pr_name = pr_name, @column_name = column_name FROM #objgrants
			
			INSERT INTO #AllPermissions ([codeSQL]) 
			SELECT @state_desc + ' ' + @perm_name +
				' ON [' + @sch_name + '].[' + @maj_ID + ']' + CASE WHEN @column_name = '' THEN '' ELSE CONCAT('(', @column_name, ')') END + ' TO [' + @name + @pr_name
		
			DELETE FROM #objgrants WHERE state_desc = @state_desc AND perm_name = @perm_name 
				AND sch_name = @sch_name AND maj_ID = @maj_ID AND name = @name AND pr_name = @pr_name AND column_name = @column_name
			 
		END
	DROP TABLE #objgrants
	
END
		

--=================================================================
--Now script all the database roles the user have permissions to
--=================================================================
IF		(SELECT COUNT(*) FROM sys.database_principals dp
		INNER JOIN sys.database_role_members drm ON dp.principal_id = drm.role_principal_id
		INNER JOIN sys.database_principals dp2 ON drm.member_principal_id = dp2.principal_id
		WHERE dp2.principal_id > 4 AND dp2.type <> 'R') = 0
BEGIN
		--INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*No database user role GRANTS found*/'
		PRINT '/*No database user role GRANTS found*/'
END 
ELSE
BEGIN
	
		INSERT INTO #AllPermissions ([codeSQL]) SELECT '/*Scripting all database user role permissions */';
		INSERT INTO #AllPermissions ([codeSQL]) 
		SELECT 'EXEC sp_addrolemember ''' + dp.name + ''', ''' + dp2.name + ''''
		FROM sys.database_principals dp
		INNER JOIN sys.database_role_members drm ON dp.principal_id = drm.role_principal_id
		INNER JOIN sys.database_principals dp2 ON drm.member_principal_id = dp2.principal_id
		WHERE dp2.principal_id > 4 AND dp2.type <> 'R'
			AND (dp2.name COLLATE Latin1_General_CI_AS in (SELECT F.userName FROM #FilterByUsers F) OR (SELECT COUNT(*) FROM #FilterByUsers) = 0) 

END

--UPDATE #AllPermissions SET codeSQL = REPLACE(codeSQL, ' old login ', ' new login ');

SELECT codeSQL FROM #AllPermissions ORDER BY id
