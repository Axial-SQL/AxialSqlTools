-- Source: https://learn.microsoft.com/en-us/troubleshoot/sql/database-engine/security/transfer-logins-passwords-between-instances
-- Made the output more compact
USE [master]
GO
IF OBJECT_ID('dbo.sp_hexadecimal') IS NOT NULL
    DROP PROCEDURE dbo.sp_hexadecimal
GO
CREATE PROCEDURE dbo.sp_hexadecimal
    @binvalue [varbinary](256)
    ,@hexvalue [nvarchar] (514) OUTPUT
AS
BEGIN
    DECLARE @i [smallint]
    DECLARE @length [smallint]
    DECLARE @hexstring [nchar](16)
    SELECT @hexvalue = N'0x'
    SELECT @i = 1
    SELECT @length = DATALENGTH(@binvalue)
    SELECT @hexstring = N'0123456789ABCDEF'
    WHILE (@i < =  @length)
    BEGIN
        DECLARE @tempint   [smallint]
        DECLARE @firstint  [smallint]
        DECLARE @secondint [smallint]
        SELECT @tempint = CONVERT([smallint], SUBSTRING(@binvalue, @i, 1))
        SELECT @firstint = FLOOR(@tempint / 16)
        SELECT @secondint = @tempint - (@firstint * 16)
        SELECT @hexvalue = @hexvalue
            + SUBSTRING(@hexstring, @firstint  + 1, 1)
            + SUBSTRING(@hexstring, @secondint + 1, 1)
        SELECT @i = @i + 1
    END
END
GO
IF OBJECT_ID('dbo.sp_help_revlogin') IS NOT NULL
    DROP PROCEDURE dbo.sp_help_revlogin
GO
CREATE PROCEDURE dbo.sp_help_revlogin
    @login_name [sysname] = NULL
AS
BEGIN
    DECLARE @name                  [sysname]
    DECLARE @type                  [nvarchar](1)
    DECLARE @hasaccess             [int]
    DECLARE @denylogin             [int]
    DECLARE @is_disabled           [int]
    DECLARE @PWD_varbinary         [varbinary](256)
    DECLARE @PWD_string            [nvarchar](514)
    DECLARE @SID_varbinary         [varbinary](85)
    DECLARE @SID_string            [nvarchar](514)
    DECLARE @tmpstr                [nvarchar](4000)
    DECLARE @is_policy_checked     [nvarchar](3)
    DECLARE @is_expiration_checked [nvarchar](3)
    DECLARE @Prefix                [nvarchar](4000)
    DECLARE @defaultdb             [sysname]
    DECLARE @defaultlanguage       [sysname]
    DECLARE @tmpstrRole            [nvarchar](4000)
    IF @login_name IS NULL
    BEGIN
        DECLARE login_curs CURSOR
        FOR
        SELECT p.[sid],p.[name],p.[type],p.is_disabled,p.default_database_name,l.hasaccess,l.denylogin,default_language_name = ISNULL(p.default_language_name,@@LANGUAGE)
        FROM sys.server_principals p
        LEFT JOIN sys.syslogins l ON l.[name] = p.[name]
        WHERE p.[type] IN ('S' /* SQL_LOGIN */,'G' /* WINDOWS_GROUP */,'U' /* WINDOWS_LOGIN */)
            AND p.[name] <> 'sa'
            AND p.[name] not like '##%'
        ORDER BY p.[name]
    END
    ELSE
        DECLARE login_curs CURSOR
        FOR
        SELECT p.[sid],p.[name],p.[type],p.is_disabled,p.default_database_name,l.hasaccess,l.denylogin,default_language_name = ISNULL(p.default_language_name,@@LANGUAGE)
        FROM sys.server_principals p
        LEFT JOIN sys.syslogins l ON l.[name] = p.[name]
        WHERE p.[type] IN ('S' /* SQL_LOGIN */,'G' /* WINDOWS_GROUP */,'U' /* WINDOWS_LOGIN */)
            AND p.[name] <> 'sa'
            AND p.[name] NOT LIKE '##%'
            AND p.[name] = @login_name
        ORDER BY p.[name]
    OPEN login_curs
    FETCH NEXT FROM login_curs INTO @SID_varbinary,@name,@type,@is_disabled,@defaultdb,@hasaccess,@denylogin,@defaultlanguage
    IF (@@fetch_status = - 1)
    BEGIN
        PRINT '/* No login(s) found for ' + QUOTENAME(@login_name) + N'. */'
        CLOSE login_curs
        DEALLOCATE login_curs
        RETURN - 1
    END
    SET @tmpstr = N'/* sp_help_revlogin script
** Generated ' + CONVERT([nvarchar], GETDATE()) + N' on ' + @@SERVERNAME + N'
*/'
    PRINT @tmpstr
    WHILE (@@fetch_status <> - 1)
    BEGIN
        IF (@@fetch_status <> - 2)
        BEGIN
            PRINT ''
            SET @tmpstr = N'/* Login ' + QUOTENAME(@name) + N' */'
            PRINT @tmpstr
            SET @tmpstr = N'IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE [name] = N''' + @name + N''')
BEGIN'
            PRINT @tmpstr
            IF @type IN ('G','U') -- NT-authenticated Group/User
            BEGIN -- NT authenticated account/group 
                SET @tmpstr = N'    CREATE LOGIN ' + QUOTENAME(@name) + N' FROM WINDOWS WITH DEFAULT_DATABASE = ' + QUOTENAME(@defaultdb) + N', DEFAULT_LANGUAGE = ' + QUOTENAME(@defaultlanguage)
            END
            ELSE
            BEGIN -- SQL Server authentication
                -- obtain password and sid
                SET @PWD_varbinary = CAST(LOGINPROPERTY(@name, 'PasswordHash') AS [varbinary](256))
                EXEC dbo.sp_hexadecimal @PWD_varbinary, @PWD_string OUT
                EXEC dbo.sp_hexadecimal @SID_varbinary, @SID_string OUT
                -- obtain password policy state
                SELECT @is_policy_checked = CASE is_policy_checked WHEN 1 THEN 'ON' WHEN 0 THEN 'OFF' ELSE NULL END
                FROM sys.sql_logins
                WHERE [name] = @name

                SELECT @is_expiration_checked = CASE is_expiration_checked WHEN 1 THEN 'ON' WHEN 0 THEN 'OFF' ELSE NULL END
                FROM sys.sql_logins
                WHERE [name] = @name

                SET @tmpstr = NCHAR(9) + N'CREATE LOGIN ' + QUOTENAME(@name) + N' WITH PASSWORD = ' + @PWD_string + N' HASHED,
        SID = ' + @SID_string + N', DEFAULT_DATABASE = ' + QUOTENAME(@defaultdb) + N', DEFAULT_LANGUAGE = ' + QUOTENAME(@defaultlanguage)

                IF @is_policy_checked IS NOT NULL
                BEGIN
                    SET @tmpstr = @tmpstr + N', CHECK_POLICY = ' + @is_policy_checked
                END

                IF @is_expiration_checked IS NOT NULL
                BEGIN
                    SET @tmpstr = @tmpstr + N', CHECK_EXPIRATION = ' + @is_expiration_checked
                END
                SET @tmpstr = @tmpstr + N';';
            END
            IF (@denylogin = 1)
            BEGIN -- login is denied access
                SET @tmpstr = @tmpstr
                    -- + NCHAR(13) + NCHAR(10) + NCHAR(9) + N''
                    + NCHAR(13) + NCHAR(10) + NCHAR(9) + N'DENY CONNECT SQL TO ' + QUOTENAME(@name)
            END
            ELSE IF (@hasaccess = 0)
            BEGIN -- login exists but does not have access
                SET @tmpstr = @tmpstr
                    -- + NCHAR(13) + NCHAR(10) + NCHAR(9) + N''
                    + NCHAR(13) + NCHAR(10) + NCHAR(9) + N'REVOKE CONNECT SQL TO ' + QUOTENAME(@name)
            END
            IF (@is_disabled = 1)
            BEGIN -- login is disabled
                SET @tmpstr = @tmpstr
                    -- + NCHAR(13) + NCHAR(10) + NCHAR(9) + N''
                    + NCHAR(13) + NCHAR(10) + NCHAR(9) + N'ALTER LOGIN ' + QUOTENAME(@name) + N' DISABLE'
            END
            SET @Prefix =
                N'' -- NCHAR(13) + NCHAR(10) + NCHAR(9) + N''
                -- + NCHAR(13) + NCHAR(10) 
                + NCHAR(9) 
                + N'EXEC [master].dbo.sp_addsrvrolemember @loginame = N'''
            SET @tmpstrRole = N''
            SELECT @tmpstrRole = @tmpstrRole
                + CASE WHEN sysadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''sysadmin''' ELSE '' END
                + CASE WHEN securityadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''securityadmin''' ELSE '' END
                + CASE WHEN serveradmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''serveradmin''' ELSE '' END
                + CASE WHEN setupadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''setupadmin''' ELSE '' END
                + CASE WHEN processadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''processadmin''' ELSE '' END
                + CASE WHEN diskadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''diskadmin''' ELSE '' END
                + CASE WHEN dbcreator = 1 THEN @Prefix + LoginName + N''', @rolename = N''dbcreator''' ELSE '' END
                + CASE WHEN bulkadmin = 1 THEN @Prefix + LoginName + N''', @rolename = N''bulkadmin''' ELSE '' END
            FROM (
                SELECT
                    SUSER_SNAME([sid])AS LoginName
                    ,sysadmin
                    ,securityadmin
                    ,serveradmin
                    ,setupadmin
                    ,processadmin
                    ,diskadmin
                    ,dbcreator
                    ,bulkadmin
                FROM sys.syslogins
                WHERE (    sysadmin <> 0
                        OR securityadmin <> 0
                        OR serveradmin <> 0
                        OR setupadmin <> 0
                        OR processadmin <> 0
                        OR diskadmin <> 0
                        OR dbcreator <> 0
                        OR bulkadmin <> 0
                        )
                    AND [name] = @name
                ) L
            IF @tmpstr <> '' PRINT @tmpstr
            IF @tmpstrRole <> '' PRINT @tmpstrRole
            PRINT 'END'
        END
        FETCH NEXT FROM login_curs INTO @SID_varbinary,@name,@type,@is_disabled,@defaultdb,@hasaccess,@denylogin,@defaultlanguage
    END
    CLOSE login_curs
    DEALLOCATE login_curs
    RETURN 0
END
GO
