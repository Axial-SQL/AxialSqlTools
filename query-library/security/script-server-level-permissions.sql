/* Generate statements to create server permissions for SQL logins, Windows Logins, and Groups */ 
SET NOCOUNT ON; 
 
SELECT  'USE' + SPACE(1) + QUOTENAME('MASTER') AS '--Database Context'; 
 
-- Role Members 
SELECT  'EXEC sp_addsrvrolemember @rolename =' + SPACE(1) + QUOTENAME(usr1.name, '''') + ', @loginame =' + SPACE(1) + QUOTENAME(usr2.name, '''') AS '--Role Memberships'
FROM    sys.server_principals AS usr1
        INNER JOIN sys.server_role_members AS rm ON usr1.principal_id = rm.role_principal_id
        INNER JOIN sys.server_principals AS usr2 ON rm.member_principal_id = usr2.principal_id
ORDER BY rm.role_principal_id ASC; 
 
-- Permissions 
SELECT  server_permissions.state_desc COLLATE SQL_Latin1_General_CP1_CI_AS + ' ' + server_permissions.permission_name COLLATE SQL_Latin1_General_CP1_CI_AS
        + ' TO [' + server_principals.name COLLATE SQL_Latin1_General_CP1_CI_AS + ']' AS '--Server Level Permissions'
FROM    sys.server_permissions AS server_permissions WITH (NOLOCK)
        INNER JOIN sys.server_principals AS server_principals WITH (NOLOCK) ON server_permissions.grantee_principal_id = server_principals.principal_id
WHERE   server_principals.type IN ('S', 'U', 'G')
ORDER BY server_principals.name ,
        server_permissions.state_desc ,
        server_permissions.permission_name;
