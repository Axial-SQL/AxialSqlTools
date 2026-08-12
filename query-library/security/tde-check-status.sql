SELECT
  d.[name] AS DatabaseName,
  encryption_state,
  encryption_state_desc =
  CASE encryption_state
    WHEN '0'  THEN  'No database encryption key present, no encryption'
    WHEN '1'  THEN  'Unencrypted'
    WHEN '2'  THEN  'Encryption in progress'
    WHEN '3'  THEN  'Encrypted'
    WHEN '4'  THEN  'Key change in progress'
    WHEN '5'  THEN  'Decryption in progress'
    WHEN '6'  THEN  'Protection change in progress (The certificate or asymmetric key that'
                    + ' is encrypting the database encryption key is being changed.)'
         ELSE 'No Status' END,
  percent_complete,
  c.[name] AS CertificateName,
  c.[start_date] AS CertificateStartDate,
  c.[expiry_date] AS CertificateExpirationDate,
  encryptor_thumbprint,
  encryptor_type,
  CONCAT('USE [', d.[name], ']
GO
CREATE DATABASE ENCRYPTION KEY WITH ALGORITHM = AES_256 ENCRYPTION BY SERVER CERTIFICATE [TDE];
GO
ALTER DATABASE [', d.[name], '] SET ENCRYPTION ON;
GO
-- rotate the cert only without re-ecnrypting the entire database	
-- ALTER DATABASE ENCRYPTION KEY ENCRYPTION BY SERVER CERTIFICATE [TDE_new];
GO
-- re-encrypt the entire database without changing the certificate
--ALTER DATABASE ENCRYPTION KEY REGENERATE WITH ALGORITHM = AES_256;') AS encryptCommand
FROM sys.databases AS d
LEFT JOIN sys.dm_database_encryption_keys AS dek
    ON dek.database_id = d.database_id
LEFT JOIN master.sys.certificates AS c
    ON c.thumbprint = dek.encryptor_thumbprint
WHERE d.source_database_id IS NULL
  AND d.database_id > 4
ORDER BY d.[name];
