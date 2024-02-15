SELECT DB_NAME() as DatabaseName,
       CONCAT(s.name, '.', t.name) AS TableName,
       c.name AS ColumnName,
       ty.Name + '(' + CAST (c.max_length AS VARCHAR (10)) + ')' AS ColumnType,
       cmk.[name] AS CMK_name,
       cmk.create_date AS CMK_createDate,
	   cmk.[key_path] as CMK_path, 
       k.[name] AS CEK_name,
       c.encryption_type_desc
FROM sys.columns AS c
     INNER JOIN sys.types AS ty
		ON c.user_type_id = ty.user_type_id
     INNER JOIN sys.column_encryption_keys AS k
		ON c.column_encryption_key_id = k.column_encryption_key_id
     INNER JOIN sys.tables AS t
		ON c.object_id = t.object_id
     INNER JOIN sys.schemas AS s
		ON t.schema_id = s.schema_id
     INNER JOIN sys.column_encryption_key_values AS cekv
		ON cekv.column_encryption_key_id = k.column_encryption_key_id
     INNER JOIN sys.column_master_keys AS cmk
		ON cekv.column_master_key_id = cmk.column_master_key_id
WHERE encryption_type IS NOT NULL;
