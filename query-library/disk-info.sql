SELECT DISTINCT 
	vs.volume_mount_point,
	vs.file_system_type,
	vs.logical_volume_name,
	CONVERT(DECIMAL (18, 2), vs.total_bytes / 1024.0 / 1024 / 1024) AS [Total Size (GB)],
	CONVERT(DECIMAL (18, 2), vs.available_bytes / 1024.0 / 1024 / 1024) AS [Available Size (GB)],
	CONVERT(DECIMAL (18, 2), vs.available_bytes * 1. / vs.total_bytes * 100.) AS [Space Free %]
FROM sys.master_files AS f WITH (NOLOCK) CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.[file_id]) AS vs
ORDER BY vs.volume_mount_point
OPTION (RECOMPILE);
