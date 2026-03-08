## Install / Uninstall (SSMS EnvTabs)

Open PowerShell, then:

Install:
```powershell
& "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" "C:\Users\blake\source\repos\SSMS EnvTabs\SSMS EnvTabs\bin\Debug\SSMS EnvTabs.vsix"
```

Reinstall:
```
.\dev\reinstall.ps1
```

Uninstall:
```powershell
& "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" /uninstall:SSMS_EnvTabs.20d4f774-2a12-403b-a25d-1ce263e878d7
```

---

## File Paths

**User Config:**
```
%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json
```

**Logs:**
- File log: `%LocalAppData%\SSMS EnvTabs\runtime.log`
- ActivityLog: `%AppData%\Microsoft\SQL Server Management Studio\22.0\ActivityLog.xml`

**ColorByRegex Config (SSMS temp):**
```
C:\Users\blake\AppData\Local\Temp\<guid>\ColorByRegexConfig.txt
```
- SSMS creates this after opening first query tab
- Extension writes generated regex lines here