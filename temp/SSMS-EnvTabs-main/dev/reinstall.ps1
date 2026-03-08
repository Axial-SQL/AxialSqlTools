$installer = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe"
$vsix = "C:\Users\blake\source\repos\SSMS EnvTabs\SSMS EnvTabs\bin\Release\SSMS EnvTabs.vsix"
$id = "SSMS_EnvTabs.20d4f774-2a12-403b-a25d-1ce263e878d7"
$ssms = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SSMS.lnk"

function Show-Spinner($label, $doneLabel) {
    $spinner = '|/-\'
    $i = 0
    [Console]::CursorVisible = $false
    Write-Host $label
    $line = [Console]::CursorTop - 1
    while (Get-Process VSIXInstaller -ErrorAction SilentlyContinue) {
        $ch = $spinner[$i++ % 4]
        [Console]::SetCursorPosition(0, $line)
        Write-Host -NoNewline ("$label $ch   ")
        [Console]::SetCursorPosition(0, $line + 1)
        Start-Sleep -Milliseconds 250
    }
    [Console]::SetCursorPosition(0, $line)
    Write-Host "$doneLabel   "
    [Console]::SetCursorPosition(0, $line + 1)
    [Console]::CursorVisible = $true
}

Start-Process -FilePath $installer -ArgumentList "/quiet /uninstall:$id"
Show-Spinner "Uninstalling..." "Uninstalling... done!"

Start-Process -FilePath $installer -ArgumentList "/quiet `"$vsix`""
Show-Spinner "Installing..." "Installing... done!"

Start-Process -FilePath $ssms
Read-Host "Press Enter to exit"
