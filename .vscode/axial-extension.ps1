param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Build', 'Reinstall')]
    [string]$Action,

    [string]$Configuration = 'Release',

    [string]$Platform = 'Any CPU'
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$projectPath = Join-Path $repoRoot 'AxialSqlTools\AxialSqlTools.csproj'
$solutionPath = Join-Path $repoRoot 'AxialSqlTools\AxialSqlTools.sln'
$vsixPath = Join-Path $repoRoot "AxialSqlTools\bin\$Configuration\AxialSqlTools.vsix"
$vsixInstallerPath = 'C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe'
$extensionId = 'AxialSqlTools'

function Get-ProjectPlatform {
    if ($Platform -eq 'Any CPU') {
        return 'AnyCPU'
    }

    return $Platform
}

function Wait-ForProcessExit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,

        [int]$TimeoutSeconds = 15
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }

    return -not (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
}

function Assert-ReinstallPrerequisites {
    $ssmsProcesses = Get-Process -Name 'Ssms' -ErrorAction SilentlyContinue
    if ($ssmsProcesses) {
        $processList = ($ssmsProcesses | Select-Object -ExpandProperty Id) -join ', '
        throw "SSMS is running (PID: $processList). Close SSMS before reinstalling the extension."
    }

    if (-not (Wait-ForProcessExit -ProcessName 'VSIXInstaller' -TimeoutSeconds 5)) {
        throw 'Another VSIXInstaller process is still running. Wait for it to finish, then rerun the reinstall task.'
    }
}

function Assert-BuildPrerequisites {
    if (-not (Wait-ForProcessExit -ProcessName 'VSIXInstaller' -TimeoutSeconds 10)) {
        throw 'VSIXInstaller is still running and may be locking the output VSIX. Wait for it to finish, then rerun the build task.'
    }

    if (Test-Path $vsixPath) {
        try {
            Remove-Item $vsixPath -Force
        }
        catch {
            throw "The existing VSIX output is locked: $vsixPath. Close anything using it, then rerun the build task."
        }
    }
}

function Get-VsWherePath {
    $programFilesX86 = ${env:ProgramFiles(x86)}
    if (-not $programFilesX86) {
        throw 'ProgramFiles(x86) is not defined.'
    }

    $vswherePath = Join-Path $programFilesX86 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path $vswherePath)) {
        throw "vswhere.exe was not found at $vswherePath"
    }

    return $vswherePath
}

function Get-MSBuildPath {
    $vswherePath = Get-VsWherePath
    $msbuildPath = & $vswherePath -latest -products * -requires Microsoft.VisualStudio.Component.VSSDK -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1

    if (-not $msbuildPath) {
        throw 'Could not find an MSBuild installation with the Visual Studio VSSDK component.'
    }

    if (-not (Test-Path $msbuildPath)) {
        throw "Resolved MSBuild path does not exist: $msbuildPath"
    }

    return $msbuildPath
}

function Invoke-Build {
    if (-not (Test-Path $projectPath)) {
        throw "Project not found: $projectPath"
    }

    $msbuildPath = Get-MSBuildPath
    $projectPlatform = Get-ProjectPlatform

    Assert-BuildPrerequisites

    Write-Host "Building $projectPath"
    Write-Host "Using MSBuild: $msbuildPath"

    & $msbuildPath $projectPath '/t:Restore' "/p:Configuration=$Configuration" "/p:Platform=$projectPlatform"
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild restore failed with exit code $LASTEXITCODE"
    }

    & $msbuildPath $projectPath '/t:Build' "/p:Configuration=$Configuration" "/p:Platform=$projectPlatform" '/m'
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed with exit code $LASTEXITCODE"
    }

    if (-not (Test-Path $vsixPath)) {
        throw "Build completed, but the VSIX was not found at $vsixPath"
    }

    Write-Host "Built VSIX: $vsixPath"
}

function Invoke-Reinstall {
    if (-not (Test-Path $vsixInstallerPath)) {
        throw "VSIXInstaller.exe was not found at $vsixInstallerPath"
    }

    if (-not (Test-Path $vsixPath)) {
        throw "VSIX file not found: $vsixPath. Run the build task first."
    }

    Assert-ReinstallPrerequisites

    Write-Host "Uninstalling extension $extensionId"
    $uninstallProcess = Start-Process -FilePath $vsixInstallerPath -ArgumentList @('/quiet', "/uninstall:$extensionId") -Wait -PassThru
    if ($uninstallProcess.ExitCode -ne 0) {
        Write-Warning "Uninstall exited with code $($uninstallProcess.ExitCode). Continuing with install."
    }

    if (-not (Wait-ForProcessExit -ProcessName 'VSIXInstaller' -TimeoutSeconds 15)) {
        throw 'VSIXInstaller did not exit after uninstall. Wait for the installer to close, then rerun the reinstall task.'
    }

    Write-Host "Installing VSIX $vsixPath"
    $installProcess = Start-Process -FilePath $vsixInstallerPath -ArgumentList @('/quiet', $vsixPath) -Wait -PassThru
    if ($installProcess.ExitCode -ne 0) {
        throw "Install failed with exit code $($installProcess.ExitCode)"
    }

    Write-Host 'Extension reinstall completed.'
}

switch ($Action) {
    'Build' {
        Invoke-Build
    }
    'Reinstall' {
        Invoke-Reinstall
    }
}