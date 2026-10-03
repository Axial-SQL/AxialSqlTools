[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$VsixPath,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedVersion,

    [Parameter(Mandatory = $true)]
    [string]$VerifierRoot
)

$ErrorActionPreference = 'Stop'

$vsix = Get-Item -LiteralPath $VsixPath
$verifyTool = Get-ChildItem -LiteralPath $VerifierRoot -Filter 'VSIXSignTool.exe' -Recurse -File |
    Select-Object -First 1
if ($null -eq $verifyTool) {
    throw 'VSIXSignTool.exe was not found.'
}

& $verifyTool.FullName verify /v $vsix.FullName
if ($LASTEXITCODE -ne 0) {
    throw "VSIX package signature verification failed with exit code $LASTEXITCODE."
}

$verifyRoot = Join-Path $env:RUNNER_TEMP ('axial-vsix-verify-' + [Guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $verifyRoot -Force | Out-Null
    $zip = Join-Path $verifyRoot 'package.zip'
    $expanded = Join-Path $verifyRoot 'expanded'
    Copy-Item -LiteralPath $vsix.FullName -Destination $zip
    Expand-Archive -LiteralPath $zip -DestinationPath $expanded

    [xml]$manifest = Get-Content -LiteralPath (Join-Path $expanded 'extension.vsixmanifest')
    $identity = $manifest.PackageManifest.Metadata.Identity
    if ($identity.Id -ne 'AxialSqlTools' -or $identity.Version -ne $ExpectedVersion) {
        throw "Signed VSIX identity $($identity.Id) / $($identity.Version) does not match AxialSqlTools / $ExpectedVersion."
    }

    $dlls = @(Get-ChildItem -LiteralPath $expanded -Filter 'AxialSqlTools.dll' -Recurse -File)
    if ($dlls.Count -ne 1) {
        throw 'The signed VSIX must contain exactly one AxialSqlTools.dll.'
    }
    $dll = $dlls[0]
    $signature = Get-AuthenticodeSignature -LiteralPath $dll.FullName
    Write-Host "Embedded assembly: $($dll.FullName)"
    Write-Host "Signature status: $($signature.Status)"
    Write-Host "Signer: $($signature.SignerCertificate.Subject)"
    Write-Host "Thumbprint: $($signature.SignerCertificate.Thumbprint)"
    if ($signature.Status -ne 'Valid') {
        throw "Embedded AxialSqlTools.dll signature is not valid. Status: $($signature.Status)"
    }

    $parsed = [Version]::Parse($ExpectedVersion)
    $expectedAssembly = [Version]::new(
        $parsed.Major, $parsed.Minor,
        [Math]::Max(0, $parsed.Build), [Math]::Max(0, $parsed.Revision))
    $actualAssembly = [Reflection.AssemblyName]::GetAssemblyName($dll.FullName).Version
    if ($actualAssembly -ne $expectedAssembly) {
        throw "Signed assembly version $actualAssembly does not match VSIX version $expectedAssembly."
    }
    Write-Host "Verified signed AxialSqlTools VSIX and assembly version $ExpectedVersion."
}
finally {
    Remove-Item -LiteralPath $verifyRoot -Recurse -Force -ErrorAction SilentlyContinue
}
