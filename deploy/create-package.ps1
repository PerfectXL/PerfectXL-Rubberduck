# Create Chocolatey NuGet packages for a set of Inno Setup installer files.

param (
    [string]$SetupFilePattern = "setup-*.exe"
)

Get-Command NuGet

$ErrorActionPreference = "Stop"

foreach ($SetupFile in Get-Item $SetupFilePattern) {

    if ( -Not (Test-Path $SetupFile)) {
        Throw New-Object System.IO.FileNotFoundException($SetupFile)
        return
    }

    $PackageName = $SetupFile.Name -Replace "^setup-", "" -Replace "-\d.*exe$", ""
    $Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFile).FileVersion

    if ( Test-Path temp ) {
        Remove-Item temp -Recurse
    }
    New-Item temp -type directory
    Set-Location temp

@"
<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://schemas.microsoft.com/packaging/2010/07/nuspec.xsd">
    <metadata>
        <id>$PackageName</id>
        <version>$Version</version>
        <authors>Infotron B.V.</authors>
        <requireLicenseAcceptance>false</requireLicenseAcceptance>
        <description>$PackageName</description>
    </metadata>
    <files>
        <file src="tools\**" target="tools" />
    </files>
</package>
"@ | Out-File "$PackageName.nuspec" -encoding UTF8

    New-Item tools -type directory
    Copy-Item $SetupFile tools
    New-Item "tools\$($SetupFile.Name).ignore" -type file

@"
`$packageName = "$PackageName"
`$fileType = "exe"
`$silentArgs = "/SILENT"
`$scriptPath =  `$(Split-Path `$MyInvocation.MyCommand.Path)
`$fileFullPath = Join-Path `$scriptPath "$($SetupFile.Name)"

try {
  Install-ChocolateyInstallPackage `$packageName `$fileType `$silentArgs `$fileFullPath
} catch {
  Write-ChocolateyFailure `$packageName `$(`$_.Exception.Message)
  throw
}
"@ | Out-File tools\chocolateyInstall.ps1 -encoding ASCII

    & cpack
    $nuPkg = ( Get-Item *.nupkg ).Name
    Move-Item $nuPkg -destination .. -force
    Set-Location ..
    Remove-Item temp -Recurse
}
