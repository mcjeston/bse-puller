param(
    [switch]$SkipPublish
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$publishDir = Join-Path $projectRoot 'bin\\Release\\net8.0-windows\\publish'
$installerDir = Join-Path $projectRoot 'installer'
$distDir = Join-Path $projectRoot 'dist'
$stageDir = Join-Path $distDir 'installer-stage'
$sedPath = Join-Path $distDir 'BsePullerInstaller.sed'
$targetExe = Join-Path $distDir 'BsePullerSetup.exe'

if (-not $SkipPublish) {
    $appData = Join-Path $projectRoot '.appdata'
    $localAppData = Join-Path $projectRoot '.localappdata'
    $nugetDir = Join-Path $appData 'NuGet'
    $nugetPackages = Join-Path $appData 'nuget-packages'

    New-Item -ItemType Directory -Path $appData -Force | Out-Null
    New-Item -ItemType Directory -Path $localAppData -Force | Out-Null
    New-Item -ItemType Directory -Path $nugetDir -Force | Out-Null
    New-Item -ItemType Directory -Path $nugetPackages -Force | Out-Null
    Copy-Item (Join-Path $projectRoot 'NuGet.Config') (Join-Path $nugetDir 'NuGet.Config') -Force

    $env:APPDATA = $appData
    $env:LOCALAPPDATA = $localAppData
    $env:DOTNET_CLI_HOME = $appData
    $env:HOME = $appData
    $env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    $env:NUGET_PACKAGES = $nugetPackages

    & (Join-Path $projectRoot '.dotnet\\dotnet.exe') publish (Join-Path $projectRoot 'BsePuller.csproj') -c Release --no-restore -o $publishDir
    if ($LASTEXITCODE -ne 0) {
        throw 'Publish failed.'
    }
}

if (-not (Test-Path $publishDir)) {
    throw "Publish output not found: $publishDir"
}

if (Test-Path $stageDir) {
    Remove-Item -Path $stageDir -Recurse -Force
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

if (Test-Path $targetExe) {
    Remove-Item -Path $targetExe -Force
}

Copy-Item (Join-Path $publishDir '*') $stageDir -Recurse -Force
Copy-Item (Join-Path $installerDir 'Install-BsePuller.cmd') $stageDir -Force
Copy-Item (Join-Path $installerDir 'Install-BsePuller.ps1') $stageDir -Force

$files = Get-ChildItem -Path $stageDir -File | Sort-Object Name
$stringLines = New-Object System.Collections.Generic.List[string]
$sourceLines = New-Object System.Collections.Generic.List[string]

for ($i = 0; $i -lt $files.Count; $i++) {
    $token = "FILE$i"
    $stringLines.Add("$token=""$($files[$i].Name)""")
    $sourceLines.Add("%$token%=")
}

$sedLines = @(
    '[Version]',
    'Class=IEXPRESS',
    'SEDVersion=3',
    '',
    '[Options]',
    'PackagePurpose=InstallApp',
    'ShowInstallProgramWindow=1',
    'HideExtractAnimation=0',
    'UseLongFileName=1',
    'InsideCompressed=0',
    'CAB_FixedSize=0',
    'CAB_ResvCodeSigning=0',
    'RebootMode=N',
    'InstallPrompt=%InstallPrompt%',
    'DisplayLicense=%DisplayLicense%',
    'FinishMessage=%FinishMessage%',
    'TargetName=%TargetName%',
    'FriendlyName=%FriendlyName%',
    'AppLaunched=%AppLaunched%',
    'PostInstallCmd=%PostInstallCmd%',
    'AdminQuietInstCmd=%AdminQuietInstCmd%',
    'UserQuietInstCmd=%UserQuietInstCmd%',
    'SourceFiles=SourceFiles'
)

$sedLines += @('', '[Strings]',
    'InstallPrompt=',
    'DisplayLicense=',
    'FinishMessage=',
    "TargetName=$targetExe",
    'FriendlyName=BSE Puller Setup',
    'AppLaunched=cmd.exe /c Install-BsePuller.cmd',
    'PostInstallCmd=<None>',
    'AdminQuietInstCmd=',
    'UserQuietInstCmd=')
$sedLines += $stringLines
$sedLines += @('', '[SourceFiles]', "SourceFiles0=$stageDir\", '', '[SourceFiles0]')
$sedLines += $sourceLines

Set-Content -Path $sedPath -Value $sedLines -Encoding ASCII

Push-Location $projectRoot
try {
    cmd /c "iexpress.exe /N dist\BsePullerInstaller.sed"
}
finally {
    Pop-Location
}

if (-not (Test-Path $targetExe)) {
    throw 'IExpress did not produce the expected installer EXE.'
}

Write-Host "Installer created: $targetExe"
