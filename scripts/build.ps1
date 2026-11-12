param(
    [switch]$SkipTests,
    [string]$BuildRoot = (Join-Path ([System.IO.Path]::GetTempPath()) "RenewalReviewBuilderBuild"),
    [string]$NuGetPackagesPath = ""
)

$ErrorActionPreference = "Stop"

function Assert-Success {
    param([string]$Step)

    if ($LASTEXITCODE -ne 0) {
        throw "$Step failed with exit code $LASTEXITCODE"
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedBuildRoot = [System.IO.Path]::GetFullPath($BuildRoot)
if ($resolvedBuildRoot.Length -lt 8) {
    throw "BuildRoot is unexpectedly short: $resolvedBuildRoot"
}

# The staging copy deliberately stays outside the synced workspace.  Publishing
# directly in the repository would generate thousands of bin/obj files there;
# putting the staging copy under this repository would make OneDrive sync them.
$normalizedRepoRoot = $repoRoot.TrimEnd('\', '/')
$normalizedBuildRoot = $resolvedBuildRoot.TrimEnd('\', '/')
if ($normalizedBuildRoot.Equals($normalizedRepoRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
    $normalizedBuildRoot.StartsWith("$normalizedRepoRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "BuildRoot must be outside the OneDrive-synced repository. Use the default local temp folder or a path such as C:\\BuildRenewal-Staging."
}

$localDotnet = Join-Path $repoRoot ".dotnet-sdk\dotnet.exe"
$dotnet = if (Test-Path $localDotnet) { $localDotnet } else { "dotnet" }

$env:DOTNET_CLI_HOME = Join-Path $repoRoot ".dotnet-home"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:NUGET_PACKAGES = if ([string]::IsNullOrWhiteSpace($NuGetPackagesPath)) {
    Join-Path $repoRoot ".nuget\packages"
}
else {
    [System.IO.Path]::GetFullPath($NuGetPackagesPath)
}
$env:APPDATA = Join-Path $repoRoot ".dotnet-home\AppData\Roaming"
$env:LOCALAPPDATA = Join-Path $repoRoot ".dotnet-home\AppData\Local"

try {
    & $dotnet --info | Out-Host
    Assert-Success "dotnet --info"
}
catch {
    Write-Host ""
    Write-Host "The .NET SDK is not installed yet." -ForegroundColor Yellow
    Write-Host "Install .NET 10 SDK from Microsoft, then run this script again:" -ForegroundColor Yellow
    Write-Host "https://dotnet.microsoft.com/download/dotnet/10.0" -ForegroundColor Yellow
    exit 1
}

$distRoot = Join-Path $repoRoot "dist"
$dist = Join-Path $distRoot "RenewalReviewBuilderPortable"
$zipPath = Join-Path $distRoot "RenewalReviewBuilderPortable.zip"

$buildSource = Join-Path $resolvedBuildRoot "source"
$buildDist = Join-Path $resolvedBuildRoot "RenewalReviewBuilderPortable"

if (Test-Path $buildSource) {
    Remove-Item -LiteralPath $buildSource -Recurse -Force
}
if (Test-Path $buildDist) {
    Remove-Item -LiteralPath $buildDist -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $buildSource | Out-Null
New-Item -ItemType Directory -Force -Path $buildDist | Out-Null

robocopy $repoRoot $buildSource /E `
    /XD ".artifacts" ".build" ".build-20260529210944" ".build-20260529211057" ".build-cache" ".codex-backups" ".codex_tmp" ".dotnet" ".dotnet-sdk" ".dotnet-home" ".manual-build" ".nuget" ".offline-packages" ".tmp" ".tmp-render" ".vs" "Backups" "bin" "buildtmp" "dist" "obj" "output" "outputs" "Renewal App - Test" "TestResults" "tmp" `
    /XF ".dotnet-install.ps1" ".dotnet-sdk-*.zip" ".dotnet-sdk-win-x64.tar.gz" "restore-diag.log" | Out-Host
if ($LASTEXITCODE -gt 7) {
    throw "Could not copy source files to the temp build folder. Robocopy exit code: $LASTEXITCODE"
}

# OneDrive Files On-Demand can leave these small support trees as directory-only
# placeholders after robocopy. Copy-Item hydrates their contents into local staging.
foreach ($supportFolderName in @("config", "proposal_engine")) {
    $supportSource = Join-Path $repoRoot $supportFolderName
    $supportDestination = Join-Path $buildSource $supportFolderName

    if (Test-Path -LiteralPath $supportDestination) {
        Remove-Item -LiteralPath $supportDestination -Recurse -Force
    }

    Copy-Item -LiteralPath $supportSource -Destination $supportDestination -Recurse -Force -ErrorAction Stop
}

$solution = Join-Path $buildSource "RenewalReviewBuilder.sln"
$project = Join-Path $buildSource "src\RenewalReviewBuilder.App\RenewalReviewBuilder.App.csproj"
$launcher = Join-Path $buildSource "packaging\Start Renewal Review Builder.cmd"
$readme = Join-Path $buildSource "packaging\README - How to Use.txt"
$offlinePackages = Join-Path $buildSource ".offline-packages"

$restoreTarget = if ($SkipTests) { $project } else { $solution }

if (Test-Path $offlinePackages) {
    & $dotnet restore $restoreTarget --source $offlinePackages
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Offline package cache is missing one or more packages. Falling back to NuGet.org." -ForegroundColor Yellow
        & $dotnet restore $restoreTarget --configfile (Join-Path $buildSource "NuGet.config")
    }
}
else {
    & $dotnet restore $restoreTarget --configfile (Join-Path $buildSource "NuGet.config")
}
Assert-Success "dotnet restore"
if (-not $SkipTests) {
    $testResults = Join-Path $resolvedBuildRoot "TestResults"
    New-Item -ItemType Directory -Force -Path $testResults | Out-Null
    & $dotnet test $solution `
        --configuration Release `
        --no-restore `
        --results-directory $testResults `
        --logger "console;verbosity=normal" `
        --logger "trx;LogFileName=RenewalReviewBuilder.Tests.trx"

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Tests failed, so the portable app was not published yet." -ForegroundColor Yellow
        Write-Host "Test result file:" -ForegroundColor Yellow
        Write-Host (Join-Path $testResults "RenewalReviewBuilder.Tests.trx")
        Write-Host ""
        Write-Host "To make the portable zip now while we inspect the test failure, run:" -ForegroundColor Yellow
        Write-Host "powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1 -SkipTests"
        throw "dotnet test failed with exit code $LASTEXITCODE"
    }
}

New-Item -ItemType Directory -Force -Path $distRoot | Out-Null
if (Test-Path $dist) {
    Remove-Item -LiteralPath $dist -Recurse -Force
}

& $dotnet publish $project `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -p:PublishReadyToRun=false `
    --output $buildDist
Assert-Success "dotnet publish"

Copy-Item -LiteralPath $launcher -Destination (Join-Path $buildDist "Start Renewal Review Builder.cmd") -Force
Copy-Item -LiteralPath $readme -Destination (Join-Path $buildDist "README - How to Use.txt") -Force

Copy-Item -LiteralPath $buildDist -Destination $distRoot -Recurse -Force

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -LiteralPath $dist -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Published portable app:" -ForegroundColor Green
Write-Host (Join-Path $dist "RenewalReviewBuilder.App.exe")
Write-Host ""
Write-Host "Share this zip with the team:" -ForegroundColor Green
Write-Host $zipPath
