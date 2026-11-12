$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$localDotnet = Join-Path $repoRoot ".dotnet-sdk\dotnet.exe"
$dotnet = if (Test-Path $localDotnet) { $localDotnet } else { "dotnet" }
$project = Join-Path $repoRoot "src\RenewalReviewBuilder.App\RenewalReviewBuilder.App.csproj"

$env:DOTNET_CLI_HOME = Join-Path $repoRoot ".dotnet-home"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:NUGET_PACKAGES = Join-Path $repoRoot ".nuget\packages"
$env:APPDATA = Join-Path $repoRoot ".dotnet-home\AppData\Roaming"
$env:LOCALAPPDATA = Join-Path $repoRoot ".dotnet-home\AppData\Local"

& $dotnet run --project $project
