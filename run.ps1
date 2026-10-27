#!/usr/bin/env pwsh
# Run ToggJirazer. Any arguments are forwarded to the program.
# Runs from the ToggJirazer project directory so relative report paths land in
# ToggJirazer/. Settings live in the per-user folder (%APPDATA%\ToggJirazer).
#
# Examples:
#   .\run.ps1                 # open the app
#   .\run.ps1 --headless      # non-interactive report run
#   .\run.ps1 --help

$ErrorActionPreference = 'Stop'

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Error '.NET 10 SDK is required. Install it, then run this script again.'
    exit 127
}

$projectDir = Join-Path $PSScriptRoot 'ToggJirazer'
Push-Location $projectDir
try {
    dotnet run --project ToggJirazer.csproj -- @args
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
