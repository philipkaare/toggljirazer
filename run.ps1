#!/usr/bin/env pwsh
# Run ToggJirazer. Any arguments are forwarded to the program.
# Runs from the ToggJirazer project directory so the app reads/writes the single
# canonical appsettings.json there and writes report.xlsx alongside it.
#
# Examples:
#   .\run.ps1                 # Windows GUI
#   .\run.ps1 --tui           # interactive terminal UI
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
