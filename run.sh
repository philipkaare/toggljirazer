#!/usr/bin/env bash
# Run ToggJirazer. Any arguments are forwarded to the program.
# Runs from the ToggJirazer project directory so the app reads/writes the single
# canonical appsettings.json there and writes report.xlsx alongside it.
set -euo pipefail

if ! command -v dotnet >/dev/null 2>&1; then
    echo "Error: .NET 10 SDK is required. Install it, then run this script again." >&2
    exit 127
fi

cd "$(dirname "$0")/ToggJirazer"

exec dotnet run --project ToggJirazer.csproj -- "$@"
