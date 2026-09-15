#!/usr/bin/env bash
# Run ToggJirazer. Any arguments are forwarded to the program.
# Runs from the ToggJirazer project directory so the app reads/writes the single
# canonical appsettings.json there and writes report.xlsx alongside it.
set -euo pipefail

cd "$(dirname "$0")/ToggJirazer"

dotnet run --project ToggJirazer.csproj -- "$@"
