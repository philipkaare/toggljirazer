#!/usr/bin/env bash
# Run ToggJirazer. Any arguments are forwarded to the program.
set -euo pipefail

cd "$(dirname "$0")"

dotnet run --project ToggJirazer/ToggJirazer.csproj -- "$@"
