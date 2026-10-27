#!/usr/bin/env bash
# Run ToggJirazer. Any arguments are forwarded to the program.
# Runs from the ToggJirazer project directory so relative report paths land in
# ToggJirazer/. Settings live in the per-user folder
# (~/Library/Application Support/ToggJirazer on macOS).
#
# Examples:
#   ./run.sh                 # open the app
#   ./run.sh --headless      # non-interactive report run
#   ./run.sh --help
set -euo pipefail

if ! command -v dotnet >/dev/null 2>&1; then
    echo "Error: .NET 10 SDK is required. Install it, then run this script again." >&2
    exit 127
fi

cd "$(dirname "$0")/ToggJirazer"

exec dotnet run --project ToggJirazer.csproj -- "$@"
