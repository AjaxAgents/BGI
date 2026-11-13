#!/usr/bin/env bash
# Tests, bundles the page into one HTML file, and cross-compiles the Windows EXE.
set -euo pipefail
cd "$(dirname "$0")/.."
version="$(tr -d '[:space:]' < VERSION)"
node --test tests/engine.test.cjs
node scripts/build-web.mjs
go vet ./...
GOOS=windows GOARCH=amd64 go vet ./...
GOOS=windows GOARCH=amd64 CGO_ENABLED=0 go build -trimpath \
  -ldflags "-s -w -H=windowsgui -X main.version=${version}" \
  -o dist/BGI-Underbilling.exe ./cmd/bgi-underbilling
echo "dist/BGI-Underbilling.exe  v${version}"
