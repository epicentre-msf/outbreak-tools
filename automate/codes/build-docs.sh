#!/usr/bin/env bash
set -eu
cd "$(git rev-parse --show-toplevel)"

echo "==> Cleaning previous output..."
rm -rf src/docs

echo "==> Parsing VBA annotations..."
Rscript automate/codes/create-docs.R --all

echo "==> Building HTML site..."
Rscript automate/codes/build-site.R

echo "==> Copying to site/dev/ for Quarto integration..."
rm -rf site/dev
cp -r src/docs/site site/dev

echo "==> Done. Open src/docs/site/index.html"
