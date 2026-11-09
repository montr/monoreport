#!/usr/bin/env bash
set -euo pipefail

# Verification tooling only: restore the archived changes into fresh source trees.
archive=$(cd "$(dirname "$0")" && pwd)
repository=$(git -C "$archive" rev-parse --show-toplevel)
if [ "$#" -ne 1 ] || [ -e "$1" ]; then
  echo 'Usage: bash reproduce.sh <new-workspace-directory>' >&2
  exit 2
fi
mkdir -p "$1"
work=$(cd "$1" && pwd)
mkdir -p "$work/tools" "$work/downloads" "$work/logs" "$work/upstream" "$work/monoreport"
export GIT_CEILING_DIRECTORIES="$work"
export CARGO_HOME="$work/tools/cargo"
export RUSTUP_HOME="$work/tools/rustup"
export WASM_PACK_CACHE="$work/tools/wasm-pack-cache"
export TMPDIR="$work/tmp"
export CARGO_BUILD_JOBS=2
export npm_config_cache="$work/tools/npm-cache"
export npm_config_userconfig=/dev/null
mkdir -p "$WASM_PACK_CACHE" "$TMPDIR"

value() {
  python3 - "$archive/manifest.json" "$1" <<'PY'
import json, sys
value = json.load(open(sys.argv[1]))
for key in sys.argv[2].split('.'):
    value = value[key]
print(value)
PY
}
verify_hash() {
  python3 - "$1" "$2" <<'PY'
import hashlib, pathlib, sys
actual = hashlib.sha256(pathlib.Path(sys.argv[1]).read_bytes()).hexdigest()
if actual != sys.argv[2]:
    raise SystemExit('SHA-256 mismatch: ' + sys.argv[1])
PY
}
phase() {
  local name="$1"
  shift
  printf '%s\n' "Running $name"
  if "$@" 2>&1 | tee "$work/logs/$name.log"; then
    printf '%s\t0\n' "$name" >> "$work/results.tsv"
  else
    printf '%s\t1\n' "$name" >> "$work/results.tsv"
    return 1
  fi
}
fetch() { curl --fail --location --retry 2 --silent --show-error "$1" --output "$2"; }

case "$(uname -s)-$(uname -m)" in
  Darwin-arm64) node_target=darwin-arm64; rust_target=aarch64-apple-darwin; pack_target=aarch64-apple-darwin ;;
  Darwin-x86_64) node_target=darwin-x64; rust_target=x86_64-apple-darwin; pack_target=x86_64-apple-darwin ;;
  Linux-x86_64) node_target=linux-x64; rust_target=x86_64-unknown-linux-gnu; pack_target=x86_64-unknown-linux-musl ;;
  *) echo 'Unsupported bootstrap platform; use the manual instructions in reproduction.md.' >&2; exit 2 ;;
esac

verify_hash "$archive/print-area-0.89.8.patch" "$(value upstream.patch_sha256)"
verify_hash "$archive/monoreport-integration.patch" "$(value monoreport.integration_patch_sha256)"
node_version=$(value toolchain.node)
node_archive="node-v$node_version-$node_target.tar.gz"
phase download-node fetch "https://nodejs.org/dist/v$node_version/$node_archive" "$work/downloads/$node_archive"
fetch "https://nodejs.org/dist/v$node_version/SHASUMS256.txt" "$work/downloads/node-SHASUMS256.txt"
node_hash=$(awk -v filename="$node_archive" '$2 == filename { print $1 }' "$work/downloads/node-SHASUMS256.txt")
verify_hash "$work/downloads/$node_archive" "$node_hash"
tar -xzf "$work/downloads/$node_archive" -C "$work/tools"
export PATH="$work/tools/node-v$node_version-$node_target/bin:$work/tools:$CARGO_HOME/bin:$PATH"
phase install-npm npm install --global --prefix "$work/tools/node-v$node_version-$node_target" --ignore-scripts --no-audit --no-fund --registry https://registry.npmjs.org "npm@$(value toolchain.npm)"
hash -r
node --version
npm --version

phase download-rustup fetch "https://static.rust-lang.org/rustup/archive/1.28.2/$rust_target/rustup-init" "$work/tools/rustup-init"
chmod +x "$work/tools/rustup-init"
phase install-rust "$work/tools/rustup-init" -y --no-modify-path --profile minimal --default-toolchain "$(value toolchain.rust)" --target wasm32-unknown-unknown
pack_version=$(value toolchain.wasm_pack)
phase download-wasm-pack fetch "https://github.com/wasm-bindgen/wasm-pack/releases/download/v$pack_version/wasm-pack-v$pack_version-$pack_target.tar.gz" "$work/downloads/wasm-pack.tar.gz"
mkdir "$work/tools/wasm-pack-release"
tar -xzf "$work/downloads/wasm-pack.tar.gz" --strip-components=1 -C "$work/tools/wasm-pack-release"
export PATH="$work/tools/wasm-pack-release:$PATH"
phase install-pnpm npm install --prefix "$work/tools/pnpm" --ignore-scripts --no-package-lock --no-audit --no-fund --registry https://registry.npmjs.org "pnpm@$(value toolchain.pnpm)"
pnpm="$work/tools/pnpm/node_modules/pnpm/bin/pnpm.cjs"

phase download-upstream fetch "$(value upstream.source_archive_url)" "$work/downloads/upstream.tar.gz"
verify_hash "$work/downloads/upstream.tar.gz" "$(value upstream.source_archive_sha256)"
tar -xzf "$work/downloads/upstream.tar.gz" --strip-components=1 -C "$work/upstream"
git -C "$work/upstream" apply --check --whitespace=error-all "$archive/print-area-0.89.8.patch"
git -C "$work/upstream" apply "$archive/print-area-0.89.8.patch"
git -C "$repository" archive "$(value monoreport.base_commit)" | tar -x -C "$work/monoreport"
git -C "$work/monoreport" apply --check --whitespace=error-all "$archive/monoreport-integration.patch"
git -C "$work/monoreport" apply "$archive/monoreport-integration.patch"
python3 - "$archive/manifest.json" "$work" <<'PY'
import hashlib, json, pathlib, sys
manifest = json.load(open(sys.argv[1]))
root = pathlib.Path(sys.argv[2])
for section, directory in [('upstream', 'upstream'), ('monoreport', 'monoreport')]:
    for entry in manifest[section]['files']:
        path = root / directory / entry['path']
        if hashlib.sha256(path.read_bytes()).hexdigest() != entry['sha256']:
            raise SystemExit('Restored source mismatch: ' + str(path))
print('Every archived upstream and MonoReport target file matches its manifest hash.')
PY

cd "$work/upstream"
phase install-upstream node "$pnpm" install --frozen-lockfile --ignore-scripts --filter @silurus/ooxml --filter @silurus/ooxml-xlsx... --store-dir "$work/tools/pnpm-store"
cd "$work/monoreport/web"
phase install-frontend npm ci --no-audit --no-fund --registry https://registry.npmjs.org
phase build-library npm run ooxml:local -- "$work/upstream" build-library
cd "$work/upstream"
phase parser-tests cargo test --locked -p xlsx-parser defined_name_scope_tests -- --nocapture
phase legacy-typecheck cargo check --locked -p legacy-office-parser --features direct-xls
phase library-tests ./node_modules/.bin/vitest run packages/xlsx/src/print-area.test.ts packages/xlsx/src/internal/sheet-surface.test.ts packages/xlsx/src/sheet-viewer.test.ts packages/xlsx/src/selection.test.ts packages/xlsx/src/resize-zoom.test.ts packages/xlsx/src/viewer-render-coalesce.test.ts packages/xlsx/src/viewer-hit-test-axis.test.ts packages/xlsx/src/viewer-destroy.test.ts packages/xlsx/src/viewer-reload-orphan.test.ts packages/xlsx/src/viewer-destroy-during-render.test.ts packages/xlsx/src/viewer-outline-dom-parity.test.ts
phase library-typecheck ./node_modules/.bin/tsc -p packages/xlsx/tsconfig.json --noEmit --pretty false
phase api-symmetry node scripts/check-viewer-api-symmetry.mjs
cd "$work/monoreport/web"
phase frontend-tests npm run ooxml:local -- "$work/upstream" test --run
phase frontend-typecheck npm run ooxml:local -- "$work/upstream" typecheck
phase frontend-build npm run ooxml:local -- "$work/upstream" build
printf '%s\n' "Reproduction complete. Results: $work/results.tsv"
