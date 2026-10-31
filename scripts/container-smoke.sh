#!/usr/bin/env bash
set -euo pipefail

image="${1:?usage: container-smoke.sh IMAGE [HOST_PORT]}"
port="${2:-8080}"
name="monoreport-smoke-$RANDOM-$$"
root="http://127.0.0.1:$port"
work="$(mktemp -d)"

cleanup() {
  docker rm -f "$name" >/dev/null 2>&1 || true
  rm -rf "$work"
}
trap cleanup EXIT

configured_user="$(docker image inspect "$image" --format '{{.Config.User}}')"
if [[ -z "$configured_user" || "$configured_user" == "0" || "$configured_user" == "root" ]]; then
  echo "Image must configure a non-root runtime user; found '${configured_user:-empty}'." >&2
  exit 1
fi

docker run --rm -d --name "$name" -p "$port:8080" "$image" >/dev/null
ready=false
for _ in $(seq 1 60); do
  if curl --silent --fail --max-time 2 "$root/openapi/v1.json" >"$work/openapi.json"; then
    ready=true
    break
  fi
  if [[ "$(docker inspect "$name" --format '{{.State.Running}}')" != "true" ]]; then
    docker logs "$name" >&2
    exit 1
  fi
  sleep 1
done
if [[ "$ready" != "true" ]]; then
  docker logs "$name" >&2
  echo "Container did not become ready within 60 seconds." >&2
  exit 1
fi

curl --silent --show-error --fail --max-time 10 -D "$work/root.headers" "$root/" -o "$work/index.html"
grep -qi '^Content-Type: text/html' "$work/root.headers"
asset="$(grep -oE '/assets/[^" ]+\.js' "$work/index.html" | head -n 1)"
test -n "$asset"
curl --silent --show-error --fail --max-time 10 -D "$work/asset.headers" "$root$asset" -o "$work/asset.js"
grep -qi '^Content-Type: text/javascript' "$work/asset.headers"
grep -q '"openapi"' "$work/openapi.json"
curl --silent --show-error --fail --location --max-time 10 -D "$work/docs.headers" "$root/docs/" -o "$work/docs.html"
grep -qi '^Content-Type: text/html' "$work/docs.headers"
grep -q 'MonoReport API' "$work/docs.html"

analyze_status="$(curl --silent --show-error --max-time 20 -o "$work/analyze.json" -w '%{http_code}' -F "template=@Samples/sample.template.docx" "$root/api/documents/analyze")"
test "$analyze_status" = "200"
grep -q '"valid":true' "$work/analyze.json"
grep -q '"fields":' "$work/analyze.json"

for path in /api/not-real /docs/not-real /openapi/not-real /missing.js; do
  status="$(curl --silent --show-error --max-time 10 -D "$work/not-found.headers" -o "$work/not-found.body" -w '%{http_code}' "$root$path")"
  test "$status" = "404"
  ! grep -qi '^Content-Type: text/html' "$work/not-found.headers"
  ! grep -qi '<!doctype html' "$work/not-found.body"
done

docker exec "$name" sh -c 'test "$(id -u)" != 0 && ! command -v node && ! command -v npm && test ! -d /usr/share/dotnet/sdk && test ! -d /app/web'
echo "Container smoke checks passed for $image on $root."
