#!/usr/bin/env bash
set -euo pipefail

IMAGE="${1:?usage: scripts/smoke-test.sh <image>}"
NAME="outlook-mcp-smoke-$$"
PORT="${SMOKE_PORT:-3999}"

cleanup() {
  docker logs "$NAME" 2>&1 | tail -40 || true
  docker rm -f "$NAME" >/dev/null 2>&1 || true
}
trap cleanup EXIT

docker run -d --name "$NAME" -p "127.0.0.1:${PORT}:3000" \
  -e MS365_MCP_CLIENT_ID=smoke-test-client \
  "$IMAGE" >/dev/null

status=starting
for _ in $(seq 1 60); do
  status=$(docker inspect -f '{{.State.Health.Status}}' "$NAME")
  [ "$status" = healthy ] && break
  [ "$status" = unhealthy ] && break
  sleep 2
done
if [ "$status" != healthy ]; then
  echo "container did not become healthy: $status" >&2
  exit 1
fi

token=$(node -e '
const b = o => Buffer.from(JSON.stringify(o)).toString("base64url");
const now = Math.floor(Date.now() / 1000);
console.log([b({ alg: "RS256" }), b({ iss: "smoke", aud: "smoke", tid: "smoke", oid: "smoke", exp: now + 600 }), "x"].join("."));
')

initialize='{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}'

for attempt in 1 2 3; do
  body=$(curl -fsS "http://127.0.0.1:${PORT}/mcp" \
    -H "Authorization: Bearer ${token}" \
    -H 'Accept: application/json, text/event-stream' \
    -H 'Content-Type: application/json' \
    -d "$initialize")
  if ! grep -q '"serverInfo"' <<<"$body"; then
    echo "initialize attempt ${attempt} failed: ${body}" >&2
    exit 1
  fi
done

echo "smoke test passed: healthy and initialize answered 3 times"
