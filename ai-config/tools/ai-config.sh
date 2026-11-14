#!/usr/bin/env bash
# AI 接入配置命令行：把 HTTP 接口包一层，省得手写 curl/转义。
#
#   ./ai-config.sh get                     看生产配置（key 脱敏）
#   ./ai-config.sh put '{"active":"anthropic"}'
#   ./ai-config.sh put '{"providers":{"anthropic":{"api_key":"sk-ant-xxx"}}}'
#   ./ai-config.sh test                    用当前 provider 试一次
#   ./ai-config.sh test anthropic          指定 provider 试
#   ./ai-config.sh get --local             看本机 8020（本机容器/调试用）
#
# 默认操作平台（192.168.6.245 上的 ai-config，写的是 worker 真正读的那份配置）；
# 加 --local 则操作本机 8020。
set -euo pipefail

REMOTE_HOST="${AI_CONFIG_REMOTE:-ops@192.168.6.245}"
REMOTE_PORT="${AI_CONFIG_REMOTE_PORT:-2222}"
LOCAL_URL="${AI_CONFIG_LOCAL_URL:-http://127.0.0.1:8020}"

SSH_BIN="${SSH_BIN:-ssh}"
if ! command -v "$SSH_BIN" >/dev/null 2>&1; then
  SSH_BIN="/c/Users/AI_ProgramFiles/OpenClaw/WorkBuddy/resources/vendor/PortableGit/usr/bin/ssh.exe"
fi

LOCAL=0
ARGS=()
for a in "$@"; do
  if [ "$a" = "--local" ]; then LOCAL=1; else ARGS+=("$a"); fi
done

cmd="${ARGS[0]:-get}"

# 远端执行：请求体从 stdin 传，避免 JSON 里引号被 shell 二次解析
remote_curl() {
  local method="$1" path="$2" body="${3:-}"
  if [ -n "$body" ]; then
    printf '%s' "$body" | "$SSH_BIN" -p "$REMOTE_PORT" -o BatchMode=yes -o StrictHostKeyChecking=no "$REMOTE_HOST" \
      "curl -s -X $method --data-binary @- -H 'Content-Type: application/json' http://127.0.0.1:8020$path"
  else
    "$SSH_BIN" -p "$REMOTE_PORT" -o BatchMode=yes -o StrictHostKeyChecking=no "$REMOTE_HOST" \
      "curl -s -X $method http://127.0.0.1:8020$path"
  fi
}

local_curl() {
  local method="$1" path="$2" body="${3:-}"
  if [ -n "$body" ]; then
    curl -s -X "$method" --data-binary "$body" -H 'Content-Type: application/json' "$LOCAL_URL$path"
  else
    curl -s -X "$method" "$LOCAL_URL$path"
  fi
}

call() {  # call METHOD PATH [BODY]
  if [ "$LOCAL" = "1" ]; then local_curl "$@"; else remote_curl "$@"; fi
}

pretty() { python -c "import sys,json;print(json.dumps(json.load(sys.stdin),ensure_ascii=False,indent=2))" 2>/dev/null || cat; }

case "$cmd" in
  get)
    call GET /ai-config | pretty
    ;;
  put)
    body="${ARGS[1]:-}"
    [ -z "$body" ] && { echo "用法：ai-config.sh put '<json>' [--local]" >&2; exit 2; }
    call PUT /ai-config "$body" | pretty
    ;;
  test)
    prov="${ARGS[1]:-}"
    if [ -n "$prov" ]; then body="{\"provider\":\"$prov\"}"; else body='{}'; fi
    call POST /ai-config/test "$body" | pretty
    ;;
  raw)
    call GET "${ARGS[1]:-/health}" | pretty
    ;;
  *)
    echo "未知命令：$cmd（get | put | test | raw）" >&2
    exit 2
    ;;
esac
