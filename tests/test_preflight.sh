#!/usr/bin/env bash
# cwlogs-export.sh の認証とスイッチバックを、偽の aws / python3 で確認する。
set -u
set -o pipefail

ROOT="$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
REAL_PYTHON="/usr/bin/python3"
if [[ ! -x "$REAL_PYTHON" ]]; then
  echo "FAIL /usr/bin/python3 がありません" >&2
  exit 1
fi

WORK="$(mktemp -d)"
HOME_DIR="$WORK/home"
BIN_DIR="$WORK/bin"
MOCK_DIR="$WORK/mock"
ORIG_PATH="$PATH"
mkdir -p "$HOME_DIR" "$BIN_DIR" "$MOCK_DIR"
PASSES=0
FAILS=0

cleanup() {
  rm -rf "$WORK"
}
trap cleanup EXIT

cat > "$BIN_DIR/aws" <<'EOF'
#!/usr/bin/env bash
if [[ "${1:-}" == "--version" ]]; then
  printf 'aws-cli/%s Python/3.9.21 Linux/5.14 source/x86_64\n' "${MOCK_AWS_VERSION:-2.33.1}"
  exit 0
fi
args=("$@")
filtered=()
i=0
while [[ $i -lt ${#args[@]} ]]; do
  token="${args[$i]}"
  case "$token" in
    --output|--profile|--region|--next-token|--limit|--log-group-name|--start-time|--end-time|--log-group-name-prefix)
      i=$((i + 2))
      ;;
    --no-cli-pager|--no-paginate|--interleaved)
      i=$((i + 1))
      ;;
    *)
      filtered+=("$token")
      i=$((i + 1))
      ;;
  esac
done
key="${filtered[*]}"
case "$key" in
  "configure get login_session")
    if [[ "${MOCK_NO_LOGIN:-}" == 1 || -z "${MOCK_LOGIN_SESSION:-}" ]]; then
      echo "could not find login_session" >&2
      exit 1
    fi
    printf '%s\n' "$MOCK_LOGIN_SESSION"
    ;;
  "configure get region")
    printf '%s\n' "${MOCK_REGION:-ap-northeast-1}"
    ;;
  "configure list")
    printf 'access_key : ****************ABCD : %s : login\n' "${MOCK_CRED_TYPE:-login}"
    printf 'secret_key : ****************WXYZ : %s : login\n' "${MOCK_CRED_TYPE:-login}"
    ;;
  "sts get-caller-identity")
    if [[ "${MOCK_STS_FAIL:-}" == 1 ]]; then
      echo "Unable to load a existing login session. Please reauthenticate." >&2
      exit 255
    fi
    printf '%s\n' '{"Account":"123456789012","Arn":"arn:aws:sts::123456789012:assumed-role/Login/user","UserId":"AIDA"}'
    ;;
  "logs describe-log-groups")
    count=0
    if [[ -f "$MOCK_DIR/logs_count" ]]; then
      count="$(cat "$MOCK_DIR/logs_count")"
    fi
    count=$((count + 1))
    printf '%s\n' "$count" > "$MOCK_DIR/logs_count"
    if [[ "${MOCK_LOGS_DENY_ALWAYS:-}" == 1 ]]; then
      echo "An error occurred (AccessDeniedException) when calling the DescribeLogGroups operation: not authorized DETAILTOKEN" >&2
      exit 254
    fi
    if [[ "${MOCK_LOGS_DENY_FIRST:-}" == 1 && "$count" -eq 1 ]]; then
      echo "An error occurred (AccessDeniedException) when calling the DescribeLogGroups operation: not authorized DETAILTOKEN" >&2
      exit 254
    fi
    printf '%s\n' '{"logGroups":[]}'
    ;;
  *)
    printf 'unexpected aws: %s\n' "$key" >&2
    exit 99
    ;;
esac
EOF

cat > "$BIN_DIR/python3" <<EOF
#!/usr/bin/env bash
if [[ "\${1:-}" == "--version" || "\${1:-}" == "-V" || "\${1:-}" == "-c" ]]; then
  exec "$REAL_PYTHON" "\$@"
fi
printf '%s\n' "\$@" > "\$MOCK_DIR/exporter_args"
printf '1\n' > "\$MOCK_DIR/exporter_called"
exit 0
EOF

cat > "$MOCK_DIR/switch.sh" <<'EOF'
#!/usr/bin/env bash
echo sourced >> "$MOCK_DIR/sourced"
shopt -u extglob
export AWS_ACCESS_KEY_ID=AKIATEST
export AWS_SECRET_ACCESS_KEY=secret
export AWS_SESSION_TOKEN=token
EOF

cat > "$MOCK_DIR/switch-fail.sh" <<'EOF'
#!/usr/bin/env bash
echo sourced >> "$MOCK_DIR/sourced"
return 9
EOF

chmod 755 "$BIN_DIR/aws" "$BIN_DIR/python3" "$MOCK_DIR/switch.sh" "$MOCK_DIR/switch-fail.sh"

write_cache() {
  local client_id="$1"
  local session="${2:-remote-session-1}"
  local hash
  hash="$("$REAL_PYTHON" -c 'import hashlib,sys; print(hashlib.sha256(sys.argv[1].encode("utf-8")).hexdigest())' "$session")"
  mkdir -p "$HOME_DIR/.aws/login/cache"
  cat > "$HOME_DIR/.aws/login/cache/${hash}.json" <<EOF
{"clientId":"${client_id}","accessToken":"SUPERSECRETTOKEN"}
EOF
}

clear_cache() {
  rm -rf "$HOME_DIR/.aws"
}

invoke() {
  local envs=()
  local args=()
  local phase="env"
  local item
  for item in "$@"; do
    if [[ "$item" == "--" ]]; then
      phase="args"
      continue
    fi
    if [[ "$phase" == "env" ]]; then
      envs+=("$item")
    else
      args+=("$item")
    fi
  done
  rm -f "$MOCK_DIR/logs_count" "$MOCK_DIR/exporter_called" "$MOCK_DIR/exporter_args" "$MOCK_DIR/sourced"
  set +e
  LAST_OUTPUT="$(
    env \
      -u AWS_ACCESS_KEY_ID \
      -u AWS_SECRET_ACCESS_KEY \
      -u AWS_SESSION_TOKEN \
      -u AWS_WEB_IDENTITY_TOKEN_FILE \
      -u AWS_CONTAINER_CREDENTIALS_RELATIVE_URI \
      -u AWS_CONTAINER_CREDENTIALS_FULL_URI \
      -u AWS_PROFILE \
      -u AWS_REGION \
      -u AWS_DEFAULT_REGION \
      -u CWLOGS_PREFLIGHT_OK \
      -u CWLOGS_SWITCHBACK_MODE \
      -u CWLOGS_SWITCHBACK_SCRIPT \
      HOME="$HOME_DIR" \
      PATH="$BIN_DIR:$ORIG_PATH" \
      MOCK_DIR="$MOCK_DIR" \
      "${envs[@]}" \
      bash "$ROOT/cwlogs-export.sh" "${args[@]}" 2>&1
  )"
  LAST_STATUS=$?
  set -e
}

pass() {
  echo "PASS $1"
  PASSES=$((PASSES + 1))
}

fail() {
  echo "FAIL $1"
  printf '%s\n' "$LAST_OUTPUT"
  FAILS=$((FAILS + 1))
}

expect_status() {
  local name="$1"
  local expected="$2"
  if [[ "$LAST_STATUS" -eq "$expected" ]]; then
    pass "$name"
  else
    fail "$name status=$LAST_STATUS expected=$expected"
  fi
  if [[ "$LAST_OUTPUT" == *SUPERSECRETTOKEN* ]]; then
    fail "$name leaked the login cache secret"
  fi
}

expect_text() {
  local name="$1"
  local needle="$2"
  if [[ "$LAST_OUTPUT" == *"$needle"* ]]; then
    pass "$name"
  else
    fail "$name missing [$needle]"
  fi
}

expect_absent() {
  local name="$1"
  local needle="$2"
  if [[ "$LAST_OUTPUT" == *"$needle"* ]]; then
    fail "$name unexpectedly contained [$needle]"
  else
    pass "$name"
  fi
}

base_args() {
  BASE=(
    --start 2026-04-08
    --end 2026-04-08
    --output "$MOCK_DIR/out.xlsx"
    --region ap-northeast-1
  )
}

# 未ログイン
clear_cache
base_args
invoke MOCK_NO_LOGIN=1 -- "${BASE[@]}"
expect_status "no login session exits 2" 2
expect_text "no login tells the user to run remote login" "aws login --remote"
expect_absent "no login does not call the exporter" "Excel を出力しました"
if [[ -f "$MOCK_DIR/exporter_called" ]]; then
  fail "no login invoked the exporter"
else
  pass "no login did not invoke the exporter"
fi

# 同一端末ログイン
clear_cache
write_cache "arn:aws:signin:::devtools/same-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 -- "${BASE[@]}"
expect_status "same-device login exits 2" 2
expect_text "same-device asks for remote login" "aws login --remote"

# ログイン以外の認証情報
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_CRED_TYPE=iam-role -- "${BASE[@]}"
expect_status "non-login credential type exits 2" 2
expect_text "non-login names the detected type" "iam-role"

# セッションは無効
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_STS_FAIL=1 -- "${BASE[@]}"
expect_status "sts failure exits 2" 2
expect_text "sts failure says the session cannot connect" "接続できません"

# キャッシュ欠落
clear_cache
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 -- "${BASE[@]}"
expect_status "missing cache exits 2" 2
expect_text "missing cache explains the cache" "認証キャッシュ"

# 権限不足は警告して終了し、拒否理由を残す
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_LOGS_DENY_FIRST=1 -- \
  "${BASE[@]}" --switchback-mode warn --switchback-script "$MOCK_DIR/switch.sh"
expect_status "warn mode exits 3" 3
expect_text "warn mode names switchback" "スイッチバック"
expect_text "warn mode keeps the access-denied detail" "DETAILTOKEN"
expect_text "warn mode shows the source path" "$MOCK_DIR/switch.sh"
if [[ -f "$MOCK_DIR/sourced" ]]; then
  fail "warn mode sourced the switchback script"
else
  pass "warn mode did not source the switchback script"
fi

# 自動スイッチバック。最初の拒否のあと source し、再確認して出力へ進む
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_LOGS_DENY_FIRST=1 -- \
  "${BASE[@]}" --switchback-mode auto --switchback-script "$MOCK_DIR/switch.sh" \
  --highlight=ERROR --highlight "disk full"
expect_status "auto mode continues" 0
if [[ ! -f "$MOCK_DIR/sourced" ]]; then
  fail "auto mode did not source the switchback script"
else
  pass "auto mode sourced the switchback script"
fi
if [[ ! -f "$MOCK_DIR/exporter_called" ]]; then
  fail "auto mode did not call the exporter"
else
  pass "auto mode called the exporter"
fi
args_text="$(cat "$MOCK_DIR/exporter_args")"
if [[ "$args_text" == *"--output"* && "$args_text" == *"--start"* && "$args_text" == *"ERROR"* && "$args_text" == *"disk full"* ]]; then
  pass "auto mode forwarded the export arguments"
else
  LAST_OUTPUT="$args_text"
  fail "auto mode export arguments were incomplete"
fi
if [[ "$args_text" == *"--profile"* ]]; then
  LAST_OUTPUT="$args_text"
  fail "auto mode kept --profile after the script exported keys"
else
  pass "auto mode dropped --profile after exported keys"
fi

# 切り替えても拒否される
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_LOGS_DENY_ALWAYS=1 -- \
  "${BASE[@]}" --switchback-mode auto --switchback-script "$MOCK_DIR/switch.sh"
expect_status "auto mode still denied exits 4" 4
expect_text "auto mode reports the second denial" "スイッチバック後も"
if [[ -f "$MOCK_DIR/exporter_called" ]]; then
  fail "denied auto mode called the exporter"
else
  pass "denied auto mode did not call the exporter"
fi

# スイッチバック用シェルが失敗
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 MOCK_LOGS_DENY_FIRST=1 -- \
  "${BASE[@]}" --switchback-mode auto --switchback-script "$MOCK_DIR/switch-fail.sh"
expect_status "failed switchback exits 4" 4
expect_text "failed switchback shows the script status" "終了コード 9"

# 既に権限があるときは source しない。環境変数でも指定できる
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke \
  MOCK_LOGIN_SESSION=remote-session-1 \
  CWLOGS_SWITCHBACK_MODE=auto \
  CWLOGS_SWITCHBACK_SCRIPT="$MOCK_DIR/switch.sh" \
  -- "${BASE[@]}"
expect_status "authorized auto mode exits 0" 0
if [[ -f "$MOCK_DIR/sourced" ]]; then
  fail "authorized auto mode sourced the script"
else
  pass "authorized auto mode did not source the script"
fi
if [[ ! -f "$MOCK_DIR/exporter_called" ]]; then
  fail "authorized auto mode did not call the exporter"
else
  pass "authorized auto mode called the exporter"
fi

# 事前のアクセスキーはログイン確認を妨げる
clear_cache
write_cache "arn:aws:signin:::devtools/cross-device"
base_args
invoke MOCK_LOGIN_SESSION=remote-session-1 AWS_ACCESS_KEY_ID=AKIAPRESET -- "${BASE[@]}"
expect_status "preset access key exits 2" 2
expect_text "preset access key names the variable" "AWS_ACCESS_KEY_ID"

# 古い CLI
clear_cache
base_args
invoke MOCK_AWS_VERSION=2.31.0 -- "${BASE[@]}"
expect_status "old aws cli exits 1" 1
expect_text "old aws cli asks for 2.32" "2.32.0"

# ファイル名だけの source は拒否
clear_cache
base_args
invoke -- "${BASE[@]}" --switchback-script switch.sh
expect_status "bare switchback name exits 1" 1
expect_text "bare switchback name asks for a path" "ディレクトリ"

# 続くトークンがオプションなら値ではない
invoke -- --start 2026-04-08 --end 2026-04-08 --output "$MOCK_DIR/out.xlsx" --highlight --debug
expect_status "highlight without a value exits 1" 1
expect_text "highlight without a value explains it" "値がありません"

echo
echo "passed=$PASSES failed=$FAILS"
if [[ "$FAILS" -ne 0 ]]; then
  exit 1
fi
