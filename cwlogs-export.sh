#!/usr/bin/env bash
# CloudWatch Logs を期間指定で Excel に書き出す。
# RHEL 9.8 / bash / Python 3.9 標準ライブラリ。追加の pip パッケージは使わない。
#
# 開始時に aws login --remote の認証を確認する。
# CodeCommit では足りず CloudWatch Logs を読めないときは、
# 警告して終了するか、指定したシェルを source してスイッチバックする。
# source は環境変数をこのプロセスに残すため、Python ではなくこのシェルで行う。

set -euo pipefail
shopt -s extglob

readonly EXIT_ERROR=1
readonly EXIT_LOGIN=2
readonly EXIT_SWITCHBACK=3
readonly EXIT_SWITCHBACK_FAILED=4

readonly SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
readonly EXPORTER="${CWLOGS_EXPORTER:-$SCRIPT_DIR/cwlogs_export.py}"
readonly PYTHON_BIN="${CWLOGS_PYTHON:-python3}"

profile="${AWS_PROFILE:-default}"
profile_explicit=0
region=""
region_explicit=0
start_time=""
end_time=""
output=""
timezone_name="${CWLOGS_TIMEZONE:-Asia/Tokyo}"
switchback_mode="${CWLOGS_SWITCHBACK_MODE:-warn}"
switchback_script="${CWLOGS_SWITCHBACK_SCRIPT:-}"
ignore_case=0
force=0
max_events=0
highlights=()
log_groups=()
prefixes=()
switchback_args=()

usage() {
  cat <<'EOF'
使い方:
  cwlogs-export.sh --start 時刻 --end 時刻 --output ファイル.xlsx [オプション]

指定した時間範囲に CloudWatch Logs へ出力されたログを Excel に書き出します。
ロググループごとにシートを分け、ログストリームはシート内のブロックにします。
フォントは Meiryo UI です。開く Windows に Meiryo UI がある必要があります。

開始時の確認:
  1. aws login --remote で認証済みか
  2. 今の権限で CloudWatch Logs を読めるか
     読めない場合はスイッチバックを案内して終了します。
     --switchback-mode auto なら、指定シェルを source して続行します。

必須:
  --start 時刻            この時刻を含む。日付だけならその日の 0:00:00
  --end 時刻              この時刻を含む。秒だけならその秒の末尾まで
  --output ファイル.xlsx  出力先。既存ファイルは --force がなければ拒否

時刻の例:
  2026-04-08
  "2026-04-08 00:00:00"
  2026-04-08T00:00:00+09:00
  1775574000              epoch 秒
  1775574000000           epoch ミリ秒

オプション:
  --region 名前           未指定時は AWS_REGION、AWS_DEFAULT_REGION、プロファイルの順
  --profile 名前          未指定時は AWS_PROFILE、なければ default
  --timezone 名前         表示タイムゾーン。既定は Asia/Tokyo
  --highlight 文言        本文中のこの文言だけを太字・別色にする。繰り返せる
  --ignore-case           強調文言の大文字小文字を区別しない
  --log-group 名前        このロググループだけを対象にする。繰り返せる
  --log-group-prefix 前方一致
                          未指定かつ --log-group も無いときは、リージョン内の全グループ
  --max-events 件数       0 は無制限（既定）。検証用の上限
  --force                 既存の xlsx を上書きする
  --switchback-mode warn|auto
                          warn（既定）: 権限不足なら警告して終了
                          auto: 警告のあと、スイッチバック用シェルを source して続行
  --switchback-script パス
                          別チームのシェル。source するファイルの場所
                          環境変数 CWLOGS_SWITCHBACK_SCRIPT でも指定できる
  --switchback-arg 値     スイッチバック用シェルに渡す引数。繰り返せる
  -h, --help              このヘルプ

終了コード:
  0  出力した
  1  引数、実行環境、AWS CLI のその他の失敗
  2  aws login --remote が未実施、またはセッションが無効
  3  権限不足。スイッチバックが必要（warn）
  4  auto で source したあとも CloudWatch Logs を読めない

スイッチバック用シェルの契約:
  source されるため、認証情報は export してください。
  成功時に exit するとこのスクリプトも終了します。終えるときは return してください。
  例:
    export AWS_ACCESS_KEY_ID=...
    export AWS_SECRET_ACCESS_KEY=...
    export AWS_SESSION_TOKEN=...

環境変数:
  CWLOGS_SWITCHBACK_MODE, CWLOGS_SWITCHBACK_SCRIPT, CWLOGS_TIMEZONE
  AWS_PROFILE, AWS_REGION, AWS_DEFAULT_REGION, AWS_LOGIN_CACHE_DIRECTORY

例:
  cwlogs-export.sh \
    --start "2026-04-08 00:00:00" \
    --end "2026-04-08 23:59:59" \
    --region ap-northeast-1 \
    --output "$HOME/cwlogs.xlsx" \
    --highlight ERROR \
    --highlight "Exception" \
    --switchback-mode warn \
    --switchback-script /opt/aws-team/switchback.sh

  cwlogs-export.sh \
    --start 2026-04-08 --end 2026-04-08 \
    --output "$HOME/cwlogs.xlsx" \
    --switchback-mode auto \
    --switchback-script /opt/aws-team/switchback.sh \
    --switchback-arg production
EOF
}

die() {
  printf '[エラー] %s\n' "$*" >&2
  exit "$EXIT_ERROR"
}

info() {
  printf '[情報] %s\n' "$*" >&2
}

trim() {
  local value="$1"
  value="${value//$'\r'/}"
  value="${value//$'\n'/}"
  value="${value##+([[:space:]])}"
  value="${value%%+([[:space:]])}"
  printf '%s' "$value"
}

require_value() {
  local option="$1"
  local value="${2-}"
  if [[ -z "$value" ]]; then
    die "${option} の値がありません"
  fi
  printf '%s' "$value"
}

next_value() {
  local option="$1"
  local value="${2-}"
  if [[ -z "$value" || "$value" == --* ]]; then
    die "${option} の値がありません"
  fi
  printf '%s' "$value"
}

login_command() {
  local command=(aws login --remote)
  if [[ "$profile_explicit" -eq 1 || "$profile" != "default" ]]; then
    command+=(--profile "$profile")
  fi
  if [[ -n "$region" ]]; then
    command+=(--region "$region")
  fi
  printf '%q ' "${command[@]}"
  printf '\n'
}

print_login_required() {
  local reason="$1"
  local detail="${2:-}"
  cat >&2 <<EOF
================================================================
[警告] aws login --remote による認証が確認できないため終了します
================================================================
${reason}

このサーバーではブラウザを使えないため、先に次を実行してください。
認証が終わるまで CloudWatch Logs にはアクセスしません。

    $(login_command)

認証が完了すると、プロファイルに login_session が設定され、
~/.aws/login/cache にクロスデバイス（--remote）のキャッシュができます。
EOF
  if [[ -n "$detail" ]]; then
    printf '\n詳細:\n%s\n' "$detail" >&2
  fi
  printf '================================================================\n' >&2
  exit "$EXIT_LOGIN"
}

switchback_source_line() {
  if [[ -z "$switchback_script" ]]; then
    printf '%s\n' "（--switchback-script または CWLOGS_SWITCHBACK_SCRIPT でパスを指定してください）"
    return
  fi
  printf 'source %q' "$switchback_script"
  local arg
  for arg in "${switchback_args[@]}"; do
    printf ' %q' "$arg"
  done
  printf '\n'
}

print_switchback_required() {
  local detail="$1"
  cat >&2 <<EOF
================================================================
[警告] 現在の権限では CloudWatch Logs を操作できないため終了します
================================================================
この処理に CodeCommit の操作は不要です。
logs:DescribeLogGroups と logs:FilterLogEvents が必要です。
今のプリンシパルでは AWS 操作ができないため、操作権限のあるロールへ
スイッチバックしてください。

現在のプリンシパル:
    ${CALLER_ARN:-（不明）}

スイッチバックは、別チームが用意したシェルを source します。
    $(switchback_source_line)

手動で切り替えてから、同じ引数で再実行してください。
終了せずに source して続行する場合:

    cwlogs-export.sh --switchback-mode auto --switchback-script <シェルのパス> ...

専用シェルは source されるため、成功時に exit せず return してください。
EOF
  if [[ -n "$detail" ]]; then
    printf '\n詳細:\n%s\n' "$detail" >&2
  fi
  printf '================================================================\n' >&2
  exit "$EXIT_SWITCHBACK"
}

parse_args() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      -h|--help)
        usage
        exit 0
        ;;
      --start)
        start_time="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --start=*)
        start_time="$(require_value "--start" "${1#*=}")"
        shift
        ;;
      --end)
        end_time="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --end=*)
        end_time="$(require_value "--end" "${1#*=}")"
        shift
        ;;
      --output)
        output="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --output=*)
        output="$(require_value "--output" "${1#*=}")"
        shift
        ;;
      --region)
        region="$(next_value "$1" "${2:-}")"
        region_explicit=1
        shift 2
        ;;
      --region=*)
        region="$(require_value "--region" "${1#*=}")"
        region_explicit=1
        shift
        ;;
      --profile)
        profile="$(next_value "$1" "${2:-}")"
        profile_explicit=1
        shift 2
        ;;
      --profile=*)
        profile="$(require_value "--profile" "${1#*=}")"
        profile_explicit=1
        shift
        ;;
      --timezone)
        timezone_name="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --timezone=*)
        timezone_name="$(require_value "--timezone" "${1#*=}")"
        shift
        ;;
      --highlight)
        highlights+=("$(next_value "$1" "${2:-}")")
        shift 2
        ;;
      --highlight=*)
        highlights+=("$(require_value "--highlight" "${1#*=}")")
        shift
        ;;
      --log-group)
        log_groups+=("$(next_value "$1" "${2:-}")")
        shift 2
        ;;
      --log-group=*)
        log_groups+=("$(require_value "--log-group" "${1#*=}")")
        shift
        ;;
      --log-group-prefix)
        prefixes+=("$(next_value "$1" "${2:-}")")
        shift 2
        ;;
      --log-group-prefix=*)
        prefixes+=("$(require_value "--log-group-prefix" "${1#*=}")")
        shift
        ;;
      --switchback-mode)
        switchback_mode="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --switchback-mode=*)
        switchback_mode="$(require_value "--switchback-mode" "${1#*=}")"
        shift
        ;;
      --switchback-script)
        switchback_script="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --switchback-script=*)
        switchback_script="$(require_value "--switchback-script" "${1#*=}")"
        shift
        ;;
      --switchback-arg)
        switchback_args+=("$(next_value "$1" "${2:-}")")
        shift 2
        ;;
      --switchback-arg=*)
        switchback_args+=("$(require_value "--switchback-arg" "${1#*=}")")
        shift
        ;;
      --max-events)
        max_events="$(next_value "$1" "${2:-}")"
        shift 2
        ;;
      --max-events=*)
        max_events="$(require_value "--max-events" "${1#*=}")"
        shift
        ;;
      --ignore-case)
        ignore_case=1
        shift
        ;;
      --force)
        force=1
        shift
        ;;
      --)
        shift
        break
        ;;
      *)
        die "不明な引数です: $1"
        ;;
    esac
  done
  if [[ $# -gt 0 ]]; then
    die "引数が残っています: $*"
  fi
}

check_tools() {
  if ! command -v aws >/dev/null 2>&1; then
    die "aws コマンドが見つかりません。AWS CLI v2.32.0 以上が必要です。"
  fi
  if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    die "python3 が見つかりません。RHEL 9.8 の Python 3.9 を使います。追加パッケージは不要です。"
  fi
  if ! "$PYTHON_BIN" -c 'import sys; raise SystemExit(0 if sys.version_info[:2] >= (3, 9) else 1)'; then
    die "Python 3.9 以上が必要です。PyPI の追加インストールは行いません。"
  fi
  local version_line version
  version_line="$(aws --version 2>&1 || true)"
  version="$(printf '%s\n' "$version_line" | sed -n 's/.*aws-cli\/\([0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\).*/\1/p')"
  if [[ -z "$version" ]]; then
    die "AWS CLI のバージョンを確認できません: ${version_line}"
  fi
  if ! "$PYTHON_BIN" -c 'import sys
current = [int(part) for part in sys.argv[1].split(".")[:3]]
needed = [2, 32, 0]
current += [0] * (3 - len(current))
raise SystemExit(0 if current >= needed else 1)' "$version"; then
    die "AWS CLI 2.32.0 以上が必要です（aws login のため）。検出したバージョン: ${version}"
  fi
  if [[ ! -f "$EXPORTER" ]]; then
    die "Excel 出力プログラムが見つかりません: $EXPORTER"
  fi
}

reject_overriding_credentials() {
  local found=()
  local name
  for name in \
    AWS_ACCESS_KEY_ID \
    AWS_SECRET_ACCESS_KEY \
    AWS_SESSION_TOKEN \
    AWS_WEB_IDENTITY_TOKEN_FILE \
    AWS_CONTAINER_CREDENTIALS_RELATIVE_URI \
    AWS_CONTAINER_CREDENTIALS_FULL_URI
  do
    if [[ -n "${!name:-}" ]]; then
      found+=("$name")
    fi
  done
  if [[ ${#found[@]} -eq 0 ]]; then
    return 0
  fi
  local joined=""
  name=""
  for name in "${found[@]}"; do
    joined+=" ${name}"
  done
  print_login_required \
    "環境変数${joined} が設定されています。これらは aws login --remote のセッションより優先されるため、ログイン状態を確認できません。環境変数を解除してから aws login --remote を実行してください。" \
    "例: unset${joined}"
}

credential_type() {
  local listing="$1"
  printf '%s\n' "$listing" | awk -F' : ' '
    $1 ~ /access_key/ {
      gsub(/^[[:space:]]+|[[:space:]]+$/, "", $3)
      print $3
      exit
    }
  '
}

classify_failure() {
  local text
  text="$(printf '%s' "$1" | tr '[:upper:]' '[:lower:]')"
  if printf '%s' "$text" | grep -Eq 'expiredtoken|requestexpired|nocredentials|unable to locate credentials|unable to load a existing login session|please reauthenticate|invalidclienttokenid|unrecognizedclient'; then
    printf 'auth'
    return
  fi
  if printf '%s' "$text" | grep -Eq 'accessdenied|not authorized|unauthorizedoperation'; then
    printf 'denied'
    return
  fi
  printf 'other'
}

# ログイン確認は現在のシェルで行う。$(関数) の中で exit すると、
# bash の set -e に依存して親が落ちるだけで、変数は親に残らない。
read_login_session() {
  local error_file session_value status
  error_file="$(mktemp "$WORK/aws-err.XXXXXX")"
  set +e
  session_value="$(aws configure get login_session --profile "$profile" 2>"$error_file")"
  status=$?
  set -e
  LOGIN_DETAIL="$(cat "$error_file" 2>/dev/null || true)"
  if [[ "$status" -ne 0 || -z "$(trim "$session_value")" ]]; then
    LOGIN_SESSION=""
    return 1
  fi
  LOGIN_SESSION="$(trim "$session_value")"
  return 0
}

cache_client_id() {
  local login_session="$1"
  local cache_dir cache_key cache_file client_id error_file status
  if [[ -n "${AWS_LOGIN_CACHE_DIRECTORY:-}" ]]; then
    cache_dir="$("$PYTHON_BIN" -c 'import os, sys; print(os.path.abspath(os.path.expanduser(os.path.expandvars(sys.argv[1]))))' "$AWS_LOGIN_CACHE_DIRECTORY")"
  else
    cache_dir="${HOME}/.aws/login/cache"
  fi
  cache_key="$("$PYTHON_BIN" -c 'import hashlib, sys; print(hashlib.sha256(sys.argv[1].encode("utf-8")).hexdigest())' "$login_session")"
  cache_file="${cache_dir}/${cache_key}.json"
  if [[ ! -f "$cache_file" ]]; then
    CLIENT_ID=""
    CACHE_STATUS="missing"
    LOGIN_DETAIL="想定したキャッシュ: ${cache_file}"
    return 1
  fi
  if [[ ! -r "$cache_file" ]]; then
    die "認証キャッシュを読み取れません: ${cache_file}"
  fi
  error_file="$(mktemp "$WORK/cache-err.XXXXXX")"
  set +e
  client_id="$("$PYTHON_BIN" -c '
import json, sys
path = sys.argv[1]
try:
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
except Exception as exc:
    print("キャッシュの JSON を読めません: %s" % exc, file=sys.stderr)
    raise SystemExit(2)
value = data.get("clientId")
if not isinstance(value, str):
    value = ""
sys.stdout.write(value)
' "$cache_file" 2>"$error_file")"
  status=$?
  set -e
  if [[ "$status" -ne 0 ]]; then
    CLIENT_ID=""
    CACHE_STATUS="broken"
    LOGIN_DETAIL="$(cat "$error_file")"
    return 1
  fi
  CLIENT_ID="$client_id"
  CACHE_STATUS="ok"
  return 0
}

check_remote_login() {
  local listing cred_type error_file identity_json status
  if ! read_login_session; then
    print_login_required \
      "プロファイル ${profile} に aws login のセッションがありません。まだ aws login --remote が実行されていないか、別の種類の認証情報です。" \
      "$LOGIN_DETAIL"
  fi
  if ! cache_client_id "$LOGIN_SESSION"; then
    case "$CACHE_STATUS" in
      missing)
        print_login_required \
          "login_session は設定されていますが、認証キャッシュがありません。セッションが削除されているか、期限切れです。" \
          "$LOGIN_DETAIL"
        ;;
      broken)
        print_login_required \
          "認証キャッシュが壊れています。aws login --remote をやり直してください。" \
          "$LOGIN_DETAIL"
        ;;
      *)
        die "認証キャッシュの確認に失敗しました。"
        ;;
    esac
  fi
  case "$CLIENT_ID" in
    *cross-device*)
      ;;
    *same-device*)
      print_login_required \
        "検出したログインは同一端末の aws login です。このサーバーでは aws login --remote が必要です。"
      ;;
    *)
      print_login_required \
        "キャッシュの clientId が aws login --remote（cross-device）ではありません。再認証してください。" \
        "clientId: ${CLIENT_ID:-（空）}"
      ;;
  esac

  error_file="$(mktemp "$WORK/aws-err.XXXXXX")"
  set +e
  listing="$(aws configure list --profile "$profile" 2>"$error_file")"
  status=$?
  set -e
  if [[ "$status" -ne 0 ]]; then
    print_login_required \
      "aws login のセッションを読み込めません。キャッシュの期限が切れている可能性があります。" \
      "$(cat "$error_file")"
  fi
  cred_type="$(credential_type "$listing")"
  cred_type="$(trim "$cred_type")"
  if [[ "$cred_type" != "login" ]]; then
    print_login_required \
      "有効な認証情報の取得元が aws login ではありません（検出: ${cred_type:-なし}）。インスタンスロールや別の鍵が優先されているため、aws login --remote のセッションが使われていません。"
  fi

  if [[ -z "$region" ]]; then
    if [[ -n "${AWS_REGION:-}" ]]; then
      region="$AWS_REGION"
    elif [[ -n "${AWS_DEFAULT_REGION:-}" ]]; then
      region="$AWS_DEFAULT_REGION"
    else
      set +e
      region="$(aws configure get region --profile "$profile" 2>/dev/null)"
      set -e
      region="$(trim "${region:-}")"
    fi
  fi
  if [[ -z "$region" ]]; then
    die "リージョンが未設定です。--region を指定してください。"
  fi

  error_file="$(mktemp "$WORK/aws-err.XXXXXX")"
  set +e
  identity_json="$(aws sts get-caller-identity --output json --no-cli-pager --profile "$profile" --region "$region" 2>"$error_file")"
  status=$?
  set -e
  if [[ "$status" -ne 0 ]]; then
    print_login_required \
      "aws login --remote のセッションで AWS に接続できません。再認証してください。" \
      "$(cat "$error_file")"
  fi
  CALLER_ARN="$("$PYTHON_BIN" -c 'import json,sys; print(json.load(sys.stdin).get("Arn",""))' <<<"$identity_json")"
  info "aws login --remote を確認しました: ${CALLER_ARN}"
}

probe_logs() {
  local error_file status
  error_file="$(mktemp "$WORK/aws-err.XXXXXX")"
  set +e
  aws logs describe-log-groups --limit 1 --no-paginate --output json --no-cli-pager \
    "${profile_args[@]}" --region "$region" >/dev/null 2>"$error_file"
  status=$?
  set -e
  if [[ "$status" -eq 0 ]]; then
    PROBE_RESULT="ok"
    PROBE_DETAIL=""
    return 0
  fi
  # 失敗理由は親シェルの変数に残す。コマンド置換だとここで代入しても消える。
  PROBE_DETAIL="$(cat "$error_file")"
  PROBE_RESULT="$(classify_failure "$PROBE_DETAIL")"
}

resolve_switchback_script() {
  local raw="$1"
  case "$raw" in
    */*|~*)
      ;;
    *)
      die "スイッチバック用シェルはディレクトリを含むパスで指定してください。ファイル名だけは source の PATH 検索になるため使えません: ${raw}"
      ;;
  esac
  "$PYTHON_BIN" -c 'import os, sys; print(os.path.abspath(os.path.expanduser(sys.argv[1])))' "$raw"
}

validate_switchback_script() {
  local resolved="$1"
  if [[ ! -f "$resolved" ]]; then
    die "スイッチバック用シェルが見つかりません: ${resolved}"
  fi
  if [[ ! -r "$resolved" ]]; then
    die "スイッチバック用シェルを読み取れません: ${resolved}"
  fi
}

adopt_credentials_after_switchback() {
  if [[ -n "${AWS_ACCESS_KEY_ID:-}${AWS_SECRET_ACCESS_KEY:-}${AWS_SESSION_TOKEN:-}" ]]; then
    # 環境変数の一時鍵は --profile より優先される。古いログインプロファイルを固定しない。
    profile_args=()
    forward_profile=""
  else
    if [[ -n "${AWS_PROFILE:-}" ]]; then
      profile="$AWS_PROFILE"
    fi
    profile_args=(--profile "$profile")
    forward_profile="$profile"
  fi
  if [[ "$region_explicit" -eq 0 ]]; then
    if [[ -n "${AWS_REGION:-}" ]]; then
      region="$AWS_REGION"
    elif [[ -n "${AWS_DEFAULT_REGION:-}" ]]; then
      region="$AWS_DEFAULT_REGION"
    fi
  fi
}

run_switchback() {
  local status error_file
  info "現在の権限では CloudWatch Logs を操作できないため、スイッチバックを実行します。"
  info "次を source します: $(switchback_source_line)"
  info "専用シェルが成功時に exit すると、このスクリプトもそこで終わります。"
  set +e
  set +u
  set +o pipefail
  # チーム提供シェルの中身はここには無い。source して環境変数だけを受け取る。
  # shellcheck disable=SC1090
  if [[ ${#switchback_args[@]} -gt 0 ]]; then
    source "$switchback_script" "${switchback_args[@]}"
    status=$?
  else
    source "$switchback_script"
    status=$?
  fi
  set -e
  set -u
  set -o pipefail
  # source 先が extglob を落とすと、このシェルの trim が壊れる。
  shopt -s extglob
  if [[ "$status" -ne 0 ]]; then
    printf '[エラー] スイッチバック用シェルが終了コード %s で失敗しました: %s\n' "$status" "$switchback_script" >&2
    exit "$EXIT_SWITCHBACK_FAILED"
  fi
  adopt_credentials_after_switchback
  error_file="$(mktemp "$WORK/aws-err.XXXXXX")"
  set +e
  identity_json="$(aws sts get-caller-identity --output json --no-cli-pager "${profile_args[@]}" --region "$region" 2>"$error_file")"
  status=$?
  set -e
  if [[ "$status" -ne 0 ]]; then
    printf '[エラー] スイッチバック後の認証情報で AWS に接続できません。\n%s\n' "$(cat "$error_file")" >&2
    exit "$EXIT_SWITCHBACK_FAILED"
  fi
  CALLER_ARN="$("$PYTHON_BIN" -c 'import json,sys; print(json.load(sys.stdin).get("Arn",""))' <<<"$identity_json")"
  info "スイッチバック後のプリンシパル: ${CALLER_ARN}"
}

check_logs_or_switchback() {
  profile_args=(--profile "$profile")
  forward_profile="$profile"
  probe_logs
  case "$PROBE_RESULT" in
    ok)
      info "CloudWatch Logs の参照権限を確認しました。"
      return
      ;;
    auth)
      print_login_required "CloudWatch Logs の確認中にセッションが無効になりました。" "$PROBE_DETAIL"
      ;;
    denied)
      ;;
    *)
      die "CloudWatch Logs の権限確認に失敗しました。${PROBE_DETAIL}"
      ;;
  esac

  case "$switchback_mode" in
    warn)
      print_switchback_required "$PROBE_DETAIL"
      ;;
    auto)
      if [[ -z "$switchback_script" ]]; then
        die "スイッチバックを自動実行するには --switchback-script が必要です。"
      fi
      run_switchback
      probe_logs
      case "$PROBE_RESULT" in
        ok)
          info "スイッチバック後の権限で CloudWatch Logs を参照できます。"
          ;;
        auth)
          print_login_required "スイッチバック後の認証情報が無効です。" "$PROBE_DETAIL"
          ;;
        denied)
          printf '[エラー] スイッチバック後も CloudWatch Logs を操作できません。\n%s\n' "$PROBE_DETAIL" >&2
          exit "$EXIT_SWITCHBACK_FAILED"
          ;;
        *)
          die "スイッチバック後の権限確認に失敗しました。${PROBE_DETAIL}"
          ;;
      esac
      ;;
    *)
      die "--switchback-mode は warn または auto です: ${switchback_mode}"
      ;;
  esac
}

build_python_command() {
  py_args=(
    --output "$output"
    --start "$start_time"
    --end "$end_time"
    --region "$region"
    --timezone "$timezone_name"
  )
  if [[ -n "$forward_profile" ]]; then
    py_args+=(--profile "$forward_profile")
  fi
  if [[ "$ignore_case" -eq 1 ]]; then
    py_args+=(--ignore-case)
  fi
  if [[ "$force" -eq 1 ]]; then
    py_args+=(--force)
  fi
  if [[ "$max_events" != "0" ]]; then
    py_args+=(--max-events "$max_events")
  fi
  local value
  for value in "${highlights[@]}"; do
    py_args+=(--highlight "$value")
  done
  for value in "${log_groups[@]}"; do
    py_args+=(--log-group "$value")
  done
  for value in "${prefixes[@]}"; do
    py_args+=(--log-group-prefix "$value")
  done
}

main() {
  parse_args "$@"
  if [[ -z "$start_time" || -z "$end_time" || -z "$output" ]]; then
    die "--start、--end、--output は必須です。--help を参照してください。"
  fi
  case "$output" in
    *.xlsx) ;;
    *) die "出力は .xlsx にしてください: ${output}" ;;
  esac
  case "$switchback_mode" in
    warn|auto) ;;
    *) die "--switchback-mode は warn または auto です: ${switchback_mode}" ;;
  esac
  if [[ ! "$max_events" =~ ^[0-9]+$ ]]; then
    die "--max-events は 0 以上の整数です: ${max_events}"
  fi
  local phrase
  for phrase in "${highlights[@]}"; do
    if [[ -z "$(trim "$phrase")" ]]; then
      die "--highlight に空の文言は指定できません。"
    fi
  done

  export AWS_PAGER=""
  export AWS_CLI_AUTO_PROMPT=off
  if [[ -z "${LANG:-}" ]]; then
    export LANG=C.UTF-8
  fi
  export PYTHONIOENCODING=UTF-8

  WORK="$(mktemp -d "${TMPDIR:-/tmp}/cwlogs-preflight.XXXXXX")"
  cleanup() {
    rm -rf "$WORK"
  }
  trap cleanup EXIT

  check_tools
  output="$("$PYTHON_BIN" -c 'import os, sys; print(os.path.abspath(os.path.expanduser(sys.argv[1])))' "$output")"

  if [[ -n "$switchback_script" ]]; then
    switchback_script="$(resolve_switchback_script "$switchback_script")"
    validate_switchback_script "$switchback_script"
  elif [[ "$switchback_mode" == "auto" ]]; then
    die "スイッチバックを自動実行するには --switchback-script が必要です。"
  fi

  reject_overriding_credentials
  check_remote_login
  check_logs_or_switchback

  if [[ -e "$output" && "$force" -eq 0 ]]; then
    die "出力先が既にあります。上書きするときは --force を付けてください: ${output}"
  fi

  build_python_command
  cleanup
  trap - EXIT
  export CWLOGS_PREFLIGHT_OK=1
  exec "$PYTHON_BIN" "$EXPORTER" "${py_args[@]}"
}

main "$@"
