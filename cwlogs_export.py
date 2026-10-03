#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CloudWatch Logs を期間指定で取り、cwlogs_xlsx に渡す。

AWS への通信は AWS CLI だけを使う。Python は RHEL 9.8 の標準ライブラリ
（sqlite3, zoneinfo, zipfile）に限る。エントリポイントは cwlogs-export.sh。
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from cwlogs_xlsx import (
    GroupSource,
    StreamSource,
    WorkbookMeta,
    write_xlsx,
)

EXIT_ERROR = 1
EXIT_LOGIN = 2
EXIT_SWITCHBACK = 3


class AwsCommandError(RuntimeError):
    def __init__(self, message: str, returncode: int = 1) -> None:
        super().__init__(message)
        self.returncode = returncode


class AwsAuthExpired(AwsCommandError):
    """ログインセッションが無効。aws login --remote の再実行が必要。"""


class AwsAccessDenied(AwsCommandError):
    """認証は通るが、今の権限では当該 API を呼べない。"""


class AwsNotFound(AwsCommandError):
    """ロググループが存在しない。"""


def classify_aws_failure(stderr: str) -> str:
    """AWS CLI の失敗を auth / missing / throttle / denied / other に分ける。"""
    text = (stderr or "").lower()
    if any(
        token in text
        for token in (
            "expiredtoken",
            "requestexpired",
            "nocredentials",
            "unable to locate credentials",
            "unable to load a existing login session",
            "please reauthenticate",
            "invalidclienttokenid",
            "unrecognizedclientexception",
            "security token included in the request is expired",
            "security token included in the request is invalid",
        )
    ):
        return "auth"
    if "resourcenotfound" in text or "log group does not exist" in text:
        return "missing"
    if any(
        token in text
        for token in (
            "throttl",
            "toomanyrequests",
            "requestlimitexceeded",
            "slowdown",
            "rate exceeded",
        )
    ):
        return "throttle"
    if any(
        token in text
        for token in (
            "accessdenied",
            "not authorized",
            "unauthorizedoperation",
            "is not authorized",
        )
    ):
        return "denied"
    return "other"


def one_line(text: str, limit: int = 280) -> str:
    collapsed = " ".join((text or "").split())
    if len(collapsed) <= limit:
        return collapsed
    return collapsed[: limit - 1] + "…"


@dataclass
class ExportOptions:
    output: Path
    start: str
    end: str
    region: str
    profile: str = ""
    timezone_name: str = "Asia/Tokyo"
    highlights: list[str] = field(default_factory=list)
    log_groups: list[str] = field(default_factory=list)
    prefixes: list[str] = field(default_factory=list)
    ignore_case: bool = False
    force: bool = False
    max_events: int = 0
    aws_executable: str = "aws"
    sleeper: Callable[[float], None] = time.sleep
    max_sheet_rows: int = 1048576


@dataclass
class TimeWindow:
    start_ms: int
    end_ms: int
    start_display: str
    end_display: str
    timezone_name: str


def load_timezone(name: str) -> ZoneInfo:
    try:
        return ZoneInfo(name)
    except ZoneInfoNotFoundError as exc:
        raise SystemExit(
            f"[エラー] タイムゾーン {name} を解釈できません。"
            "RHEL の tzdata に含まれる名前（例: Asia/Tokyo）を指定してください。"
        ) from exc


def _to_millis(moment: datetime) -> int:
    utc = moment.astimezone(timezone.utc)
    epoch = datetime(1970, 1, 1, tzinfo=timezone.utc)
    delta = utc - epoch
    return (delta.days * 86400 + delta.seconds) * 1000 + delta.microseconds // 1000


def _format_local(moment: datetime) -> str:
    return moment.strftime("%Y-%m-%d %H:%M:%S.") + f"{moment.microsecond // 1000:03d}"


def _parse_user_moment(value: str, tz: ZoneInfo) -> tuple[datetime, str]:
    """利用者の時刻文字列を (その精度の開始時刻, 精度名) にする。"""
    text = value.strip()
    if not text:
        raise ValueError("時刻が空です")
    if text.isdigit() and len(text) >= 10:
        raw = int(text)
        if len(text) >= 13:
            millis = raw
            precision = "millisecond"
        else:
            millis = raw * 1000
            precision = "second"
        moment = datetime(1970, 1, 1, tzinfo=timezone.utc) + timedelta(milliseconds=millis)
        return moment.astimezone(tz), precision

    normalized = text.replace("/", "-")
    if normalized.endswith("Z") or normalized.endswith("z"):
        normalized = normalized[:-1] + "+00:00"
    date_only = len(normalized) == 10 and normalized[4] == "-" and normalized[7] == "-"
    if date_only:
        moment = datetime.fromisoformat(normalized).replace(tzinfo=tz)
        return moment, "day"

    body, sep, offset = _split_offset(normalized)
    if "T" not in body and " " in body:
        body = body.replace(" ", "T", 1)
    has_fraction = "." in body
    combined = body + (sep + offset if sep else "")
    moment = datetime.fromisoformat(combined)
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=tz)
    else:
        moment = moment.astimezone(tz)
    if has_fraction:
        millis = moment.microsecond // 1000
        moment = moment.replace(microsecond=millis * 1000)
        return moment, "millisecond"
    clock = body.split("T", 1)[1]
    if clock.count(":") == 1:
        return moment.replace(microsecond=0), "minute"
    return moment.replace(microsecond=0), "second"


def _split_offset(text: str) -> tuple[str, str, str]:
    """末尾の ±HH:MM を本体から切り離す。日付のハイフンは対象にしない。"""
    for index in range(len(text) - 1, 0, -1):
        if text[index] in "+-" and index >= 10:
            return text[:index], text[index], text[index + 1 :]
    return text, "", ""


def _advance(moment: datetime, precision: str) -> datetime:
    if precision == "day":
        return moment + timedelta(days=1)
    if precision == "minute":
        return moment + timedelta(minutes=1)
    if precision == "second":
        return moment + timedelta(seconds=1)
    if precision == "millisecond":
        return moment + timedelta(milliseconds=1)
    raise ValueError("unknown precision: " + precision)


def parse_time_window(start: str, end: str, timezone_name: str) -> TimeWindow:
    tz = load_timezone(timezone_name)
    try:
        start_at, _start_precision = _parse_user_moment(start, tz)
        end_at, end_precision = _parse_user_moment(end, tz)
    except ValueError as exc:
        raise SystemExit(
            "[エラー] 時刻の形式が不正です。"
            "YYYY-MM-DD、YYYY-MM-DD HH:MM:SS、ISO 8601、epoch 秒、epoch ミリ秒を使えます。\n"
            f"詳細: {exc}"
        ) from exc
    # CloudWatch の endTime は「この時刻以降を含めない」。
    # 指定した終了の精度（その日、その秒、そのミリ秒）までは含める。
    end_exclusive = _advance(end_at, end_precision)
    inclusive_end = end_exclusive - timedelta(milliseconds=1)
    start_ms = _to_millis(start_at)
    end_ms = _to_millis(end_exclusive)
    if start_ms >= end_ms:
        raise SystemExit("[エラー] 終了は開始より後にしてください。")
    return TimeWindow(
        start_ms=start_ms,
        end_ms=end_ms,
        start_display=_format_local(start_at),
        end_display=_format_local(inclusive_end),
        timezone_name=timezone_name,
    )


def make_formatter(timezone_name: str) -> Callable[[int], str]:
    tz = load_timezone(timezone_name)

    def format_ts(millis: int) -> str:
        moment = datetime(1970, 1, 1, tzinfo=timezone.utc) + timedelta(milliseconds=int(millis))
        return _format_local(moment.astimezone(tz))

    return format_ts


def run_aws(cmd: Sequence[str], sleeper: Callable[[float], None], attempts: int = 6) -> dict:
    delay = 1.0
    last_error = ""
    returncode = 1
    for attempt in range(1, attempts + 1):
        try:
            completed = subprocess.run(
                list(cmd),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
        except FileNotFoundError as exc:
            raise AwsCommandError(
                "aws コマンドが見つかりません。AWS CLI v2.32.0 以上をインストールしてください。"
            ) from exc
        returncode = completed.returncode
        if returncode == 0:
            payload = (completed.stdout or "").strip()
            if not payload:
                return {}
            try:
                loaded = json.loads(payload)
            except json.JSONDecodeError as exc:
                raise AwsCommandError(
                    "AWS CLI の出力が JSON ではありません: " + payload[:300]
                ) from exc
            if not isinstance(loaded, dict):
                raise AwsCommandError("AWS CLI の JSON がオブジェクトではありません")
            return loaded
        last_error = ((completed.stderr or "") + "\n" + (completed.stdout or "")).strip()
        kind = classify_aws_failure(last_error)
        if kind == "throttle" and attempt < attempts:
            sleeper(delay)
            delay = min(delay * 2, 16.0)
            continue
        message = one_line(last_error) or f"aws が終了コード {returncode} を返しました"
        if kind == "auth":
            raise AwsAuthExpired(message, returncode)
        if kind == "denied":
            raise AwsAccessDenied(message, returncode)
        if kind == "missing":
            raise AwsNotFound(message, returncode)
        raise AwsCommandError(message or last_error, returncode)
    raise AwsCommandError(one_line(last_error) or "AWS CLI が失敗しました", returncode)


def aws_base(options: ExportOptions, service_args: Sequence[str]) -> list[str]:
    command = [options.aws_executable, "--output", "json", "--no-cli-pager"]
    if options.profile:
        command.extend(["--profile", options.profile])
    if options.region:
        command.extend(["--region", options.region])
    command.extend(service_args)
    return command


def paginate(options: ExportOptions, service_args: Sequence[str]) -> Iterable[dict]:
    seen: set[str] = set()
    token: Optional[str] = None
    while True:
        args = list(service_args)
        args.append("--no-paginate")
        if token:
            args.extend(["--next-token", token])
        page = run_aws(aws_base(options, args), options.sleeper)
        yield page
        raw = page.get("nextToken")
        if not isinstance(raw, str) or not raw or raw in seen:
            return
        seen.add(raw)
        token = raw


def log(message: str) -> None:
    print("[情報] " + message, file=sys.stderr, flush=True)


def warn(message: str) -> None:
    print("[警告] " + message, file=sys.stderr, flush=True)


def list_log_groups(options: ExportOptions, prefix: Optional[str]) -> list[dict]:
    args = ["logs", "describe-log-groups", "--limit", "50"]
    if prefix:
        args.extend(["--log-group-name-prefix", prefix])
    found: list[dict] = []
    for page in paginate(options, args):
        groups = page.get("logGroups") or []
        if isinstance(groups, list):
            found.extend(item for item in groups if isinstance(item, dict))
    return found


def resolve_group_names(
    options: ExportOptions,
) -> tuple[list[str], dict[str, Optional[int]]]:
    retention: dict[str, Optional[int]] = {}
    ordered: list[str] = []
    seen: set[str] = set()
    for name in options.log_groups:
        if name not in seen:
            ordered.append(name)
            seen.add(name)
    discovered: list[str] = []
    if options.prefixes or not options.log_groups:
        prefixes: Sequence[Optional[str]] = options.prefixes or [None]
        for prefix in prefixes:
            for group in list_log_groups(options, prefix):
                name = group.get("logGroupName")
                if not isinstance(name, str) or not name:
                    continue
                days = group.get("retentionInDays")
                retention[name] = days if isinstance(days, int) else None
                if name not in seen:
                    discovered.append(name)
                    seen.add(name)
    discovered.sort()
    ordered.extend(discovered)
    for name in options.log_groups:
        if name in retention:
            continue
        try:
            for group in list_log_groups(options, name):
                candidate = group.get("logGroupName")
                if candidate == name:
                    days = group.get("retentionInDays")
                    retention[name] = days if isinstance(days, int) else None
                    break
        except AwsAuthExpired:
            raise
        except (AwsAccessDenied, AwsNotFound, AwsCommandError):
            retention.setdefault(name, None)
    return ordered, retention


def fetch_group_events(
    options: ExportOptions,
    connection: sqlite3.Connection,
    group_name: str,
    window: TimeWindow,
    remaining: Optional[int],
) -> tuple[int, bool]:
    """期間内イベントを SQLite に入れる。戻り値は (件数, 上限で打ち切ったか)。"""
    args = [
        "logs",
        "filter-log-events",
        "--log-group-name",
        group_name,
        "--start-time",
        str(window.start_ms),
        "--end-time",
        str(window.end_ms),
        "--interleaved",
        "--limit",
        "10000",
    ]
    stored = 0
    truncated = False
    for page in paginate(options, args):
        events = page.get("events") or []
        batch = []
        if not isinstance(events, list):
            continue
        for event in events:
            if not isinstance(event, dict):
                continue
            timestamp = event.get("timestamp")
            if not isinstance(timestamp, int):
                continue
            if remaining is not None and stored >= remaining:
                truncated = True
                break
            message = event.get("message")
            stream_name = event.get("logStreamName")
            event_id = event.get("eventId")
            batch.append(
                (
                    group_name,
                    stream_name if isinstance(stream_name, str) and stream_name else "(unknown)",
                    timestamp,
                    message if isinstance(message, str) else "",
                    event_id if isinstance(event_id, str) else "",
                )
            )
            stored += 1
        if batch:
            connection.executemany(
                "INSERT INTO events (group_name, stream_name, ts, message, event_id) "
                "VALUES (?, ?, ?, ?, ?)",
                batch,
            )
            connection.commit()
        if truncated:
            break
    return stored, truncated


def _stream_factory(
    connection: sqlite3.Connection, group_name: str
) -> Callable[[], Iterable[StreamSource]]:
    def iter_streams() -> Iterable[StreamSource]:
        rows = connection.execute(
            """
            SELECT stream_name, COUNT(*), MIN(ts), MAX(ts)
              FROM events
             WHERE group_name = ?
             GROUP BY stream_name
             ORDER BY MIN(ts), stream_name
            """,
            (group_name,),
        ).fetchall()
        for stream_name, count, first_ms, last_ms in rows:
            def iter_events(bound_name: str = stream_name) -> Iterable[tuple]:
                cursor = connection.execute(
                    """
                    SELECT ts, message, event_id
                      FROM events
                     WHERE group_name = ? AND stream_name = ?
                     ORDER BY ts, id
                    """,
                    (group_name, bound_name),
                )
                yield from cursor

            yield StreamSource(
                name=str(stream_name),
                event_count=int(count),
                first_ms=int(first_ms) if first_ms is not None else None,
                last_ms=int(last_ms) if last_ms is not None else None,
                iter_events=iter_events,
            )

    return iter_streams


def _bounds(connection: sqlite3.Connection, group_name: str) -> tuple[Optional[int], Optional[int]]:
    row = connection.execute(
        "SELECT MIN(ts), MAX(ts) FROM events WHERE group_name = ?",
        (group_name,),
    ).fetchone()
    if row is None or row[0] is None:
        return None, None
    return int(row[0]), int(row[1])


def _stream_count(connection: sqlite3.Connection, group_name: str) -> int:
    row = connection.execute(
        "SELECT COUNT(DISTINCT stream_name) FROM events WHERE group_name = ?",
        (group_name,),
    ).fetchone()
    return int(row[0] if row else 0)


def export_logs(options: ExportOptions) -> int:
    if options.max_events < 0:
        raise SystemExit("[エラー] --max-events は 0 以上にしてください。0 は無制限です。")
    phrases = []
    for phrase in options.highlights:
        if phrase.strip() == "":
            raise SystemExit("[エラー] --highlight に空の文言は指定できません。")
        if len(phrase) > 2000:
            raise SystemExit("[エラー] --highlight の文言が長すぎます。")
        phrases.append(phrase)
    output = Path(options.output)
    if output.exists() and not options.force:
        raise SystemExit(
            f"[エラー] 出力先が既にあります。上書きするときは --force を付けてください: {output}"
        )
    if output.exists() and output.is_dir():
        raise SystemExit(f"[エラー] 出力先がディレクトリです: {output}")

    os.environ.setdefault("AWS_PAGER", "")
    os.environ.setdefault("AWS_CLI_AUTO_PROMPT", "off")
    window = parse_time_window(options.start, options.end, options.timezone_name)
    format_ts = make_formatter(options.timezone_name)

    try:
        identity = run_aws(
            aws_base(options, ["sts", "get-caller-identity"]),
            options.sleeper,
        )
    except AwsAuthExpired as exc:
        _print_login_required(str(exc))
        return EXIT_LOGIN
    except AwsCommandError as exc:
        print(f"[エラー] 呼び出し元の確認に失敗しました: {exc}", file=sys.stderr)
        return EXIT_ERROR

    account_id = str(identity.get("Account") or "")
    caller_arn = str(identity.get("Arn") or "")
    log(f"プリンシパル {caller_arn or '（不明）'}")
    log(
        f"期間 {window.start_display} 〜 {window.end_display} ({options.timezone_name})"
        f"  リージョン {options.region}"
    )

    explicit = set(options.log_groups)
    try:
        names, retention = resolve_group_names(options)
    except AwsAuthExpired as exc:
        _print_login_required(str(exc))
        return EXIT_LOGIN
    except AwsAccessDenied as exc:
        _print_switchback_required(str(exc))
        return EXIT_SWITCHBACK
    except AwsCommandError as exc:
        print(f"[エラー] ロググループの一覧取得に失敗しました: {exc}", file=sys.stderr)
        return EXIT_ERROR

    log(f"対象ロググループ {len(names)}")
    groups: list[GroupSource] = []
    empty_groups = 0
    denied_groups = 0
    scanned = 0
    stored_total = 0
    truncated = False
    notice = ""

    with tempfile.TemporaryDirectory(prefix="cwlogs-export-") as temporary:
        connection = sqlite3.connect(Path(temporary) / "events.sqlite")
        try:
            connection.execute("PRAGMA journal_mode=MEMORY")
            connection.execute("PRAGMA synchronous=OFF")
            connection.execute(
                """
                CREATE TABLE events (
                    id INTEGER PRIMARY KEY,
                    group_name TEXT NOT NULL,
                    stream_name TEXT NOT NULL,
                    ts INTEGER NOT NULL,
                    message TEXT NOT NULL,
                    event_id TEXT NOT NULL
                )
                """
            )
            for index, name in enumerate(names, start=1):
                if options.max_events and stored_total >= options.max_events:
                    truncated = True
                    notice = (
                        f"--max-events {options.max_events} に達したため、"
                        "残りのロググループは取得していません。"
                    )
                    warn(notice)
                    break
                scanned += 1
                remaining = None
                if options.max_events:
                    remaining = options.max_events - stored_total
                try:
                    count, group_truncated = fetch_group_events(
                        options, connection, name, window, remaining
                    )
                except AwsAuthExpired as exc:
                    _print_login_required(str(exc))
                    return EXIT_LOGIN
                except AwsNotFound:
                    groups.append(
                        _noted_group(
                            name,
                            retention.get(name),
                            "ロググループが見つかりません。",
                        )
                    )
                    warn(f"{index}/{len(names)}  {name}  見つかりません")
                    continue
                except AwsAccessDenied as exc:
                    denied_groups += 1
                    groups.append(
                        _noted_group(
                            name,
                            retention.get(name),
                            "このロググループを読む権限がありません。 " + one_line(str(exc), 180),
                        )
                    )
                    warn(f"{index}/{len(names)}  {name}  権限がありません")
                    continue
                except AwsCommandError as exc:
                    groups.append(
                        _noted_group(
                            name,
                            retention.get(name),
                            "取得に失敗しました。 " + one_line(str(exc), 180),
                        )
                    )
                    warn(f"{index}/{len(names)}  {name}  取得失敗")
                    continue

                stored_total += count
                if group_truncated:
                    truncated = True
                    notice = (
                        f"--max-events {options.max_events} に達したため、"
                        f"{name} は途中までです。以降のロググループは取得していません。"
                    )
                if count == 0:
                    empty_groups += 1
                    log(f"{index}/{len(names)}  {name}  イベント 0")
                    if name in explicit:
                        groups.append(
                            _noted_group(
                                name,
                                retention.get(name),
                                "この期間に出力されたイベントはありません。",
                            )
                        )
                    continue
                first_ms, last_ms = _bounds(connection, name)
                note = ""
                if group_truncated:
                    note = "指定の --max-events に達したため、このロググループは途中までです。"
                groups.append(
                    GroupSource(
                        name=name,
                        retention_days=retention.get(name),
                        stream_count=_stream_count(connection, name),
                        event_count=count,
                        first_ms=first_ms,
                        last_ms=last_ms,
                        include_sheet=True,
                        note=note,
                        iter_streams=_stream_factory(connection, name),
                    )
                )
                log(f"{index}/{len(names)}  {name}  イベント {count}")
                if group_truncated:
                    warn(notice)
                    break

            connection.execute(
                "CREATE INDEX idx_events ON events(group_name, stream_name, ts, id)"
            )
            connection.commit()
            generated = datetime.now(load_timezone(options.timezone_name))
            meta = WorkbookMeta(
                account_id=account_id,
                caller_arn=caller_arn,
                region=options.region,
                start_display=window.start_display,
                end_display=window.end_display,
                timezone_name=options.timezone_name,
                output_path=str(output.resolve()) if output.is_absolute() else str(output),
                generated_display=_format_local(generated),
                generated_iso=generated.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                scanned_groups=scanned,
                empty_groups=empty_groups,
                ignore_case=options.ignore_case,
                phrases=tuple(phrases),
                notice=notice,
            )
            # 出力パスは絶対パスで表紙に残す。相対のままだと後から場所が分からない。
            meta.output_path = str(output if output.is_absolute() else output.resolve())
            results = write_xlsx(
                output,
                meta=meta,
                groups=groups,
                format_ts=format_ts,
                max_sheet_rows=options.max_sheet_rows,
            )
        finally:
            connection.close()

    event_total = sum(item.events for item in results)
    log(
        f"Excel を出力しました: {output}  "
        f"シート {len(results)}  イベント {event_total}"
    )
    if names and event_total == 0 and denied_groups == len(names) and denied_groups > 0:
        _print_switchback_required("対象のロググループをすべて拒否されました。")
        return EXIT_SWITCHBACK
    return 0


def _noted_group(name: str, retention: Optional[int], note: str) -> GroupSource:
    return GroupSource(
        name=name,
        retention_days=retention,
        stream_count=0,
        event_count=0,
        first_ms=None,
        last_ms=None,
        include_sheet=True,
        note=note,
    )


def _print_login_required(detail: str) -> None:
    print(
        "\n".join(
            [
                "================================================================",
                "[警告] aws login --remote のセッションが無効なため終了します",
                "================================================================",
                "先に次を実行して認証し、このスクリプトを再実行してください。",
                "",
                "    aws login --remote",
                "",
                "詳細:",
                "    " + one_line(detail, 500),
                "================================================================",
            ]
        ),
        file=sys.stderr,
    )


def _print_switchback_required(detail: str) -> None:
    print(
        "\n".join(
            [
                "================================================================",
                "[警告] 現在の権限では CloudWatch Logs を操作できません",
                "================================================================",
                "この処理に CodeCommit の権限は不要です。",
                "logs:DescribeLogGroups と logs:FilterLogEvents が必要です。",
                "操作権限のあるロールへスイッチバックしてから再実行してください。",
                "自動で切り替える場合は、シェルの --switchback-mode auto と",
                "--switchback-script を指定してください。",
                "",
                "詳細:",
                "    " + one_line(detail, 500),
                "================================================================",
            ]
        ),
        file=sys.stderr,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="cwlogs-export.sh",
        description="CloudWatch Logs を Excel に書き出します。通常は cwlogs-export.sh から起動します。",
    )
    parser.add_argument("--output", required=True)
    parser.add_argument("--start", required=True)
    parser.add_argument("--end", required=True)
    parser.add_argument("--region", required=True)
    parser.add_argument("--profile", default="")
    parser.add_argument("--timezone", default="Asia/Tokyo")
    parser.add_argument("--highlight", action="append", default=[])
    parser.add_argument("--log-group", action="append", default=[])
    parser.add_argument("--log-group-prefix", action="append", default=[])
    parser.add_argument("--ignore-case", action="store_true")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--max-events", type=int, default=0)
    return parser


def options_from_args(argv: Sequence[str]) -> ExportOptions:
    namespace = build_parser().parse_args(list(argv))
    return ExportOptions(
        output=Path(namespace.output),
        start=namespace.start,
        end=namespace.end,
        region=namespace.region,
        profile=namespace.profile,
        timezone_name=namespace.timezone,
        highlights=list(namespace.highlight),
        log_groups=list(namespace.log_group),
        prefixes=list(namespace.log_group_prefix),
        ignore_case=bool(namespace.ignore_case),
        force=bool(namespace.force),
        max_events=int(namespace.max_events),
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    if os.environ.get("CWLOGS_PREFLIGHT_OK") != "1":
        print(
            "[エラー] cwlogs-export.sh から実行してください。"
            "認証の確認とスイッチバックの source はシェル側で行います。",
            file=sys.stderr,
        )
        return EXIT_ERROR
    try:
        options = options_from_args(sys.argv[1:] if argv is None else argv)
        return export_logs(options)
    except SystemExit as exc:
        if isinstance(exc.code, int):
            return exc.code
        if exc.code:
            print(exc.code, file=sys.stderr)
        return EXIT_ERROR
    except BrokenPipeError:
        return EXIT_ERROR


if __name__ == "__main__":
    sys.exit(main())
