#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""時間範囲、強調、Excel、AWS CLI ページングの検証。RHEL の標準ライブラリだけ。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import cwlogs_export  # noqa: E402
import cwlogs_xlsx  # noqa: E402
from cwlogs_export import (  # noqa: E402
    AwsAuthExpired,
    ExportOptions,
    classify_aws_failure,
    export_logs,
    parse_time_window,
    resolve_group_names,
)
from cwlogs_xlsx import (  # noqa: E402
    FONT_NAME,
    GroupSource,
    StreamSource,
    WorkbookMeta,
    compile_highlights,
    highlight_runs,
    iter_events_list,
    sanitize_sheet_name,
    slice_runs,
    unique_sheet_name,
    write_xlsx,
)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def epoch_ms(moment: datetime) -> int:
    utc = moment.astimezone(timezone.utc)
    delta = utc - datetime(1970, 1, 1, tzinfo=timezone.utc)
    return (delta.days * 86400 + delta.seconds) * 1000 + delta.microseconds // 1000


def ok(payload: dict) -> tuple:
    return 0, json.dumps(payload), ""


def fail(text: str, code: int = 254) -> tuple:
    return code, "", text


class ScriptedAws:
    def __init__(self, handler):
        self.handler = handler
        self.calls = []

    def run(self, cmd, capture_output=True, text=True, encoding="utf-8", errors="replace"):
        recorded = list(cmd)
        self.calls.append(recorded)
        code, stdout, stderr = self.handler(recorded)
        return subprocess.CompletedProcess(recorded, code, stdout or "", stderr or "")


def arg_value(cmd, flag):
    if flag not in cmd:
        return None
    return cmd[cmd.index(flag) + 1]


class TimeWindowTests(unittest.TestCase):
    def test_second_precision_includes_that_second(self):
        tz = ZoneInfo("Asia/Tokyo")
        window = parse_time_window(
            "2026-04-08 00:00:00", "2026-04-08 00:00:00", "Asia/Tokyo"
        )
        start = datetime(2026, 4, 8, 0, 0, 0, tzinfo=tz)
        self.assertEqual(window.start_ms, epoch_ms(start))
        self.assertEqual(window.end_ms, epoch_ms(start + timedelta(seconds=1)))
        self.assertEqual(window.end_display, "2026-04-08 00:00:00.999")

    def test_date_only_includes_the_whole_day(self):
        tz = ZoneInfo("Asia/Tokyo")
        window = parse_time_window("2026/04/08", "2026-04-08", "Asia/Tokyo")
        start = datetime(2026, 4, 8, tzinfo=tz)
        self.assertEqual(window.start_ms, epoch_ms(start))
        self.assertEqual(window.end_ms, epoch_ms(start + timedelta(days=1)))
        self.assertEqual(window.end_display, "2026-04-08 23:59:59.999")

    def test_minute_offset_and_zulu(self):
        tz = ZoneInfo("Asia/Tokyo")
        minute = parse_time_window(
            "2026-04-08 00:00", "2026-04-08T00:00+09:00", "Asia/Tokyo"
        )
        start = datetime(2026, 4, 8, 0, 0, tzinfo=tz)
        self.assertEqual(minute.end_ms, epoch_ms(start + timedelta(minutes=1)))
        self.assertEqual(minute.end_display, "2026-04-08 00:00:59.999")
        zulu = parse_time_window(
            "2026-04-08T00:00:00Z", "2026-04-08T00:00:00Z", "Asia/Tokyo"
        )
        utc_start = datetime(2026, 4, 8, 0, 0, tzinfo=timezone.utc).astimezone(tz)
        self.assertEqual(zulu.start_ms, epoch_ms(utc_start))
        self.assertEqual(zulu.start_display, "2026-04-08 09:00:00.000")

    def test_epoch_seconds_and_millis(self):
        window = parse_time_window("1775574000", "1775574000000", "Asia/Tokyo")
        self.assertEqual(window.start_ms, 1775574000000)
        self.assertEqual(window.end_ms, 1775574000000 + 1)

    def test_end_before_start_is_rejected(self):
        with self.assertRaises(SystemExit):
            parse_time_window("2026-04-09", "2026-04-08", "Asia/Tokyo")

    def test_unknown_timezone_is_rejected(self):
        with self.assertRaises(SystemExit):
            parse_time_window("2026-04-08", "2026-04-08", "Not/AZone")


class HighlightTests(unittest.TestCase):
    def test_longer_phrase_wins(self):
        phrases = ["ERROR", "ERROR CODE"]
        pattern = compile_highlights(phrases, False)
        runs = highlight_runs("see ERROR CODE now", phrases, pattern, False)
        bold = [run.text for run in runs if run.bold]
        self.assertEqual(bold, ["ERROR CODE"])
        # 色は一致した文言の指定順。長い方を優先しても、短い方の色にはしない。
        self.assertEqual(runs[1].color, "D35400")

    def test_ignore_case(self):
        phrases = ["Error"]
        pattern = compile_highlights(phrases, True)
        runs = highlight_runs("error", phrases, pattern, True)
        self.assertTrue(runs[0].bold)
        self.assertEqual(runs[0].color, "C0392B")
        plain = highlight_runs("error", phrases, compile_highlights(phrases, False), False)
        self.assertFalse(plain[0].bold)

    def test_slice_keeps_color(self):
        pieces = slice_runs(
            [cwlogs_xlsx.Run("abcdef", bold=True, color="C0392B")], limit=4
        )
        self.assertEqual(pieces[0][0].text, "abcd")
        self.assertEqual(pieces[1][0].text, "ef")
        self.assertTrue(pieces[1][0].bold)
        self.assertEqual(pieces[1][0].color, "C0392B")

    def test_sheet_names(self):
        self.assertEqual(sanitize_sheet_name("/aws/lambda/注文"), "aws_lambda_注文")
        self.assertEqual(sanitize_sheet_name("History"), "History_log")
        self.assertEqual(sanitize_sheet_name("a/b?c*d[e]f:g\\h"), "a_b_c_d_e_f_g_h")
        used = set()
        long_name = "あ" * 40
        first = unique_sheet_name(long_name, used)
        second = unique_sheet_name(long_name, used)
        self.assertLessEqual(len(first), 31)
        self.assertLessEqual(len(second), 31)
        self.assertNotEqual(first, second)
        self.assertTrue(any("あ" in name for name in (first,)))


class WorkbookTests(unittest.TestCase):
    def test_polished_workbook(self):
        events_a = [
            (1_700_000_000_000, "ZEBRA ERROR =CMD|&<>\x00日本語", "evt-z"),
            (1_700_000_000_500, "plain line", "evt-z2"),
        ]
        events_b = [(1_700_000_006_000, "ALPHA later", "evt-a")]
        long_line = "幅" * 80

        def streams_orders():
            return [
                StreamSource(
                    "zebra",
                    3,
                    events_a[0][0],
                    1_700_000_001_000,
                    iter_events_list(events_a + [(1_700_000_001_000, long_line, "evt-long")]),
                ),
                StreamSource(
                    "alpha",
                    1,
                    events_b[0][0],
                    events_b[0][0],
                    iter_events_list(events_b),
                ),
            ]

        def streams_empty():
            return []

        groups = [
            GroupSource(
                name="/app/orders",
                retention_days=30,
                stream_count=2,
                event_count=4,
                first_ms=events_a[0][0],
                last_ms=events_b[0][0],
                include_sheet=True,
                note="指定の --max-events に達したため、このロググループは途中までです。",
                iter_streams=streams_orders,
            ),
            GroupSource(
                name="/app/empty",
                retention_days=None,
                stream_count=0,
                event_count=0,
                first_ms=None,
                last_ms=None,
                include_sheet=True,
                note="この期間に出力されたイベントはありません。",
                iter_streams=streams_empty,
            ),
        ]
        meta = WorkbookMeta(
            account_id="123456789012",
            caller_arn="arn:aws:iam::123456789012:user/dev",
            region="ap-northeast-1",
            start_display="2026-04-08 00:00:00.000",
            end_display="2026-04-08 23:59:59.999",
            timezone_name="Asia/Tokyo",
            output_path="/tmp/out.xlsx",
            generated_display="2026-04-09 09:00:00.000",
            generated_iso="2026-04-09T00:00:00Z",
            scanned_groups=2,
            empty_groups=1,
            ignore_case=False,
            phrases=("ERROR", "WARN"),
            notice="取得を途中で止めています。",
        )

        def format_ts(millis: int) -> str:
            return "ts-" + str(millis)

        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "logs.xlsx"
            results = write_xlsx(path, meta=meta, groups=groups, format_ts=format_ts)
            self.assertEqual(oct(path.stat().st_mode & 0o777), "0o600")
            files, names = load_book(path)
            self.assertEqual(names[0], "概要")
            self.assertIn("app_orders", names)
            self.assertIn("app_empty", names)
            styles = files["xl/styles.xml"].decode("utf-8")
            self.assertIn(FONT_NAME, styles)
            cover = files["xl/worksheets/sheet1.xml"].decode("utf-8")
            self.assertIn("取得を途中で止めています。", cover)
            self.assertIn("途中までです", cover)
            self.assertIn("app_orders'!A1", cover)
            self.assertNotIn("<f>", cover)
            order_sheet = sheet_by_name(files, names, "app_orders")
            self.assertIn('t="inlineStr"', order_sheet)
            self.assertIn("=CMD|&amp;&lt;&gt;", order_sheet)
            self.assertNotIn("=CMD|&<>", order_sheet)
            self.assertIn("FFC0392B", order_sheet)
            self.assertIn(FONT_NAME, order_sheet)
            text = ET.fromstring(order_sheet.encode("utf-8"))
            joined = "".join(text.itertext())
            self.assertLess(joined.index("ZEBRA"), joined.index("alpha"))
            self.assertIn("ログストリーム   zebra", joined)
            self.assertIn("ログストリーム   alpha", joined)
            empty_text = "".join(
                ET.fromstring(sheet_by_name(files, names, "app_empty").encode("utf-8")).itertext()
            )
            self.assertIn("無期限", empty_text)
            self.assertIn("この期間に出力されたイベントはありません。", empty_text)
            widths = column_widths(order_sheet)
            self.assertGreaterEqual(widths[2], 48)
            self.assertLessEqual(widths[2], 120)
            self.assertIn("showGridLines=\"0\"", order_sheet)
            self.assertTrue(any(item.status.startswith("指定の") for item in results))

    def test_continuation_sheet_repeats_stream_header(self):
        events = [(1_700_000_000_000 + i, "line %s ERROR" % i, "id%s" % i) for i in range(15)]

        def streams():
            return [
                StreamSource(
                    "main",
                    len(events),
                    events[0][0],
                    events[-1][0],
                    iter_events_list(events),
                )
            ]

        group = GroupSource(
            name="orders",
            retention_days=7,
            stream_count=1,
            event_count=len(events),
            first_ms=events[0][0],
            last_ms=events[-1][0],
            include_sheet=True,
            iter_streams=streams,
        )
        meta = sample_meta()
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "split.xlsx"
            write_xlsx(
                path,
                meta=meta,
                groups=[group],
                format_ts=lambda millis: str(millis),
                max_sheet_rows=12,
            )
            _files, names = load_book(path)
            self.assertGreaterEqual(len(names), 3)
            book = zipfile.ZipFile(path)
            blob = "\n".join(
                book.read(name).decode("utf-8")
                for name in book.namelist()
                if name.startswith("xl/worksheets/")
            )
            book.close()
            self.assertIn("（続き）", blob)
            self.assertIn("ログストリーム   main    （続き）", blob)


class FetchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.out = Path(self.tmp.name) / "out.xlsx"

    def options(self, **kwargs):
        values = dict(
            output=self.out,
            start="2026-04-08 00:00:00",
            end="2026-04-08 00:00:01",
            region="ap-northeast-1",
            timezone_name="Asia/Tokyo",
            highlights=["ERROR"],
            force=True,
            sleeper=lambda _delay: None,
        )
        values.update(kwargs)
        return ExportOptions(**values)

    def patch_aws(self, handler):
        scripted = ScriptedAws(handler)
        patcher = mock.patch("cwlogs_export.subprocess.run", scripted.run)
        patcher.start()
        self.addCleanup(patcher.stop)
        return scripted

    def test_pages_order_streams_by_first_timestamp(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/Login/user"})
            if "describe-log-groups" in cmd and "--next-token" not in cmd:
                return ok(
                    {
                        "logGroups": [{"logGroupName": "/app/orders", "retentionInDays": 30}],
                        "nextToken": "groups-2",
                    }
                )
            if "describe-log-groups" in cmd:
                return ok({"logGroups": [{"logGroupName": "/app/billing", "retentionInDays": 14}]})
            group = arg_value(cmd, "--log-group-name")
            if group == "/app/billing" and "--next-token" not in cmd:
                return ok(
                    {
                        "events": [
                            {
                                "timestamp": 5000,
                                "message": "ALPHA later",
                                "logStreamName": "alpha",
                                "eventId": "a1",
                            }
                        ],
                        "nextToken": "bill-2",
                    }
                )
            if group == "/app/billing":
                return ok(
                    {
                        "events": [
                            {
                                "timestamp": 1000,
                                "message": "ZEBRA ERROR first",
                                "logStreamName": "zebra",
                                "eventId": "z1",
                            }
                        ]
                    }
                )
            if group == "/app/orders":
                return ok(
                    {
                        "events": [
                            {
                                "timestamp": 2000,
                                "message": "orders ERROR",
                                "logStreamName": "s",
                                "eventId": "o1",
                            }
                        ]
                    }
                )
            return fail("unexpected " + " ".join(cmd), 1)

        scripted = self.patch_aws(handler)
        sleeps = []
        code = export_logs(self.options(sleeper=sleeps.append))
        self.assertEqual(code, 0)
        self.assertTrue(self.out.is_file())
        files, names = load_book(self.out)
        self.assertIn("app_billing", names)
        self.assertIn("app_orders", names)
        billing = "".join(ET.fromstring(sheet_by_name(files, names, "app_billing")).itertext())
        self.assertLess(billing.index("ZEBRA"), billing.index("alpha"))
        self.assertIn("orders ERROR", "".join(ET.fromstring(sheet_by_name(files, names, "app_orders")).itertext()))
        filter_calls = [cmd for cmd in scripted.calls if "filter-log-events" in cmd]
        self.assertGreaterEqual(len(filter_calls), 3)
        self.assertIn("--interleaved", filter_calls[0])
        self.assertEqual(arg_value(filter_calls[0], "--limit"), "10000")
        window = parse_time_window("2026-04-08 00:00:00", "2026-04-08 00:00:01", "Asia/Tokyo")
        self.assertEqual(arg_value(filter_calls[0], "--start-time"), str(window.start_ms))
        self.assertEqual(arg_value(filter_calls[0], "--end-time"), str(window.end_ms))
        self.assertEqual(sleeps, [])

    def test_one_group_denied_others_continue(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd:
                name = arg_value(cmd, "--log-group-name-prefix")
                return ok({"logGroups": [{"logGroupName": name, "retentionInDays": 3}]})
            group = arg_value(cmd, "--log-group-name")
            if group == "/no":
                return fail("AccessDeniedException: not authorized to perform logs:FilterLogEvents")
            return ok(
                {
                    "events": [
                        {
                            "timestamp": 10,
                            "message": "kept ERROR",
                            "logStreamName": "s",
                            "eventId": "e",
                        }
                    ]
                }
            )

        self.patch_aws(handler)
        code = export_logs(self.options(log_groups=["/ok", "/no"]))
        self.assertEqual(code, 0)
        files, names = load_book(self.out)
        self.assertIn("ok", names)
        self.assertTrue(any(name.startswith("no") for name in names))
        denied = "".join(
            ET.fromstring(sheet_by_name(files, names, next(n for n in names if n.startswith("no")))).itertext()
        )
        self.assertIn("権限がありません", denied)
        kept = "".join(ET.fromstring(sheet_by_name(files, names, "ok")).itertext())
        self.assertIn("kept ERROR", kept)

    def test_expired_token_writes_nothing(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd:
                return ok({"logGroups": [{"logGroupName": "/app/orders"}]})
            return fail("ExpiredToken: The security token included in the request is expired")

        self.patch_aws(handler)
        code = export_logs(self.options())
        self.assertEqual(code, 2)
        self.assertFalse(self.out.exists())

    def test_throttle_then_success(self):
        state = {"filters": 0}
        sleeps = []

        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd:
                return ok({"logGroups": [{"logGroupName": "/app/orders", "retentionInDays": 1}]})
            state["filters"] += 1
            if state["filters"] == 1:
                return fail("ThrottlingException: Rate exceeded")
            return ok(
                {
                    "events": [
                        {
                            "timestamp": 10,
                            "message": "after throttle",
                            "logStreamName": "s",
                            "eventId": "e",
                        }
                    ]
                }
            )

        self.patch_aws(handler)
        code = export_logs(self.options(sleeper=sleeps.append))
        self.assertEqual(code, 0)
        self.assertEqual(sleeps, [1.0])
        _files, names = load_book(self.out)
        self.assertIn("app_orders", names)

    def test_explicit_empty_sheet_and_discovered_empty_omitted(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd:
                prefix = arg_value(cmd, "--log-group-name-prefix")
                if prefix == "/app/":
                    return ok(
                        {
                            "logGroups": [
                                {"logGroupName": "/app/quiet"},
                                {"logGroupName": "/app/loud"},
                            ]
                        }
                    )
                return ok({"logGroups": [{"logGroupName": prefix}] if prefix else []})
            group = arg_value(cmd, "--log-group-name")
            if group == "/app/loud":
                return ok(
                    {
                        "events": [
                            {
                                "timestamp": 10,
                                "message": "hello",
                                "logStreamName": "s",
                                "eventId": "e",
                            }
                        ]
                    }
                )
            if group == "/explicit/empty":
                return ok({"events": []})
            return ok({"events": []})

        self.patch_aws(handler)
        code = export_logs(
            self.options(log_groups=["/explicit/empty"], prefixes=["/app/"])
        )
        self.assertEqual(code, 0)
        files, names = load_book(self.out)
        self.assertIn("explicit_empty", names)
        self.assertIn("app_loud", names)
        self.assertNotIn("app_quiet", names)
        empty = "".join(ET.fromstring(sheet_by_name(files, names, "explicit_empty")).itertext())
        self.assertIn("この期間に出力されたイベントはありません。", empty)

    def test_every_group_denied_exits_switchback_and_keeps_workbook(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd:
                name = arg_value(cmd, "--log-group-name-prefix")
                return ok({"logGroups": [{"logGroupName": name}]})
            return fail("AccessDeniedException: not authorized")

        self.patch_aws(handler)
        code = export_logs(self.options(log_groups=["/a", "/b"]))
        self.assertEqual(code, 3)
        self.assertTrue(self.out.is_file())

    def test_retention_lookup_does_not_swallow_expired_token(self):
        def handler(cmd):
            if "describe-log-groups" in cmd:
                return fail("ExpiredToken: The security token included in the request is expired")
            return fail("unexpected", 1)

        self.patch_aws(handler)
        with self.assertRaises(AwsAuthExpired):
            resolve_group_names(self.options(log_groups=["/only"]))

    def test_max_events_notice(self):
        def handler(cmd):
            if "get-caller-identity" in cmd:
                return ok({"Account": "1", "Arn": "arn:aws:iam::1:user/a"})
            if "describe-log-groups" in cmd and "--next-token" not in cmd:
                return ok(
                    {
                        "logGroups": [
                            {"logGroupName": "/app/a"},
                            {"logGroupName": "/app/b"},
                        ]
                    }
                )
            group = arg_value(cmd, "--log-group-name")
            if group == "/app/a":
                return ok(
                    {
                        "events": [
                            {
                                "timestamp": 10,
                                "message": "one",
                                "logStreamName": "s",
                                "eventId": "e1",
                            },
                            {
                                "timestamp": 11,
                                "message": "two",
                                "logStreamName": "s",
                                "eventId": "e2",
                            },
                        ]
                    }
                )
            return ok(
                {
                    "events": [
                        {
                            "timestamp": 12,
                            "message": "should-not",
                            "logStreamName": "s",
                            "eventId": "e3",
                        }
                    ]
                }
            )

        scripted = self.patch_aws(handler)
        code = export_logs(self.options(max_events=1))
        self.assertEqual(code, 0)
        files, names = load_book(self.out)
        cover = files["xl/worksheets/sheet1.xml"].decode("utf-8")
        self.assertIn("--max-events 1", cover)
        self.assertNotIn("app_b", names)
        groups = [cmd for cmd in scripted.calls if "filter-log-events" in cmd]
        self.assertEqual(len(groups), 1)

    def test_main_refuses_without_preflight(self):
        previous = os.environ.pop("CWLOGS_PREFLIGHT_OK", None)
        try:
            code = cwlogs_export.main(
                ["--output", "a.xlsx", "--start", "2026-04-08", "--end", "2026-04-08", "--region", "ap-northeast-1"]
            )
            self.assertEqual(code, 1)
        finally:
            if previous is not None:
                os.environ["CWLOGS_PREFLIGHT_OK"] = previous

    def test_classify_order(self):
        self.assertEqual(
            classify_aws_failure("AccessDeniedException and ExpiredToken"), "auth"
        )
        self.assertEqual(classify_aws_failure("ResourceNotFoundException"), "missing")
        self.assertEqual(classify_aws_failure("ThrottlingException: Rate exceeded"), "throttle")
        self.assertEqual(classify_aws_failure("User is not authorized"), "denied")


def sample_meta() -> WorkbookMeta:
    return WorkbookMeta(
        account_id="1",
        caller_arn="arn:aws:iam::1:user/a",
        region="ap-northeast-1",
        start_display="2026-04-08 00:00:00.000",
        end_display="2026-04-08 00:00:01.999",
        timezone_name="Asia/Tokyo",
        output_path="/tmp/out.xlsx",
        generated_display="2026-04-09 09:00:00.000",
        generated_iso="2026-04-09T00:00:00Z",
        scanned_groups=1,
        empty_groups=0,
        ignore_case=False,
        phrases=("ERROR",),
    )


def load_book(path: Path):
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    for name, data in files.items():
        if name.endswith(".xml"):
            ET.fromstring(data)
    root = ET.fromstring(files["xl/workbook.xml"])
    names = [element.attrib["name"] for element in root.findall("m:sheets/m:sheet", NS)]
    return files, names


def sheet_by_name(files, names, sheet_name: str) -> str:
    index = names.index(sheet_name) + 1
    return files["xl/worksheets/sheet%d.xml" % index].decode("utf-8")


def column_widths(sheet_xml: str):
    root = ET.fromstring(sheet_xml.encode("utf-8"))
    widths = []
    for col in root.findall("m:cols/m:col", NS):
        widths.append(float(col.attrib["width"]))
    return widths


if __name__ == "__main__":
    unittest.main()
