# -*- coding: utf-8 -*-
"""台账删除能力回归测试。

台账是质量审核的审计留痕，"能删"必须与"删得可追溯、可恢复、不被误用"同时成立。
这里锁定五件事：
1. 只按「哈希+时间+文件名」定位记录（台账会因去重重排，下标不可靠）；
2. 软删除：被删记录原文写入 .trash 留痕，汇总报告里列出"已移除"小节；
3. 产物是**移动**到 .trash（默认不动产物），不是物理删除；
4. HTTP 接口受管理令牌保护，未授权一律拒绝；
5. 汇总报告的小节渲染不因留痕为空而报错。
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

import openpyxl

PROJECT = Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import audit_history as H  # noqa: E402
import server  # noqa: E402


def _entry(name: str, digest: str, ts: str, *, run_id: str = "", rows: int = 100, issues: int = 5) -> dict:
    return {
        "ts": ts, "file": name, "hash": digest, "run_id": run_id,
        "artifacts": H.ARTIFACT_AVAILABLE if run_id else "",
        "repeats": 1, "first_ts": ts, "sheet": "S", "rows": rows, "issues": issues,
        "priority": {"P0": 1, "P1": 2, "P2": 2}, "ai_status": "completed",
        "ai_candidate": 3, "ai_reviewed": 3,
        "final_counts": {"confirmed_issue": 5, "likely_false_positive": 0, "needs_human": 0},
        "conclusion": "测试结论", "incremental": False,
    }


class RemoveLedgerEntryTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.run_id = "20260101T090000_abcdef01"

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _write_ledger(self, entries: list[dict]) -> None:
        (self.dir / H.LEDGER_NAME).write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")

    def test_entry_key_uses_business_fields_not_index(self):
        entry = _entry("a.xlsx", "deadbeefdeadbeef", "2026-01-01 09:00")
        self.assertEqual("deadbeefdeadbeef|2026-01-01 09:00|a.xlsx", H.entry_key(entry))
        # 只有哈希的旧写法也要能命中（兼容旧前端/手工调用）
        self.assertIn("deadbeefdeadbeef||", H.entry_keys(entry))

    def test_remove_only_target_entry(self):
        keep = _entry("keep.xlsx", "aa" * 8, "2026-01-01 08:00")
        drop = _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00")
        self._write_ledger([keep, drop])
        result = H.remove_ledger_entries(self.dir, {H.entry_key(drop)})
        self.assertTrue(result["ok"])
        self.assertEqual(1, result["remaining"])
        self.assertEqual([drop], result["removed"])
        remaining = H.load_ledger(self.dir / H.LEDGER_NAME)
        self.assertEqual(["keep.xlsx"], [e["file"] for e in remaining])

    def test_removal_is_traced_in_trash(self):
        drop = _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00")
        self._write_ledger([drop])
        H.remove_ledger_entries(self.dir, {H.entry_key(drop)}, operator="admin-token")
        records = H.load_trash_records(self.dir)
        self.assertEqual(1, len(records))
        self.assertEqual("admin-token", records[0]["operator"])
        self.assertEqual(["drop.xlsx"], [e["file"] for e in records[0]["entries"]])
        self.assertEqual(1, records[0]["ledger_before"])
        self.assertEqual(0, records[0]["ledger_after"])
        self.assertTrue(Path(records[0]["_file"]).name.startswith("removed-"))

    def test_artifacts_are_moved_not_deleted(self):
        run_dir = self.dir / self.run_id
        (run_dir / "input").mkdir(parents=True)
        (run_dir / "input" / "drop.xlsx").write_bytes(b"x")
        (run_dir / "drop_审核报告.html").write_text("<html></html>", encoding="utf-8")
        drop = _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00", run_id=self.run_id)
        self._write_ledger([drop])

        result = H.remove_ledger_entries(self.dir, {H.entry_key(drop)}, remove_artifacts=True)
        self.assertTrue(result["ok"])
        self.assertFalse(run_dir.exists(), "run 目录应已移出 runs/")
        self.assertTrue(result["trashed_paths"], "应返回产物移动后的路径以便恢复")
        for moved in result["trashed_paths"]:
            self.assertTrue(Path(moved).is_dir(), "产物必须是移动而非删除")
            self.assertIn(H.TRASH_DIR_NAME, moved)

    def test_artifacts_kept_by_default(self):
        run_dir = self.dir / self.run_id
        run_dir.mkdir(parents=True)
        drop = _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00", run_id=self.run_id)
        self._write_ledger([drop])
        H.remove_ledger_entries(self.dir, {H.entry_key(drop)})
        self.assertTrue(run_dir.is_dir(), "默认只移除记录，产物必须原地保留")
        self.assertEqual([], H.load_trash_records(self.dir)[0]["trashed_paths"])

    def test_unknown_key_reports_miss_without_touching_ledger(self):
        keep = _entry("keep.xlsx", "aa" * 8, "2026-01-01 08:00")
        self._write_ledger([keep])
        result = H.remove_ledger_entries(self.dir, {"nosuchkey||"})
        self.assertFalse(result["ok"])
        self.assertIn("没有匹配", result["error"])
        self.assertEqual(1, len(H.load_ledger(self.dir / H.LEDGER_NAME)))
        self.assertEqual([], H.load_trash_records(self.dir), "未删除时不应留痕")

    def test_render_shows_removal_section(self):
        drop = _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00")
        self._write_ledger([drop])
        H.remove_ledger_entries(self.dir, {H.entry_key(drop)})
        out = self.dir / "summary.html"
        H.render_history_html(
            out, H.load_ledger(self.dir / H.LEDGER_NAME), removed=H.load_trash_records(self.dir),
        )
        html = out.read_text(encoding="utf-8")
        self.assertIn("已移除的台账记录（留痕）", html)
        self.assertIn("drop.xlsx", html)
        self.assertIn("已移除 1 条", html)

    def test_render_without_trail_still_works(self):
        keep = _entry("keep.xlsx", "aa" * 8, "2026-01-01 08:00")
        out = self.dir / "summary.html"
        H.render_history_html(out, [keep])
        html = out.read_text(encoding="utf-8")
        self.assertIn("审核历史明细", html)
        self.assertNotIn("已移除的台账记录（留痕）", html)


class WorkbookHandleTests(unittest.TestCase):
    """审核结束后必须释放输入 xlsx 的文件句柄。

    这曾经是真 bug：openpyxl 的 read_only 模式在 `wb.close()` 前一直占着文件，
    审核跑完 `runs/<run_id>/input/*.xlsx` 仍被占用，Windows 上"移除产物 / 移动
    run 目录"报 WinError 32。用"能不能重命名该目录"来证明句柄已释放。
    """

    def test_audit_releases_input_workbook(self):
        import run_audit
        import report_business

        d = Path(tempfile.mkdtemp())
        try:
            src = d / "small.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "KKS编码"
            ws.append(["KKS编码", "父级KKS", "设备名称", "机组", "专业", "系统"])
            ws.append(["00LBA10GP001", "00LBA10", "轴流风机", "全厂公用", "暖通", "通风系统"])
            wb.save(src)
            wb.close()

            out = d / "runs" / f"{datetime.now():%Y%m%dT%H%M%S}_abcdef01"
            run_audit.audit_file(src, out, run_ai_review=False, write_outputs=True)

            # 报告侧的"重码/待处理副表"读取曾因提前 return 漏关工作簿，
            # 这里直接断言它不再持有句柄：能重命名源文件即说明句柄已释放
            report_business._read_aux_dup_sheet(str(src))
            probe = src.with_name(src.name + ".probe")
            try:
                src.rename(probe)
            except OSError as exc:  # pragma: no cover - 只在句柄泄漏时触发
                self.fail(f"重码副表读取后仍占用源工作簿：{exc}")
            probe.rename(src)

            moved = out.with_name(out.name + ".moved")
            try:
                out.rename(moved)
            except OSError as exc:  # pragma: no cover - 只在句柄泄漏时触发
                self.fail(f"审核结束后输入工作簿仍被占用，无法移动产物目录：{exc}")
            self.assertTrue(moved.is_dir())
            self.assertTrue((moved / "small_审核报告.html").is_file())
        finally:
            shutil.rmtree(d, ignore_errors=True)


class TrendCacheTests(unittest.TestCase):
    """趋势总结：台账没变就不该再花一次模型调用。"""

    def _ledger(self, issues: int = 5) -> list[dict]:
        return [_entry("a.xlsx", "aa" * 8, "2026-01-01 08:00"), {**_entry("b.xlsx", "bb" * 8, "2026-01-01 09:00"), "issues": issues}]

    def test_fingerprint_ignores_volatile_fields(self):
        base = self._ledger()
        changed = [dict(e) for e in base]
        changed[-1]["ts"] = "2030-01-01 00:00"
        changed[-1]["repeats"] = 99
        changed[-1]["run_id"] = "20300101T000000_ffffffff"
        self.assertEqual(H.ledger_fingerprint(base), H.ledger_fingerprint(changed))

    def test_fingerprint_changes_with_business_data(self):
        self.assertNotEqual(H.ledger_fingerprint(self._ledger(5)), H.ledger_fingerprint(self._ledger(6)))

    def test_trend_switch_defaults_on(self):
        original = os.environ.pop("AI_TREND_SUMMARY_ENABLED", None)
        try:
            self.assertTrue(H.trend_enabled())
            for raw in ("0", "false", "no", "off"):
                os.environ["AI_TREND_SUMMARY_ENABLED"] = raw
                self.assertFalse(H.trend_enabled(), raw)
            os.environ["AI_TREND_SUMMARY_ENABLED"] = "1"
            self.assertTrue(H.trend_enabled())
        finally:
            os.environ.pop("AI_TREND_SUMMARY_ENABLED", None)
            if original is not None:
                os.environ["AI_TREND_SUMMARY_ENABLED"] = original

    def test_unchanged_ledger_reuses_cached_trend(self):
        """指纹一致时必须复用缓存，且不再调用 summarize_ledger。

        做法：真实跑两次 record_and_render（同一份台账、同一个 result）。
        第一次会调模型并写入带指纹的缓存；第二次业务字段完全没变，必须直接复用缓存。
        """
        d = Path(tempfile.mkdtemp())
        try:
            runs = d / "runs"
            run_dir = runs / "20260101T090000_abcdef01"
            run_dir.mkdir(parents=True)
            src = d / "a.xlsx"
            src.write_bytes(b"x")
            first = _entry("a.xlsx", H._file_digest(src), "2026-01-01 08:00", rows=1, issues=0)
            second = _entry("b.xlsx", "bb" * 8, "2026-01-01 09:00", rows=1, issues=0)
            (runs / H.LEDGER_NAME).write_text(json.dumps([first, second], ensure_ascii=False), encoding="utf-8")

            result = {
                "data_rows": first["rows"], "issue_count": first["issues"],
                "priority_counts": first["priority"], "sheet": first["sheet"],
                "ai_review": {"status": first["ai_status"], "candidate_count": first["ai_candidate"], "reviewed_count": first["ai_reviewed"]},
                "final_decision_counts": first["final_counts"], "conclusion": first["conclusion"],
                "scope_validation": {"incremental_batch": first["incremental"]},
            }

            calls: list[int] = []
            original = H.summarize_ledger
            H.summarize_ledger = lambda ledger, progress_callback=None: calls.append(1) or {
                "overall": "第一次真算的趋势结论", "points": [], "model": "demo",
            }
            try:
                H.record_and_render(src, run_dir, dict(result))
                self.assertEqual(1, len(calls), "第一次应调用模型生成趋势")
                fingerprint = H.load_trend(runs).get("fingerprint")
                self.assertTrue(fingerprint, "缓存里必须写入台账指纹")

                calls.clear()
                H.record_and_render(src, run_dir, dict(result))
                self.assertEqual([], calls, "台账业务字段未变时不应再调模型")
            finally:
                H.summarize_ledger = original
            self.assertEqual("第一次真算的趋势结论", H.load_trend(runs).get("overall"))
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def test_switch_off_skips_model_call(self):
        d = Path(tempfile.mkdtemp())
        try:
            runs = d / "runs"
            run_dir = runs / "20260101T090000_abcdef01"
            run_dir.mkdir(parents=True)
            ledger = self._ledger()
            (runs / H.LEDGER_NAME).write_text(json.dumps(ledger, ensure_ascii=False), encoding="utf-8")
            calls: list[int] = []
            original = H.summarize_ledger
            H.summarize_ledger = lambda ledger, progress_callback=None: calls.append(1) or {"overall": "不该被调用"}
            os.environ["AI_TREND_SUMMARY_ENABLED"] = "0"
            try:
                src = d / "a.xlsx"
                src.write_bytes(b"x")
                H.record_and_render(src, run_dir, {"data_rows": 1, "issue_count": 0})
            finally:
                H.summarize_ledger = original
                os.environ.pop("AI_TREND_SUMMARY_ENABLED", None)
            self.assertEqual([], calls, "开关关闭时不应调模型")
        finally:
            shutil.rmtree(d, ignore_errors=True)


class HistoryDeleteHttpTests(unittest.TestCase):
    """按真实 HTTP 语义验证删除接口的鉴权与效果。"""

    TOKEN = "delete-token-0123456789"

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.runs = Path(cls._tmp.name)
        cls.run_id = "20260101T090000_abcdef01"
        cls._httpd = server.ThreadingHTTPServer(("127.0.0.1", 0), server.AuditHandler)
        cls._httpd.runs_dir = cls.runs.resolve()
        cls._httpd.max_upload_bytes = server.MAX_UPLOAD_BYTES
        cls._httpd.jobs = {}
        cls._httpd.jobs_lock = threading.Lock()
        cls._port = cls._httpd.server_address[1]
        cls._thread = threading.Thread(target=cls._httpd.serve_forever, daemon=True)
        cls._thread.start()

    @classmethod
    def tearDownClass(cls):
        cls._httpd.shutdown()
        cls._httpd.server_close()
        cls._thread.join(timeout=5)
        cls._tmp.cleanup()

    def setUp(self):
        self._original = os.environ.get(server.ADMIN_TOKEN_ENV)
        os.environ[server.ADMIN_TOKEN_ENV] = self.TOKEN
        # 每个用例都从固定台账开始：删除类用例会真的改文件，不能靠执行顺序隔离
        shutil.rmtree(self.runs / H.TRASH_DIR_NAME, ignore_errors=True)
        shutil.rmtree(self.runs / self.run_id, ignore_errors=True)
        run_dir = self.runs / self.run_id
        (run_dir / "input").mkdir(parents=True)
        (run_dir / "input" / "drop.xlsx").write_bytes(b"x")
        ledger = [
            _entry("keep.xlsx", "aa" * 8, "2026-01-01 08:00"),
            _entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00", run_id=self.run_id),
        ]
        (self.runs / H.LEDGER_NAME).write_text(json.dumps(ledger, ensure_ascii=False), encoding="utf-8")

    def tearDown(self):
        if self._original is None:
            os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        else:
            os.environ[server.ADMIN_TOKEN_ENV] = self._original

    def _delete(self, payload: dict, token: str | None = None) -> tuple[int, dict]:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            f"http://127.0.0.1:{self._port}/api/history/delete", data=body, method="POST",
            headers={"Content-Type": "application/json"},
        )
        if token is not None:
            request.add_header("Authorization", f"Bearer {token}")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as exc:
            raw = exc.read()
            try:
                return exc.code, json.loads(raw)
            except json.JSONDecodeError:
                return exc.code, {}

    def test_rejects_anonymous_and_wrong_token(self):
        status, _ = self._delete({"keys": ["x||"]})
        self.assertEqual(401, status)
        status, _ = self._delete({"keys": ["x||"]}, token="wrong")
        self.assertEqual(401, status)

    def test_fails_closed_when_token_unset(self):
        os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        status, _ = self._delete({"keys": ["x||"]}, token=self.TOKEN)
        self.assertEqual(403, status)

    def test_missing_keys_is_rejected(self):
        status, body = self._delete({}, token=self.TOKEN)
        self.assertEqual(400, status)
        self.assertIn("keys", body.get("error", ""))

    def test_unknown_key_is_404_and_keeps_ledger(self):
        status, body = self._delete({"keys": ["nosuch||"]}, token=self.TOKEN)
        self.assertEqual(404, status)
        self.assertEqual(2, len(H.load_ledger(self.runs / H.LEDGER_NAME)))

    def test_delete_removes_entry_and_moves_artifacts(self):
        key = H.entry_key(_entry("drop.xlsx", "bb" * 8, "2026-01-01 09:00", run_id=self.run_id))
        status, body = self._delete({"keys": [key], "remove_artifacts": True}, token=self.TOKEN)
        self.assertEqual(200, status, body)
        self.assertTrue(body["ok"])
        self.assertIn("已移除 1 条", body["message"])
        self.assertEqual(["keep.xlsx"], [e["file"] for e in H.load_ledger(self.runs / H.LEDGER_NAME)])
        self.assertEqual(["keep.xlsx"], [e["file"] for e in body["entries"]])
        self.assertFalse((self.runs / self.run_id).exists())
        self.assertTrue(H.load_trash_records(self.runs), "删除必须留痕")
        summary = (self.runs / H.HISTORY_NAME).read_text(encoding="utf-8")
        self.assertIn("已移除的台账记录（留痕）", summary)

    def test_repeat_delete_is_404_not_silent_success(self):
        key = H.entry_key(_entry("keep.xlsx", "aa" * 8, "2026-01-01 08:00"))
        status, _ = self._delete({"keys": [key]}, token=self.TOKEN)
        self.assertEqual(200, status)
        status, body = self._delete({"keys": [key]}, token=self.TOKEN)
        self.assertEqual(404, status)
        self.assertIn("没有匹配", body.get("error", ""))


if __name__ == "__main__":
    unittest.main()
