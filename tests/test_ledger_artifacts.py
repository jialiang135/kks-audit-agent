# -*- coding: utf-8 -*-
"""台账产物索引与保留策略回归测试。

对应需求："审计台账每行可以直接重新下载该次生成的 HTML 报告与问题清单 Excel"。
台账按内容哈希去重，因此条目本身不携带产物位置，本次为其补上 run_id 索引，
并配套：
  - 老记录回填（按输入文件内容哈希精确匹配，不依赖时间戳）；
  - 90 天保留策略（只删名字符合 run_id 格式的目录，删后把台账标记为已归档）；
  - 受管理令牌保护的历史产物下载接口。
"""
import json
import os
import threading
import unittest
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory

import audit_history as ah
import server
from run_audit import output_artifact_names

RUN_ID_OLD = "20260101T000000_aaaaaaaa"
RUN_ID_NEW = "20260901T120000_bbbbbbbb"


class EntryRunIdTests(unittest.TestCase):
    """台账条目必须带上产物索引。"""

    def test_run_id_taken_from_output_dir_name(self):
        self.assertEqual(ah._run_id_of(Path("runs") / RUN_ID_NEW), RUN_ID_NEW)

    def test_cli_output_dir_has_no_run_id(self):
        self.assertEqual(ah._run_id_of(Path("outputs/sample")), "")

    def test_entry_records_run_id_and_available_state(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "样本.xlsx"
            source.write_bytes(b"payload")
            entry = ah._entry(source, {}, RUN_ID_NEW)
            self.assertEqual(entry["run_id"], RUN_ID_NEW)
            self.assertEqual(entry["artifacts"], ah.ARTIFACT_AVAILABLE)

    def test_entry_without_run_id_has_empty_artifact_state(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "样本.xlsx"
            source.write_bytes(b"payload")
            entry = ah._entry(source, {}, "")
            self.assertEqual(entry["run_id"], "")
            self.assertEqual(entry["artifacts"], "")

    def test_dedup_refreshes_run_id_to_latest_run(self):
        """同一文件重复上传：不新增记录，run_id 必须指向最新一次运行。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            for run_id in (RUN_ID_OLD, RUN_ID_NEW):
                run = root / run_id
                (run / "input").mkdir(parents=True)
                (run / "input" / "样本.xlsx").write_bytes(b"same-bytes")
            ledger_dir = root / RUN_ID_NEW
            ah.append_ledger(ledger_dir / "input" / "样本.xlsx", ledger_dir, {})
            ledger = ah.append_ledger(ledger_dir / "input" / "样本.xlsx", ledger_dir, {})
            self.assertEqual(len(ledger), 1)
            self.assertEqual(ledger[0]["run_id"], RUN_ID_NEW)
            self.assertEqual(ledger[0]["repeats"], 2)


class BackfillTests(unittest.TestCase):
    """老记录没有 run_id，用输入文件内容哈希回填。"""

    def _seed(self, root: Path, run_id: str, name: str, payload: bytes) -> None:
        run = root / run_id
        (run / "input").mkdir(parents=True, exist_ok=True)
        (run / "input" / name).write_bytes(payload)

    def test_backfill_matches_by_content_hash(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed(root, RUN_ID_NEW, "8.26-舟山.xlsx", b"abc")
            digest = ah._file_digest(root / RUN_ID_NEW / "input" / "8.26-舟山.xlsx")
            (root / ah.LEDGER_NAME).write_text(
                json.dumps([{"file": "8.26-舟山.xlsx", "hash": digest}], ensure_ascii=False),
                encoding="utf-8",
            )
            self.assertEqual(ah.backfill_run_ids(root), 1)
            entry = ah.load_ledger(root / ah.LEDGER_NAME)[0]
            self.assertEqual(entry["run_id"], RUN_ID_NEW)
            self.assertEqual(entry["artifacts"], ah.ARTIFACT_AVAILABLE)

    def test_backfill_is_idempotent(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed(root, RUN_ID_NEW, "a.xlsx", b"abc")
            digest = ah._file_digest(root / RUN_ID_NEW / "input" / "a.xlsx")
            (root / ah.LEDGER_NAME).write_text(
                json.dumps([{"file": "a.xlsx", "hash": digest}], ensure_ascii=False), encoding="utf-8"
            )
            self.assertEqual(ah.backfill_run_ids(root), 1)
            self.assertEqual(ah.backfill_run_ids(root), 0)

    def test_backfill_leaves_unknown_hash_alone(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed(root, RUN_ID_NEW, "a.xlsx", b"abc")
            (root / ah.LEDGER_NAME).write_text(
                json.dumps([{"file": "绝迹.xlsx", "hash": "ffffffffffffffff"}], ensure_ascii=False),
                encoding="utf-8",
            )
            self.assertEqual(ah.backfill_run_ids(root), 0)
            self.assertFalse(ah.load_ledger(root / ah.LEDGER_NAME)[0].get("run_id"))

    def test_backfill_prefers_newest_run(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed(root, RUN_ID_OLD, "a.xlsx", b"same")
            self._seed(root, RUN_ID_NEW, "a.xlsx", b"same")
            digest = ah._file_digest(root / RUN_ID_NEW / "input" / "a.xlsx")
            (root / ah.LEDGER_NAME).write_text(
                json.dumps([{"file": "a.xlsx", "hash": digest}], ensure_ascii=False), encoding="utf-8"
            )
            ah.backfill_run_ids(root)
            self.assertEqual(ah.load_ledger(root / ah.LEDGER_NAME)[0]["run_id"], RUN_ID_NEW)


class RetentionTests(unittest.TestCase):
    """90 天保留策略：只清理过期 run 目录，并把台账标记为已归档。"""

    def _seed_run(self, root: Path, run_id: str, ledger_entries=None) -> None:
        run = root / run_id
        (run / "input").mkdir(parents=True, exist_ok=True)
        (run / "input" / "样本.xlsx").write_bytes(b"payload")
        if ledger_entries is not None:
            (root / ah.LEDGER_NAME).write_text(
                json.dumps(ledger_entries, ensure_ascii=False), encoding="utf-8"
            )

    def test_expired_run_is_purged_and_marked_archived(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed_run(root, RUN_ID_OLD, [{"file": "样本.xlsx", "run_id": RUN_ID_OLD,
                                               "artifacts": ah.ARTIFACT_AVAILABLE}])
            purged = ah.purge_expired_artifacts(root, days=90)
            self.assertEqual(purged, [RUN_ID_OLD])
            self.assertFalse((root / RUN_ID_OLD).exists())
            entry = ah.load_ledger(root / ah.LEDGER_NAME)[0]
            self.assertEqual(entry["artifacts"], ah.ARTIFACT_ARCHIVED)
            self.assertEqual(entry["run_id"], RUN_ID_OLD, "台账记录本身必须保留")

    def test_recent_run_is_kept(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            recent = datetime.now().strftime("%Y%m%dT%H%M%S") + "_cccccccc"
            self._seed_run(root, recent, [{"file": "样本.xlsx", "run_id": recent,
                                           "artifacts": ah.ARTIFACT_AVAILABLE}])
            self.assertEqual(ah.purge_expired_artifacts(root, days=90), [])
            self.assertTrue((root / recent).exists())
            self.assertEqual(ah.load_ledger(root / ah.LEDGER_NAME)[0]["artifacts"],
                             ah.ARTIFACT_AVAILABLE)

    def test_non_run_directories_are_never_touched(self):
        """runs/ 下的手工目录、回归样本目录绝不能被误删。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            manual = root / "regress_align"
            manual.mkdir()
            (manual / "keep.txt").write_text("keep", encoding="utf-8")
            self._seed_run(root, RUN_ID_OLD, [{"file": "样本.xlsx", "run_id": RUN_ID_OLD}])
            purged = ah.purge_expired_artifacts(root, days=90)
            self.assertEqual(purged, [RUN_ID_OLD])
            self.assertTrue((manual / "keep.txt").is_file())

    def test_zero_days_disables_purge(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed_run(root, RUN_ID_OLD, [{"file": "样本.xlsx", "run_id": RUN_ID_OLD}])
            self.assertEqual(ah.purge_expired_artifacts(root, days=0), [])
            self.assertTrue((root / RUN_ID_OLD).exists())

    def test_retention_days_reads_env(self):
        name = ah.RETENTION_ENV
        original = os.environ.get(name)
        try:
            os.environ.pop(name, None)
            self.assertEqual(ah.retention_days(), ah.DEFAULT_RETENTION_DAYS)
            os.environ[name] = " 30 "
            self.assertEqual(ah.retention_days(), 30)
            os.environ[name] = "0"
            self.assertEqual(ah.retention_days(), 0)
            os.environ[name] = "abc"
            self.assertEqual(ah.retention_days(), ah.DEFAULT_RETENTION_DAYS)
        finally:
            if original is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = original

    def test_cutoff_boundary_uses_run_start_time(self):
        """用 run_id 里的开始时间判断，而不是目录 mtime。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._seed_run(root, RUN_ID_OLD, [])
            now = datetime(2026, 4, 1, 0, 0, 0)  # 距今 90 天（2026-01-01 起算）之内
            self.assertEqual(ah.purge_expired_artifacts(root, days=90, now=now), [])
            now = datetime(2026, 4, 2, 0, 0, 0)  # 刚好超过 90 天
            self.assertEqual(ah.purge_expired_artifacts(root, days=90, now=now), [RUN_ID_OLD])


class HistoryArtifactHttpTests(unittest.TestCase):
    """历史产物下载接口：未授权拒绝、非法编号拒绝、归档后 404、正常可下载。"""

    TOKEN = "artifact-test-token"

    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory()
        runs = Path(cls._tmp.name)
        run = runs / RUN_ID_NEW
        (run / "input").mkdir(parents=True)
        source = run / "input" / "样本.xlsx"
        source.write_bytes(b"xlsx-bytes")
        cls.html_name, cls.xlsx_name = output_artifact_names(source)
        (run / cls.html_name).write_bytes(b"<html>report</html>")
        (run / cls.xlsx_name).write_bytes(b"xlsx-workbook")
        cls._httpd = server.ThreadingHTTPServer(("127.0.0.1", 0), server.AuditHandler)
        cls._httpd.runs_dir = runs.resolve()
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

    def tearDown(self):
        if self._original is None:
            os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        else:
            os.environ[server.ADMIN_TOKEN_ENV] = self._original

    def _get(self, path: str, token: str | None = None) -> tuple[int, bytes]:
        request = urllib.request.Request(f"http://127.0.0.1:{self._port}{path}")
        if token is not None:
            request.add_header("Authorization", f"Bearer {token}")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, response.read()
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read()

    def test_anonymous_is_rejected(self):
        status, _ = self._get(f"/api/history/artifact?run_id={RUN_ID_NEW}&kind=html")
        self.assertEqual(status, 401)

    def test_wrong_token_is_rejected(self):
        status, _ = self._get(f"/api/history/artifact?run_id={RUN_ID_NEW}&kind=html", token="nope")
        self.assertEqual(status, 401)

    def test_invalid_run_id_is_rejected(self):
        status, _ = self._get(
            f"/api/history/artifact?run_id=../../etc&kind=html", token=self.TOKEN
        )
        self.assertEqual(status, 400)

    def test_traversal_with_valid_prefix_is_rejected(self):
        status, _ = self._get(
            f"/api/history/artifact?run_id={RUN_ID_NEW}/../../etc&kind=html", token=self.TOKEN
        )
        self.assertIn(status, (400, 404))

    def test_unsupported_kind_is_rejected(self):
        status, _ = self._get(
            f"/api/history/artifact?run_id={RUN_ID_NEW}&kind=pdf", token=self.TOKEN
        )
        self.assertEqual(status, 400)

    def test_missing_artifact_returns_404(self):
        status, body = self._get(
            "/api/history/artifact?run_id=20200101T000000_ffffffff&kind=html", token=self.TOKEN
        )
        self.assertEqual(status, 404)
        self.assertEqual(json.loads(body)["code"], "artifact_missing")

    def test_authorized_download_returns_file(self):
        status, body = self._get(
            f"/api/history/artifact?run_id={RUN_ID_NEW}&kind=html", token=self.TOKEN
        )
        self.assertEqual(status, 200)
        self.assertIn(b"report", body)
        status, body = self._get(
            f"/api/history/artifact?run_id={RUN_ID_NEW}&kind=xlsx", token=self.TOKEN
        )
        self.assertEqual(status, 200)
        self.assertEqual(body, b"xlsx-workbook")

    def test_public_audit_endpoint_still_works_without_token(self):
        """业务人员拿到的报告直链是公开的，不能因为本次改动被锁住。"""
        from urllib.parse import quote
        path = f"/api/audits/{RUN_ID_NEW}/{quote(self.html_name)}"
        status, body = self._get(path)
        self.assertEqual(status, 200)
        self.assertIn(b"report", body)


class ProtectedRouteRegistrationTests(unittest.TestCase):
    def test_history_artifact_endpoint_is_protected(self):
        self.assertIn("/api/history/artifact", server.PROTECTED_GET)


if __name__ == "__main__":
    unittest.main()
