# -*- coding: utf-8 -*-
"""「编码生成」Tab 的接口回归测试。

覆盖：
  · 启用 AI 需要管理令牌（防止误触发 20 分钟长任务）
  · 纯规则生成公开可用，且能跑通后台任务
  · 状态查询 / 产物下载 / 最近记录
  · 产物下载的目录穿越防护
"""
from __future__ import annotations

import io
import json
import os
import pathlib
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.parse
import urllib.request

import openpyxl

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import server  # noqa: E402

TOKEN = "gen-test-token"


def _make_xlsx(path: pathlib.Path) -> None:
    """造一个列名与标准不同的最小输入，验证模糊列识别也能跑。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KKS竣工版"
    ws.append(["序号", "图纸名称", "设计码", "设备名", "KKS码(最终)", "父级名称", "父级编码", "专业"])
    ws.append([1, "7号机组NCS测控屏", "07AAB02UH001", "1号机NCS测控屏", "07AAB02UH001",
               "1号机NCS测控屏系统", "07AAB02", "电气二次"])
    ws.append([2, "7号机远动采集控制屏", "07AAC01UH001", "1号机远动采集控制屏", "07AAC01UH001",
               "1号机远动采集控制屏系统", "07AAC01", "电气二次"])
    wb.save(path)


class GenerateApiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        runs = pathlib.Path(cls._tmp.name) / "runs"
        runs.mkdir(parents=True)
        cls.xlsx = pathlib.Path(cls._tmp.name) / "input.xlsx"
        _make_xlsx(cls.xlsx)
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
        os.environ[server.ADMIN_TOKEN_ENV] = TOKEN

    def tearDown(self):
        if self._original is None:
            os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        else:
            os.environ[server.ADMIN_TOKEN_ENV] = self._original

    # ---------- 工具 ----------
    def _post(self, use_ai: bool, token: str | None = None, filename: str = "input.xlsx"):
        data = self.xlsx.read_bytes()
        url = f"http://127.0.0.1:{self._port}/api/generate?use_ai={1 if use_ai else 0}"
        req = urllib.request.Request(
            url, data=data, method="POST",
            headers={"Content-Type": "application/octet-stream",
                     "X-Filename": urllib.parse.quote(filename)},
        )
        if token:
            req.add_header("Authorization", f"Bearer {token}")
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                return resp.status, json.loads(resp.read())
        except urllib.error.HTTPError as exc:
            raw = exc.read()
            try:
                return exc.code, json.loads(raw)
            except json.JSONDecodeError:
                return exc.code, {}

    def _get(self, path: str):
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{self._port}{path}", timeout=60) as resp:
                return resp.status, resp.read(), resp.headers.get("Content-Type", "")
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read(), ""

    def _wait(self, run_id: str, timeout: float = 180.0) -> dict:
        deadline = time.time() + timeout
        state: dict = {}
        while time.time() < deadline:
            _, raw, _ = self._get(f"/api/generate/{run_id}/status")
            state = json.loads(raw)
            if state.get("status") in ("completed", "error"):
                return state
            time.sleep(0.4)
        return state

    # ---------- 鉴权 ----------
    def test_ai_with_token_succeeds(self):
        """带令牌启用 AI 时，请求必须通过鉴权。

        🔴 实测事故：前端 `genStart` 用裸 fetch，**忘了带 Authorization 头**，
        于是服务端一律 401，页面只显示"未授权：管理令牌缺失或无效" ——
        看起来像令牌没配好，其实是请求根本没带令牌。
        （其他接口都走 jsonFetch，它会自动加头，只有这个上传漏了。）
        """
        status, body = self._post(use_ai=True, token=TOKEN)
        self.assertEqual(202, status, body)
        run_id = body.get("run_id")
        self.assertTrue(run_id)
        # 用 AI 会真调模型，测试里只确认任务已受理、不等待其跑完
        _, raw, _ = self._get(f"/api/generate/{run_id}/status")
        state = json.loads(raw)
        self.assertIn(state.get("status"), ("queued", "running", "completed"))

    def test_ai_requires_admin_token(self):
        """启用 AI 必须带令牌——否则会有人误触发 20 分钟长任务并消耗模型额度。"""
        status, body = self._post(use_ai=True)
        self.assertEqual(401, status)
        self.assertIn("未授权", body.get("error", ""))

    def test_auth_denial_reads_request_body_first(self):
        """🔴 401 必须发生在**读完请求体之后**。

        实测事故：勾了"用 AI 补名称"但浏览器没配令牌时，上传 ~1.9MB 表格，
        服务端在客户端还没上传完就断开 → 浏览器只报 `Failed to fetch`，
        用户完全看不到 401，以为服务坏了。
        这里用大文件复现：必须能正常收到 401 响应体，而不是连接被重置。
        """
        big = self.xlsx.read_bytes() + b"\x00" * (2 * 1024 * 1024)   # ~2MB
        url = f"http://127.0.0.1:{self._port}/api/generate?use_ai=1"
        req = urllib.request.Request(
            url, data=big, method="POST",
            headers={"Content-Type": "application/octet-stream",
                     "X-Filename": urllib.parse.quote("big.xlsx")},
        )
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                self.fail(f"应被拒绝，却返回 {resp.status}")
        except urllib.error.HTTPError as exc:
            self.assertEqual(401, exc.code)
            payload = json.loads(exc.read())          # 能读到响应体 = 连接没被重置
            self.assertIn("未授权", payload.get("error", ""))

    def test_oversize_upload_gets_readable_error(self):
        """> 上限的文件也要能读到明确的 413 文案，而不是 Failed to fetch。"""
        original = self._httpd.max_upload_bytes
        self._httpd.max_upload_bytes = 1024            # 临时压到 1KB
        try:
            status, body = self._post(use_ai=False)
            self.assertEqual(413, status)
            self.assertIn("限制", body.get("error", ""))
        finally:
            self._httpd.max_upload_bytes = original

    def test_ai_rejected_when_token_unset(self):
        os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        status, _ = self._post(use_ai=True, token=TOKEN)
        self.assertEqual(403, status)

    def test_rule_only_is_public(self):
        """纯规则生成公开可用（秒级完成，不消耗模型额度）。"""
        status, body = self._post(use_ai=False)
        self.assertEqual(202, status, body)
        self.assertTrue(body.get("run_id"))
        self.assertIn("status_url", body)

    # ---------- 后台任务与产物 ----------
    def test_generate_job_completes_and_exposes_artifacts(self):
        _, body = self._post(use_ai=False)
        run_id = body["run_id"]
        state = self._wait(run_id)
        self.assertEqual("completed", state.get("status"), state.get("hint"))
        self.assertGreater(state.get("nodes", 0), 0)
        self.assertGreater(state.get("data_rows", 0), 0)
        files = state.get("files") or {}
        # 2026-10-10 起：**只交付回填后的原表**（干净表 / 对照表 / JSON 不再作为下载项）
        self.assertEqual(2, len(files), f"应提供原表和待核清单，实际={list(files)}")
        self.assertIn("设备编码待核清单", files)
        self.assertTrue(any("已填写" in k for k in files), files)
        # 产物可下载且是 xlsx
        for key, url in files.items():
            code, blob, ctype = self._get(url)
            self.assertEqual(200, code, url)
            self.assertGreater(len(blob), 1000)
            self.assertIn("spreadsheet", ctype)

    def test_status_of_unknown_job_is_404(self):
        code, raw, _ = self._get("/api/generate/20260101T000000_deadbeef/status")
        self.assertEqual(404, code)
        self.assertIn("未找到", json.loads(raw).get("error", ""))

    def test_artifact_download_blocks_traversal(self):
        """产物下载只能取 runs/<run_id>/generated/ 下的文件。"""
        _, body = self._post(use_ai=False)
        run_id = body["run_id"]
        self._wait(run_id)
        for bad in ("../server.py", "..%2F..%2Fserver.py", "%2e%2e%2fserver.py"):
            code, _, _ = self._get(f"/api/generate/{run_id}/{bad}")
            self.assertEqual(404, code, f"{bad} 应被拒绝")

    def test_recent_lists_completed_generation(self):
        _, body = self._post(use_ai=False)
        run_id = body["run_id"]
        self._wait(run_id)
        code, raw, _ = self._get("/api/generate/recent")
        self.assertEqual(200, code)
        items = json.loads(raw).get("items", [])
        self.assertTrue(any(item["run_id"] == run_id for item in items), items)

    def test_non_xlsx_rejected(self):
        status, body = self._post(use_ai=False, filename="notes.txt")
        self.assertEqual(400, status)
        self.assertIn("xlsx", body.get("error", ""))


class GeneratePageMarkupTest(unittest.TestCase):
    """页面必须挂上导航、页面容器与关键脚本（防止前端漏挂）。"""

    def test_page_scripts_are_syntactically_valid(self):
        """页面内联脚本必须能通过语法检查。

        🔴 这是踩过的真事故：`deleteEntry` 里把 Python 的 `\\n` 写进了 JS 单引号字符串，
        求值后变成**真实换行** → 整个 <script> SyntaxError → **页面上所有按钮都失效**
        （连原有的上传、导航都不工作），表现却像"文件选择按钮点不动"。
        因此必须对**服务端实际返回的页面**做语法校验。
        """
        import subprocess

        result = subprocess.run(
            [sys.executable, "-B", str(PROJECT / "tools" / "verify_page_scripts.py")],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
            cwd=str(PROJECT),
        )
        self.assertEqual(0, result.returncode, f"页面脚本校验失败：\n{result.stdout}\n{result.stderr}")

    def test_no_literal_newline_inside_js_confirm_string(self):
        """回归：台账删除确认文案里的换行必须是 JS 转义序列，不能是真实换行。"""
        html = server.UPLOAD_PAGE
        self.assertIn("移除这条审核记录吗？\\n\\n文件：", html)
        self.assertIn("同时移除该次的报告 / 问题清单产物吗？\\n\\n", html)

    def test_nav_and_page_present(self):
        html = server.UPLOAD_PAGE
        self.assertIn('data-panel="generatePage"', html)
        self.assertIn('<section id="generatePage"', html)
        self.assertIn("generatePage:['KKS CODE GENERATION'", html)

    def test_script_hooks_present(self):
        html = server.UPLOAD_PAGE
        for marker in ("function genStart", "function genPoll", "function genDone",
                       "window.__initGeneratePage", "api/generate"):
            self.assertIn(marker, html, f"缺少前端逻辑：{marker}")

    def test_generate_upload_sends_admin_token(self):
        """上传生成请求必须带上 Authorization 头（否则启用 AI 永远 401）。"""
        html = server.UPLOAD_PAGE
        self.assertIn("genHeaders['Authorization']='Bearer '", html,
                      "genStart 没有附加管理令牌 → 启用 AI 会被 401 拒绝")

    def test_page_switching_calls_generate_init(self):
        html = server.UPLOAD_PAGE
        self.assertIn("page==='generatePage'&&window.__initGeneratePage", html)

    def test_deep_link_supported(self):
        self.assertIn("URLSearchParams(location.search).get('page')", server.UPLOAD_PAGE)


if __name__ == "__main__":
    unittest.main()
