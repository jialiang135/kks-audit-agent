# -*- coding: utf-8 -*-
"""管理接口鉴权回归测试。

对应评审阻断项：配置读写 / Skill 读写与上传 / 运行日志 / AI 连接测试此前均无认证，
攻击者可先改 AI 地址、再触发 AI 测试，使服务端用已保存的 API Key 向其端点发请求，造成密钥外泄。
后续评审又指出台账接口（/api/history、/api/history/download）被测试固定为公开，
会泄露原始文件名、内容哈希与 P0/P1/P2 业务统计，因此一并纳入保护名单。
本测试锁定三点：令牌解析、环境变量读取、以及"哪些端点必须在保护名单内"。
"""
import base64
import json
import os
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path

import server


class ProvidedTokenTests(unittest.TestCase):
    """令牌提取：支持 Authorization: Bearer / Basic 与 ?token= 查询串三种方式。"""

    def test_bearer_header(self):
        self.assertEqual(server._provided_token({"Authorization": "Bearer abc"}, "/api/config"), "abc")

    def test_basic_header_uses_password_part(self):
        raw = base64.b64encode(b"admin:secret").decode()
        self.assertEqual(server._provided_token({"Authorization": f"Basic {raw}"}, "/api/config"), "secret")

    def test_query_token(self):
        self.assertEqual(server._provided_token({}, "/api/config?token=q123"), "q123")

    def test_missing_token(self):
        self.assertEqual(server._provided_token({}, "/api/config"), "")

    def test_malformed_basic_does_not_raise(self):
        self.assertEqual(server._provided_token({"Authorization": "Basic !!!"}, "/api/config"), "")


class AdminTokenEnvTests(unittest.TestCase):
    def test_reads_env_and_strips_whitespace(self):
        name = server.ADMIN_TOKEN_ENV
        original = os.environ.get(name)
        try:
            os.environ[name] = "  tok  "
            self.assertEqual(server.admin_token(), "tok")
            os.environ.pop(name, None)
            self.assertEqual(server.admin_token(), "")
        finally:
            if original is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = original


class ProtectedRouteTests(unittest.TestCase):
    """保护名单必须覆盖全部管理端点，防止以后新增路由漏挂鉴权。"""

    def test_admin_get_endpoints_are_protected(self):
        self.assertEqual(
            server.PROTECTED_GET,
            {
                "/api/config",
                "/api/skill",
                "/api/logs",
                "/api/history",
                "/api/history/download",
                "/api/history/artifact",
            },
        )

    def test_admin_post_endpoints_are_protected(self):
        self.assertEqual(
            server.PROTECTED_POST,
            {"/api/config", "/api/skill", "/api/skill/upload", "/api/ai/test", "/api/history/delete"},
        )

    def test_ledger_endpoints_require_admin_token(self):
        """台账返回原始文件名、内容哈希、审核时间与 P0/P1/P2 统计，不能因页面展示需要而公开。"""
        for endpoint in ("/api/history", "/api/history/download"):
            self.assertIn(endpoint, server.PROTECTED_GET)

    def test_non_ledger_readonly_endpoints_stay_public(self):
        """探活与单次审核产物（按 run_id 取）不含台账元数据，保持公开。"""
        for endpoint in ("/healthz", "/api/audits"):
            self.assertNotIn(endpoint, server.PROTECTED_GET)
            self.assertNotIn(endpoint, server.PROTECTED_POST)


class LedgerHttpRegressionTests(unittest.TestCase):
    """进程内起真实 HTTP 服务，按真实网络语义断言台账接口的鉴权行为。

    对应评审剩余风险项："未在隔离容器中执行未授权 HTTP 请求回归"。
    这里不依赖 Docker：用 ThreadingHTTPServer 绑 127.0.0.1 的随机端口，
    直接发 HTTP 请求，覆盖匿名 / 错误令牌 / 正确令牌 / 服务端未配令牌四种情况。
    """

    TOKEN = "test-token-0123456789"

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        runs = Path(cls._tmp.name)
        (runs / "audit_ledger.json").write_text("[]", encoding="utf-8")
        (runs / "KKS审核台账汇总.html").write_text("<html>ledger</html>", encoding="utf-8")
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

    def _request(self, path: str, token: str | None = None) -> tuple[int, bytes]:
        request = urllib.request.Request(f"http://127.0.0.1:{self._port}{path}")
        if token is not None:
            request.add_header("Authorization", f"Bearer {token}")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, response.read()
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read()

    def test_ledger_endpoints_reject_anonymous(self):
        for path in ("/api/history", "/api/history/download", "/api/history/download?type=ledger"):
            status, _ = self._request(path)
            self.assertEqual(status, 401, f"{path} 未携带令牌应返回 401，实际 {status}")

    def test_ledger_endpoints_reject_wrong_token(self):
        status, _ = self._request("/api/history", token="wrong-token")
        self.assertEqual(status, 401)

    def test_ledger_endpoints_accept_admin_token(self):
        status, body = self._request("/api/history", token=self.TOKEN)
        self.assertEqual(status, 200)
        self.assertIn("entries", json.loads(body))
        status, _ = self._request("/api/history/download", token=self.TOKEN)
        self.assertEqual(status, 200)
        status, _ = self._request("/api/history/download?type=ledger", token=self.TOKEN)
        self.assertEqual(status, 200)

    def test_ledger_endpoints_fail_closed_when_token_unset(self):
        """服务端没配 KKS_ADMIN_TOKEN 时，台账接口也不能裸奔。"""
        os.environ.pop(server.ADMIN_TOKEN_ENV, None)
        status, _ = self._request("/api/history", token=self.TOKEN)
        self.assertEqual(status, 403)

    def test_healthz_still_public(self):
        status, body = self._request("/healthz")
        self.assertEqual(status, 200)
        self.assertTrue(json.loads(body)["ok"])


if __name__ == "__main__":
    unittest.main()
