# -*- coding: utf-8 -*-
"""管理接口鉴权回归测试。

对应评审阻断项：配置读写 / Skill 读写与上传 / 运行日志 / AI 连接测试此前均无认证，
攻击者可先改 AI 地址、再触发 AI 测试，使服务端用已保存的 API Key 向其端点发请求，造成密钥外泄。
本测试锁定三点：令牌解析、环境变量读取、以及"哪些端点必须在保护名单内"。
"""
import base64
import os
import unittest

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
        self.assertEqual(server.PROTECTED_GET, {"/api/config", "/api/skill", "/api/logs"})

    def test_admin_post_endpoints_are_protected(self):
        for endpoint in ("/api/config", "/api/skill", "/api/skill/upload", "/api/ai/test"):
            self.assertIn(endpoint, server.PROTECTED_POST)

    def test_business_endpoints_stay_public(self):
        for endpoint in ("/healthz", "/api/audits", "/api/history", "/api/history/download"):
            self.assertNotIn(endpoint, server.PROTECTED_GET)
            self.assertNotIn(endpoint, server.PROTECTED_POST)


if __name__ == "__main__":
    unittest.main()
