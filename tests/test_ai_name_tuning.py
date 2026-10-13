# -*- coding: utf-8 -*-
"""AI 补名称的**调优参数**：默认值 + 环境变量覆盖 + 接线（2026-10-10 实测后定档）。

背景（都是实测数字，工具 `tools/probe_ai_endpoint.py`）：
  · 小请求（16 token）串行 0.4~0.9s，**8 并发也一样** → 端点不按请求数限流
  · 真实批量（40 行 / 请求体 39 KB）**空闲时串行 28~35s**
  · 生成时并发 8 → 20 批里 **5 次撞 90s 超时**；而 AI 阶段总耗时 525s ≈ 20×30s 串行
  ⇒ **并发给不了加速，只换来超时**，故默认并发 8 → **3**、超时 90 → **180**。

以前这两个值写死在代码里，调参要改代码；现在可以在 `.env.local` 里改。
本文件锁死三件事：默认值、覆盖优先级、**接线**（环境变量真的作用到线程池）。
"""
from __future__ import annotations

import os
import pathlib
import sys
import unittest
from unittest import mock

PROJECT = pathlib.Path(__file__).resolve().parents[1]
for p in (str(PROJECT), str(PROJECT / "kks_gen")):
    if p not in sys.path:
        sys.path.insert(0, p)

import ai_names  # noqa: E402
import ai_review  # noqa: E402

ENV_KEYS = ("AI_NAME_CONCURRENCY", "AI_NAME_BATCH_SIZE")


class DefaultsTest(unittest.TestCase):
    def test_concurrency_default_is_three(self):
        """默认并发 3（实测：该端点对大请求给不了并行度，8 只会超时）。"""
        self.assertEqual(3, ai_names.MAX_CONCURRENCY)

    def test_batch_size_default_is_forty(self):
        """默认每批 40 行（请求体约 39 KB，单批约 30s）。"""
        self.assertEqual(40, ai_names.BATCH_SIZE)

    def test_timeout_default_is_180(self):
        """默认超时 180s（正常请求 30s 返回，超时值只决定"卡住时等多久"）。"""
        self.assertEqual(180, ai_review.DEFAULT_TIMEOUT_SECONDS)


class TuningPrecedenceTest(unittest.TestCase):
    """优先级：调用方显式传参 > 环境变量/.env.local > 代码默认值。"""

    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in ENV_KEYS}
        for k in ENV_KEYS:
            os.environ.pop(k, None)

    def tearDown(self):
        for k, v in self._saved.items():
            os.environ.pop(k, None)
            if v is not None:
                os.environ[k] = v

    def test_env_overrides_default(self):
        os.environ["AI_NAME_CONCURRENCY"] = "5"
        self.assertEqual(5, ai_names._tuning("AI_NAME_CONCURRENCY", None, 3, 1, 16))

    def test_explicit_argument_beats_env(self):
        os.environ["AI_NAME_CONCURRENCY"] = "5"
        self.assertEqual(2, ai_names._tuning("AI_NAME_CONCURRENCY", 2, 3, 1, 16))

    def test_out_of_range_is_clamped(self):
        os.environ["AI_NAME_CONCURRENCY"] = "999"
        self.assertEqual(16, ai_names._tuning("AI_NAME_CONCURRENCY", None, 3, 1, 16))
        os.environ["AI_NAME_CONCURRENCY"] = "0"
        self.assertEqual(1, ai_names._tuning("AI_NAME_CONCURRENCY", None, 3, 1, 16))

    def test_garbage_value_falls_back_to_default(self):
        """写错了（`abc`）不能把生成搞崩，退回默认值。"""
        os.environ["AI_NAME_CONCURRENCY"] = "abc"
        self.assertEqual(3, ai_names._tuning("AI_NAME_CONCURRENCY", None, 3, 1, 16))


class WiringTest(unittest.TestCase):
    """🔴 接线：环境变量必须**真的**作用到线程池与分批，不能只是文档里写着。"""

    def setUp(self):
        self._saved = os.environ.get("AI_NAME_CONCURRENCY")
        os.environ["AI_NAME_CONCURRENCY"] = "2"

    def tearDown(self):
        os.environ.pop("AI_NAME_CONCURRENCY", None)
        if self._saved is not None:
            os.environ["AI_NAME_CONCURRENCY"] = self._saved

    def test_env_concurrency_reaches_thread_pool(self):
        captured: dict = {}
        real_executor = ai_names.ThreadPoolExecutor

        def spy(max_workers=None, **kw):            # noqa: ANN001, ANN003
            captured["max_workers"] = max_workers
            return real_executor(max_workers=max_workers, **kw)

        class Cfg:
            enabled = True
            api_key = "test-key"
            model = "test-model"
            timeout_seconds = 5

        class Client:
            def __init__(self, _cfg): pass
            def choose_model(self): return "test-model", "env"
            def review(self, _model, _prompt, _payload): return {"items": []}

        rows = [{"row": i, "kks": f"01ABC01QA00{i}", "original_name": "",
                 "parent_name": "1号机x", "parent_kks": "01ABC01",
                 "profession": "", "drawing": ""} for i in (1, 2)]

        with mock.patch.object(ai_names, "load_config", lambda: Cfg()), \
             mock.patch.object(ai_names, "OpenAICompatibleClient", Client), \
             mock.patch.object(ai_names, "ThreadPoolExecutor", spy):
            ai_names.resolve_names(rows, batch_size=1)   # 2 行 / 每批 1 行 → 2 批

        # 并发受环境变量控制（2），而不是旧的写死 8
        self.assertEqual(2, captured.get("max_workers"))


if __name__ == "__main__":
    unittest.main()
