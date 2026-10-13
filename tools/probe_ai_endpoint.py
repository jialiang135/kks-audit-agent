# -*- coding: utf-8 -*-
"""AI 端点**延迟/超时诊断**：用真实请求量出「为什么老是超时」。

只测，不改：读 `.env.local`（不回显密钥），发少量真实请求并计时。

用法：
    python tools/probe_ai_endpoint.py                 # 默认：小请求 ×3 串行
    python tools/probe_ai_endpoint.py --concurrency 8 # 8 并发（复现生成时的压力）
    python tools/probe_ai_endpoint.py --batch 40 --file "测试文件/xx.xlsx"   # 真实批量载荷

为什么要这个：超时成因至少三种，**处理方式完全不同**——
  ① 单批太大（模型要吐 30~50 KB，本身 >90s）→ 该**调大 AI_TIMEOUT_SECONDS**
  ② 并发太高被服务端排队/限流           → 该**调小并发**（调超时没用）
  ③ 链路本身慢/抖（代理、DNS、TLS）      → 与参数无关，只能重试或换端点
猜是猜不出来的，必须分头量。
"""
from __future__ import annotations

import argparse
import json
import pathlib
import statistics
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "kks_gen"))   # ai_names / tree_rules 都在 kks_gen 下


def _config():
    import ai_review
    ai_review._load_local_env()
    return ai_review.load_config()


def _timed_post(cfg, body: dict, timeout: float) -> tuple[float, int, str]:
    """返回 (耗时秒, 状态, 说明)。status=0 表示超时/连接失败。"""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        cfg.base_url.rstrip("/") + "/chat/completions",
        data=data, method="POST",
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {cfg.api_key}",
                 "User-Agent": "kks-audit-probe/1.0"},
    )
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
        return time.time() - t0, resp.status, f"{len(raw)} bytes"
    except urllib.error.HTTPError as exc:
        return time.time() - t0, exc.code, exc.read()[:120].decode("utf-8", "replace")
    except Exception as exc:                       # noqa: BLE001 - 探针要如实反映任何失败
        return time.time() - t0, 0, f"{type(exc).__name__}: {exc}"


def _small_body(cfg) -> dict:
    return {
        "model": cfg.model,
        "temperature": 0,
        "max_tokens": 16,
        "messages": [{"role": "user", "content": "只回复两个字：收到"}],
    }


def _batch_rows(path: pathlib.Path, limit: int) -> list[dict]:
    """从真实表格里取 limit 行**真实待补名称的行**（与生成时同样的载荷形状）。"""
    import gen_step1
    from tree_rules import build_tree
    seeds, _layout = gen_step1.load(path)
    nodes = build_tree(seeds)["nodes"]
    return gen_step1._collect_ai_rows(nodes)[:limit]


def _real_body(cfg, path: pathlib.Path, limit: int) -> tuple[dict, int]:
    import ai_names
    rows = _batch_rows(path, limit)
    payload = ai_names._build_payload(rows)
    body = {
        "model": cfg.model,
        "temperature": 0,
        "messages": [{"role": "system", "content": ai_names.NAME_PROMPT},
                     {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}],
    }
    size = len(json.dumps(body, ensure_ascii=False).encode("utf-8"))
    return body, size


def _stats(label: str, results: list[tuple[float, int, str]]) -> None:
    ok = [r for r in results if r[1] == 200]
    bad = [r for r in results if r[1] != 200]
    print(f"\n== {label}：成功 {len(ok)}/{len(results)}")
    if ok:
        times = sorted(r[0] for r in ok)
        print(f"   耗时 最快 {times[0]:.1f}s ｜ 中位 {statistics.median(times):.1f}s ｜ "
              f"最慢 {times[-1]:.1f}s")
    for t, st, msg in bad:
        print(f"   [失败] {t:.1f}s 状态={st} {msg}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="")
    ap.add_argument("--model", default="")
    ap.add_argument("--timeout", type=float, default=0, help="0 = 用配置里的 AI_TIMEOUT_SECONDS")
    ap.add_argument("--times", type=int, default=3)
    ap.add_argument("--concurrency", type=int, default=0, help="0 = 串行")
    ap.add_argument("--batch", type=int, default=0, help=">0 时发一批真实大小请求")
    ap.add_argument("--file", default="")
    args = ap.parse_args()

    cfg = _config()
    if args.base:
        cfg.base_url = args.base
    if args.model:
        cfg.model = args.model
    timeout = args.timeout or cfg.timeout_seconds

    print(f"端点   : {cfg.base_url}")
    print(f"模型   : {cfg.model}")
    print(f"超时   : {timeout:.0f}s（配置里的 AI_TIMEOUT_SECONDS={cfg.timeout_seconds}）")
    print(f"密钥   : {'已配置' if cfg.api_key else '缺失！'}（不回显）")

    # ---- A) 小请求：量"链路本身" ----
    body = _small_body(cfg)
    if args.concurrency and args.concurrency > 1:
        t0 = time.time()
        with ThreadPoolExecutor(max_workers=args.concurrency) as pool:
            results = list(pool.map(lambda _i: _timed_post(cfg, body, timeout),
                                    range(args.concurrency)))
        _stats(f"A) 小请求 ×{args.concurrency} 并发（总墙钟 {time.time()-t0:.1f}s）", results)
    else:
        _stats(f"A) 小请求 ×{args.times} 串行",
               [_timed_post(cfg, body, timeout) for _ in range(args.times)])

    # ---- B) 真实批量：量"模型吐这么多字要多久" ----
    if args.batch:
        path = pathlib.Path(args.file or "")
        if not path.is_file():
            print("\n[跳过] --batch 需要 --file 指向一个真实 .xlsx")
            return 0
        real_body, size = _real_body(cfg, path, args.batch)
        print(f"\nB) 真实批量请求：{args.batch} 行，请求体 {size/1024:.0f} KB")
        results = [_timed_post(cfg, real_body, timeout) for _ in range(args.times)]
        _stats(f"B) 真实批量 ×{args.times}", results)
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
