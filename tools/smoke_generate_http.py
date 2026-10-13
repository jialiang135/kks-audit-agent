# -*- coding: utf-8 -*-
"""端到端冒烟：对**运行中的服务**走一遍「上传 → 生成 → 取交付物 → 自查」。

为什么必须有这个（而不是只跑 CLI）：
  CLI 跑通 **≠** 服务端跑通。服务端是在**同进程内**调用 `gen_step1.main()`，
  它的 stdout 是 **GBK**——曾有一个 `⚠️` 符号把整个生成任务打挂
  （"生成失败：'gbk' codec can't encode character '\\u26ad'"），CLI 却一切正常。
  所以每次动生成链路，都要真走一次 HTTP。

用法：
    python tools/smoke_generate_http.py <待上传.xlsx> [--no-ai] [--out outputs/smoke] [--timeout 1800]

令牌从 `.env.local` / 环境变量 `KKS_ADMIN_TOKEN` 读（**不回显**）。
AI 模式需要令牌；`--no-ai` 模式服务端公开，但本脚本仍会带令牌（有则不亏）。
"""
from __future__ import annotations

import argparse
import json
import urllib.parse
import os
import pathlib
import re
import sys
import time
import urllib.error
import urllib.request

ROOT = pathlib.Path(__file__).resolve().parents[1]


def _load_token() -> str:
    """读管理令牌：环境变量优先，其次 `.env.local`。**绝不打印**。"""
    tok = (os.environ.get("KKS_ADMIN_TOKEN") or "").strip()
    if tok:
        return tok
    for name in (".env.local", ".env"):
        p = ROOT / name
        if not p.exists():
            continue
        for line in p.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if k.strip() == "KKS_ADMIN_TOKEN":
                return v.strip().strip("'\"")
    return ""


def _upload_bytes(path: pathlib.Path) -> tuple[bytes, dict]:
    """按**前端的真实约定**构造上传：正文是裸字节，文件名放在 `X-Filename`（百分号编码）。

    ⚠️ 踩过：一开始这里用 multipart + `filename="…"`，服务端取的是 `X-Filename` 头
      （`server.handle_generate` 里 `raw_name = decode_filename_header(headers["X-Filename"])`），
      脚本没发这个头 → 文件名退化成 `input.xlsx`，产物叫 `input_已填写.xlsx`。
      所以冒烟必须**照抄前端** `fetch('/api/generate', {headers:{'Content-Type':'application/octet-stream',
      'X-Filename': encodeURIComponent(f.name)}})`，否则测的不是同一条路径。
    """
    headers = {
        "Content-Type": "application/octet-stream",
        "X-Filename": urllib.parse.quote(path.name),
    }
    return path.read_bytes(), headers


def _request(url: str, *, data: bytes | None = None, headers: dict | None = None,
             method: str = "GET", timeout: int = 60):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("--base", default="http://127.0.0.1:8000")
    ap.add_argument("--out", default="outputs/smoke")
    ap.add_argument("--no-ai", action="store_true")
    ap.add_argument("--timeout", type=int, default=1800)
    args = ap.parse_args()

    src = pathlib.Path(args.source)
    out_dir = ROOT / args.out
    out_dir.mkdir(parents=True, exist_ok=True)
    token = _load_token()
    auth = {"Authorization": f"Bearer {token}"} if token else {}

    status, body = _request(f"{args.base}/healthz")
    if status != 200:
        print(f"[失败] 服务没起来（/healthz -> {status}）：{body[:200]!r}")
        return 1
    print(f"[信息] 服务在线 {args.base}｜上传 {src.name}（{src.stat().st_size/1024:.0f} KB）"
          f"｜AI={'关' if args.no_ai else '开'}｜令牌={'有' if token else '无'}")

    # ---- 1) 上传并起任务（与前端同一条路径）----
    payload, up_headers = _upload_bytes(src)
    use_ai = "0" if args.no_ai else "1"
    status, body = _request(f"{args.base}/api/generate?use_ai={use_ai}", data=payload,
                            headers={**auth, **up_headers}, method="POST", timeout=300)
    if status not in (200, 202):
        print(f"[失败] 起任务被拒（{status}）：{body[:400].decode('utf-8', 'replace')}")
        return 1
    info = json.loads(body)
    run_id = info.get("run_id") or ""
    print(f"[信息] run_id={run_id}")

    # ---- 2) 轮询状态 ----
    deadline = time.time() + args.timeout
    last = ""
    while time.time() < deadline:
        status, body = _request(f"{args.base}/api/generate/{run_id}/status")
        if status != 200:
            print(f"[失败] 查状态失败（{status}）")
            return 1
        st = json.loads(body)
        line = f"{st.get('status')} {st.get('percent')}% {st.get('stage') or ''} {st.get('message') or ''}"
        if line != last:
            print(f"  {line}")
            last = line
        if st.get("status") in ("done", "completed", "failed", "interrupted"):
            break
        time.sleep(3)
    else:
        print("[失败] 超时未结束")
        return 1

    if st.get("status") not in ("done", "completed"):
        print(f"[失败] 任务未成功：{st.get('status')}｜{st.get('error') or st.get('message') or ''}")
        return 1

    files = st.get("files") or {}
    print(f"[信息] 产物 {len(files)} 项：{list(files)}")
    target = next((k for k in files if "已填写" in k), None)
    if not target:
        print(f"[失败] 没有「已填写原表」产物：{files}")
        return 1

    # ---- 3) 下载交付物 ----
    url = args.base + files[target]
    status, blob = _request(url, timeout=600)
    if status != 200:
        print(f"[失败] 下载失败（{status}）")
        return 1
    dest = out_dir / f"{src.stem}_已填写.xlsx"
    dest.write_bytes(blob)
    print(f"[信息] 已下载 {dest.relative_to(ROOT)}（{len(blob)/1024:.0f} KB）")

    # ---- 4) 交付物自查（三条纪律）----
    sys.path.insert(0, str(ROOT / "tools"))
    import verify_fill_back as V
    sys.argv = ["verify_fill_back", str(src), str(dest)]
    print("[信息] ── 交付物自查 ──")
    rc = V.main()
    print("[通过] 端到端冒烟成功" if rc == 0 else "[失败] 交付物自查未通过")
    return rc


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
