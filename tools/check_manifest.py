# -*- coding: utf-8 -*-
"""发布自检：校验 version.json 与 Release 资产是否可达、哈希是否一致。

用法：
    python tools/check_manifest.py               # 检查仓库根 version.json
    python tools/check_manifest.py <url|文件>    # 检查指定清单
"""

import hashlib
import json
import os
import sys
import urllib.request

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_LOCAL = os.path.join(REPO_ROOT, "version.json")
DEFAULT_REMOTE = ("https://raw.githubusercontent.com/xiaohaifale-QWQ/"
                  "moral-edu-automation/main/version.json")
MIRROR_REMOTE = ("https://cdn.jsdelivr.net/gh/xiaohaifale-QWQ/"
                 "moral-edu-automation@main/version.json")


def fetch_json(url, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": "deyufen-check"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fetch_bin(url, timeout=30):
    h = hashlib.sha256()
    size = 0
    req = urllib.request.Request(url, headers={"User-Agent": "deyufen-check"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        while True:
            chunk = resp.read(1 << 20)
            if not chunk:
                break
            h.update(chunk)
            size += len(chunk)
    return h.hexdigest(), size


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_LOCAL
    if os.path.exists(target):
        with open(target, encoding="utf-8") as f:
            manifest = json.load(f)
        print(f"清单来源：本地文件 {target}")
    else:
        manifest = fetch_json(target)
        print(f"清单来源：网络 {target}")

    ok = True
    ver = manifest.get("version")
    print(f"版本：v{ver}　min_supported：v{manifest.get('min_supported')}")
    if not ver:
        print("[FAIL] version 缺失")
        return 1

    win = (manifest.get("download") or {}).get("win64") or {}
    for key, label in (("url", "主地址"), ("mirror", "镜像")):
        url = win.get(key)
        if not url:
            print(f"[提示] {label}：未配置")
            continue
        try:
            digest, size = fetch_bin(url)
        except Exception as exc:
            print(f"[FAIL] {label} 不可达：{url}\n      {exc}")
            ok = False
            continue
        sha_ok = (not win.get("sha256")) or digest == win["sha256"]
        size_ok = (not win.get("size")) or size == win["size"]
        size_txt = "一致" if size_ok else "%d ≠ %s" % (size, win.get("size"))
        sha_txt = "一致" if sha_ok else "不一致！"
        print(f"[{'OK ' if sha_ok and size_ok else 'FAIL'}] {label}："
              f"size={size_txt} sha256={sha_txt}")
        ok = ok and sha_ok and size_ok

    # 主清单线上是否已同步（提醒：push 了吗？）
    if os.path.exists(target):
        for url, label in ((DEFAULT_REMOTE, "raw"), (MIRROR_REMOTE, "jsDelivr")):
            try:
                remote = fetch_json(url, timeout=8)
                same = str(remote.get("version")) == str(ver)
                print(f"[{'OK ' if same else 'WARN'}] 线上{label}清单版本："
                      f"v{remote.get('version')}"
                      + ("" if same else f"（本地 v{ver}，记得 git push）"))
            except Exception as exc:
                print(f"[WARN] 线上{label}清单拉不到：{type(exc).__name__}")

    print("\n结论：" + ("清单与资产校验通过 ✔" if ok else "存在问题，别发布 ✘"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
