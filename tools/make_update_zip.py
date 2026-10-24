# -*- coding: utf-8 -*-
"""打热更新包 release\\deyufen-<版本>-win64.zip，并回填 version.json。

  · zip 内部结构 = 安装目录结构（方案 9.3 节）
  · 名单模板.xlsx / 更新配置.json 不进更新包（用户数据，更新器也不许覆盖）
  · 打完自动把 sha256 / size / url 回填到仓库根的 version.json

用法：python tools/make_update_zip.py 1.7.0 [--notes "更新说明，\\n 换行"]
"""

import hashlib
import json
import os
import sys
import zipfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAYLOAD = os.path.join(REPO_ROOT, "_build_pkg", "payload")
RELEASE = os.path.join(REPO_ROOT, "release")
GH_BASE = "https://github.com/xiaohaifale-QWQ/moral-edu-automation/releases/download"
MIRROR_PREFIX = "https://ghfast.top/"

SKIP = {"名单模板.xlsx", "更新配置.json"}       # 用户数据不进更新包


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    if len(sys.argv) < 2 or not sys.argv[1][0].isdigit():
        print("用法：python tools/make_update_zip.py <版本号> [--notes \"...\"]")
        return 1
    version = sys.argv[1].strip()
    notes = ""
    if "--notes" in sys.argv:
        notes = sys.argv[sys.argv.index("--notes") + 1]

    if not os.path.isdir(PAYLOAD):
        print(f"[错误] payload 不存在：{PAYLOAD}\n先跑 tools/make_payload.py")
        return 1

    os.makedirs(RELEASE, exist_ok=True)
    zname = f"deyufen-{version}-win64.zip"
    zpath = os.path.join(RELEASE, zname)
    if os.path.exists(zpath):
        os.remove(zpath)
    count = 0
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for root, _dirs, files in os.walk(PAYLOAD):
            for name in files:
                if name in SKIP:
                    continue
                full = os.path.join(root, name)
                zf.write(full, os.path.relpath(full, PAYLOAD))
                count += 1

    digest = sha256_file(zpath)
    size = os.path.getsize(zpath)
    url = f"{GH_BASE}/v{version}/{zname}"

    # 回填 version.json
    vpath = os.path.join(REPO_ROOT, "version.json")
    manifest = {}
    if os.path.exists(vpath):
        try:
            with open(vpath, encoding="utf-8") as f:
                manifest = json.load(f)
        except Exception:
            manifest = {}
    manifest["version"] = version
    if notes:
        manifest["notes"] = notes
    manifest.setdefault("min_supported", "1.5.0")
    manifest["mandatory"] = bool(manifest.get("mandatory", False))
    manifest["download"] = {"win64": {
        "url": url,
        "mirror": MIRROR_PREFIX + url,
        "sha256": digest,
        "size": size,
    }}
    with open(vpath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"[OK] 更新包：{zpath}")
    print(f"     条目 {count} 个　大小 {size} 字节（{size / 1048576:.1f} MB）")
    print(f"     sha256：{digest}")
    print(f"     下载地址（上传 Release 后生效）：{url}")
    print(f"[OK] version.json 已回填：{vpath}")
    print("     发布步骤：GitHub → Releases → tag v%s → 上传 %s → Publish → git push" % (
        version, zname))
    return 0


if __name__ == "__main__":
    sys.exit(main())
