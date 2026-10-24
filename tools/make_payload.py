# -*- coding: utf-8 -*-
r"""组装配送目录 payload\ 并压出 payload.zip（安装器内嵌用）。

产出结构 = 安装后的目录结构（方案 3.2 节），但【只做显式白名单复制】，
真实学生数据 / API Key / 运行时目录天然进不来（见文件末尾的排除自检）。

前置：
  · 主程序已打包：_build_pkg/dist/德育分自动化系统/（exe + _internal）
  · 更新器已打包：_build_pkg/updater/dist/update.exe
  · MCP 已打包：  _build_pkg/mcp/dist/mcp_server.exe
  · 空白模板已生成：tools/make_blank_template.py

用法：python tools/make_payload.py
"""

import json
import os
import shutil
import sys
import zipfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(REPO_ROOT, "_build_pkg")
PAYLOAD = os.path.join(PKG, "payload")
APP_DIST = os.path.join(PKG, "dist", "德育分自动化系统")
UPDATE_EXE = os.path.join(PKG, "updater", "dist", "update.exe")
MCP_EXE = os.path.join(PKG, "mcp", "dist", "mcp_server.exe")

MANIFEST_DEFAULT = ("https://raw.githubusercontent.com/xiaohaifale-QWQ/"
                    "moral-edu-automation/main/version.json")
MIRROR_DEFAULT = ("https://cdn.jsdelivr.net/gh/xiaohaifale-QWQ/"
                  "moral-edu-automation@main/version.json")

# 绝不允许出现在 payload 里的东西（事故清单，方案 3.2 节）
FORBIDDEN = ("deepseek_config.json", "演示视频.mp4", "GIT仓库.zip",
             "识别记录日志.txt", "程序介绍.html", "德育分自动汇总系统介绍.pptx")


def need(path, what):
    if not os.path.exists(path):
        print(f"[错误] 缺少 {what}：{path}")
        print("      先跑：主程序 spec 打包 / src 两个小 spec 打包 / "
              "tools/make_blank_template.py")
        sys.exit(1)


def main():
    need(APP_DIST, "主程序打包产物")
    need(UPDATE_EXE, "更新器 update.exe")
    need(MCP_EXE, "mcp_server.exe")
    blank = os.path.join(PAYLOAD, "名单模板.xlsx")
    need(blank, "空白名单模板")

    # 清空重建 payload（名单模板先挪出来，避免误删后还要重新生成）
    tmp_keep = PAYLOAD + ".keep"
    shutil.move(blank, tmp_keep)
    shutil.rmtree(PAYLOAD, ignore_errors=True)
    os.makedirs(PAYLOAD)
    shutil.move(tmp_keep, blank)

    def cp(src, dst):
        dst = os.path.join(PAYLOAD, dst)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)

    # 主程序（exe + _internal 整目录）
    cp(os.path.join(APP_DIST, "德育分自动化系统.exe"), "德育分自动化系统.exe")
    shutil.copytree(os.path.join(APP_DIST, "_internal"),
                    os.path.join(PAYLOAD, "_internal"), dirs_exist_ok=True)
    # 更新器 / MCP
    cp(UPDATE_EXE, "update.exe")
    cp(MCP_EXE, os.path.join("AI对接", "mcp_server.exe"))
    # 文档（存在才带，缺了不致命）
    for rel, dst in [
        (os.path.join("AI对接", "使用说明.md"), os.path.join("AI对接", "使用说明.md")),
        (os.path.join("AI对接", "小欧-AI设定.md"), os.path.join("AI对接", "小欧-AI设定.md")),
        (os.path.join("软件说明", "使用说明书.md"), os.path.join("软件说明", "使用说明书.md")),
        (os.path.join("软件说明", "技术手册.md"), os.path.join("软件说明", "技术手册.md")),
        (os.path.join("软件说明", "CHANGELOG.md"), os.path.join("软件说明", "CHANGELOG.md")),
        (os.path.join("软件说明", "doro.ico"), os.path.join("软件说明", "doro.ico")),
    ]:
        src = os.path.join(REPO_ROOT, rel)
        if os.path.exists(src):
            cp(src, dst)
        else:
            print(f"[提示] 文档缺失（跳过）：{rel}")
    # 默认更新配置（用户可改）
    with open(os.path.join(PAYLOAD, "更新配置.json"), "w", encoding="utf-8") as f:
        json.dump({"manifest_url": MANIFEST_DEFAULT, "mirror_url": MIRROR_DEFAULT,
                   "last_check": 0, "skipped_version": ""},
                  f, ensure_ascii=False, indent=2)

    # 排除自检
    bad = []
    for root, _dirs, files in os.walk(PAYLOAD):
        for name in files:
            if name in FORBIDDEN or name.startswith("~$"):
                bad.append(os.path.relpath(os.path.join(root, name), PAYLOAD))
    if bad:
        print(f"[错误] payload 里混入了禁止分发的文件：{bad}")
        sys.exit(1)

    # 压 payload.zip（安装器内嵌）
    zp = os.path.join(PKG, "payload.zip")
    if os.path.exists(zp):
        os.remove(zp)
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for root, _dirs, files in os.walk(PAYLOAD):
            for name in files:
                full = os.path.join(root, name)
                zf.write(full, os.path.relpath(full, PAYLOAD))

    total = sum(os.path.getsize(os.path.join(r, n))
                for r, _d, fs in os.walk(PAYLOAD) for n in fs)
    print(f"[OK] payload 组装完成：{PAYLOAD}")
    print(f"     文件数：{sum(len(fs) for _r, _d, fs in os.walk(PAYLOAD))}"
          f"　总大小：{total / 1048576:.1f} MB")
    print(f"[OK] 安装器内嵌包：{zp}（{os.path.getsize(zp) / 1048576:.1f} MB）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
