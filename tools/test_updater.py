# -*- coding: utf-8 -*-
"""更新器单测（不联网，本地假包）。对应方案 8.6 节：

  用例 A：正常更新 → 程序文件被替换、_internal 整目录换新、备份存在、
          用户数据（名单模板/更新配置/月度数据/_ai_bridge/log）原封不动
  用例 B：sha256 不符 → 拒绝安装，目录原样
  用例 C：包内夹带用户数据文件 → 被保护名单跳过，不覆盖
  用例 D：解压/应用失败 → 回滚，目录原样

跑法：python tools/test_updater.py
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
UPDATER = os.path.join(REPO_ROOT, "src", "update.py")

FAILS = []


def check(label, ok, extra=""):
    print(f"  [{'OK ' if ok else 'FAIL'}] {label}" + (f"  → {extra}" if extra else ""))
    if not ok:
        FAILS.append(label)


def write(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def make_install(dir_):
    """构造一个假的安装目录（含程序文件 + 各种用户数据）。"""
    exe = os.path.join(dir_, "德育分自动化系统.exe")
    write(exe, "OLD-MAIN-1.6.0")
    write(os.path.join(dir_, "update.exe"), "OLD-UPDATER")
    write(os.path.join(dir_, "_internal", "a.dll"), "OLD-A")
    write(os.path.join(dir_, "_internal", "sub", "b.dll"), "OLD-B")
    write(os.path.join(dir_, "_internal", "dead.dll"), "SHOULD-DISAPPEAR")
    write(os.path.join(dir_, "AI对接", "mcp_server.exe"), "OLD-MCP")
    write(os.path.join(dir_, "名单模板.xlsx"), "USER-ROSTER")
    write(os.path.join(dir_, "更新配置.json"), "USER-CONFIG")
    write(os.path.join(dir_, "安装信息.json"), json.dumps({"version": "1.6.0"},
                                                         ensure_ascii=False))
    write(os.path.join(dir_, "2026年10月", "周表.xlsx"), "USER-MONTH")
    write(os.path.join(dir_, "_ai_bridge", "inbox", "x.json"), "USER-BRIDGE")
    write(os.path.join(dir_, "update.log"), "USER-LOG")
    return exe


def make_zip(path, with_user_data=False):
    """构造新版本假包（结构与安装目录一致）。"""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("德育分自动化系统.exe", "NEW-MAIN-1.7.0")
        zf.writestr("update.exe", "NEW-UPDATER")
        zf.writestr("_internal/a.dll", "NEW-A")
        zf.writestr("_internal/sub/b.dll", "NEW-B")
        zf.writestr("_internal/new.dll", "NEW-C")
        zf.writestr("AI对接/mcp_server.exe", "NEW-MCP")
        zf.writestr("AI对接/使用说明.md", "doc")
        if with_user_data:                       # 恶意/失误夹带的用户数据
            zf.writestr("名单模板.xlsx", "EVIL-ROSTER")
            zf.writestr("更新配置.json", "EVIL-CONFIG")


def run_updater(zip_path, dir_, digest, extra=()):
    return subprocess.run(
        [PY, UPDATER, "--zip", zip_path, "--dir", dir_,
         "--sha256", digest, "--version", "1.7.0", "--no-restart", "--silent", *extra],
        capture_output=True, text=True, timeout=120)


def case_a_and_c():
    print("### A+C. 正常更新（包里还夹带了用户数据文件）")
    work = tempfile.mkdtemp(prefix="upd-test-")
    dir_ = os.path.join(work, "install")
    os.makedirs(dir_)
    make_install(dir_)
    zp = os.path.join(work, "new.zip")
    make_zip(zp, with_user_data=True)

    r = run_updater(zp, dir_, sha256(zp))
    check("退出码 0", r.returncode == 0, r.stdout[-300:] + r.stderr[-300:])

    rd = lambda p: open(os.path.join(dir_, p), encoding="utf-8").read()
    check("主程序已替换", rd("德育分自动化系统.exe") == "NEW-MAIN-1.7.0")
    check("更新器已替换", rd("update.exe") == "NEW-UPDATER")
    check("_internal/a.dll 已替换", rd("_internal/a.dll") == "NEW-A")
    check("_internal 新增 new.dll", rd("_internal/new.dll") == "NEW-C")
    check("_internal 旧残留 dead.dll 已清除（整目录替换）",
          not os.path.exists(os.path.join(dir_, "_internal", "dead.dll")))
    check("mcp_server.exe 已替换", rd("AI对接/mcp_server.exe") == "NEW-MCP")
    # 用户数据全部原封不动（夹带的 EVIL 版本没有覆盖它们）
    check("名单模板 未被覆盖", rd("名单模板.xlsx") == "USER-ROSTER")
    check("更新配置 未被覆盖", rd("更新配置.json") == "USER-CONFIG")
    check("月度数据 未被动", rd("2026年10月/周表.xlsx") == "USER-MONTH")
    check("_ai_bridge 未被动", rd("_ai_bridge/inbox/x.json") == "USER-BRIDGE")
    check("update.log 未被动", rd("update.log") == "USER-LOG")
    check(".update-tmp 已清理", not os.path.exists(os.path.join(dir_, ".update-tmp")))
    info = json.load(open(os.path.join(dir_, "安装信息.json"), encoding="utf-8"))
    check("安装信息.json 版本=1.7.0", info.get("version") == "1.7.0", str(info))
    backups = os.listdir(os.path.join(dir_, ".backup"))
    check("备份目录存在", len(backups) == 1, str(backups))
    bdir = os.path.join(dir_, ".backup", backups[0])
    check("备份里有旧主程序", open(os.path.join(bdir, "德育分自动化系统.exe"),
                                encoding="utf-8").read() == "OLD-MAIN-1.6.0")
    check("备份里有旧 _internal", open(os.path.join(bdir, "_internal", "a.dll"),
                                   encoding="utf-8").read() == "OLD-A")
    shutil.rmtree(work, ignore_errors=True)


def case_b():
    print("### B. sha256 不符 → 拒绝")
    work = tempfile.mkdtemp(prefix="upd-test-")
    dir_ = os.path.join(work, "install")
    os.makedirs(dir_)
    make_install(dir_)
    zp = os.path.join(work, "new.zip")
    make_zip(zp)
    r = run_updater(zp, dir_, "0" * 64)
    check("退出码非 0", r.returncode != 0)
    check("主程序未被动", open(os.path.join(dir_, "德育分自动化系统.exe"),
                           encoding="utf-8").read() == "OLD-MAIN-1.6.0")
    check("_internal 未被动", open(os.path.join(dir_, "_internal", "a.dll"),
                              encoding="utf-8").read() == "OLD-A")
    shutil.rmtree(work, ignore_errors=True)


def case_d():
    print("### D. 应用失败 → 回滚")
    work = tempfile.mkdtemp(prefix="upd-test-")
    dir_ = os.path.join(work, "install")
    os.makedirs(dir_)
    make_install(dir_)
    # 制造冲突：AI对接 是一个同名【文件】，应用 AI对接/mcp_server.exe 时 makedirs 失败
    shutil.rmtree(os.path.join(dir_, "AI对接"))
    write(os.path.join(dir_, "AI对接"), "IAM-A-FILE")
    zp = os.path.join(work, "new.zip")
    make_zip(zp)
    r = run_updater(zp, dir_, sha256(zp))
    check("退出码非 0", r.returncode != 0, r.stdout[-300:])
    check("主程序回滚为旧版", open(os.path.join(dir_, "德育分自动化系统.exe"),
                              encoding="utf-8").read() == "OLD-MAIN-1.6.0")
    check("update.exe 回滚为旧版", open(os.path.join(dir_, "update.exe"),
                                   encoding="utf-8").read() == "OLD-UPDATER")
    check("_internal 回滚为旧版", open(os.path.join(dir_, "_internal", "a.dll"),
                                 encoding="utf-8").read() == "OLD-A")
    check("用户数据未被动", open(os.path.join(dir_, "名单模板.xlsx"),
                            encoding="utf-8").read() == "USER-ROSTER")
    shutil.rmtree(work, ignore_errors=True)


def case_e():
    print("### E. 断电残留 → 下次运行先从备份恢复")
    work = tempfile.mkdtemp(prefix="upd-test-")
    dir_ = os.path.join(work, "install")
    os.makedirs(dir_)
    make_install(dir_)
    # 模拟「备份搬完 _internal、还没应用就断电」的现场
    os.makedirs(os.path.join(dir_, ".backup", "1.6.0-010101"))
    shutil.move(os.path.join(dir_, "_internal"),
                os.path.join(dir_, ".backup", "1.6.0-010101", "_internal"))
    write(os.path.join(dir_, ".update-tmp", "半截"), "CRASH-HERE")
    # 用一个 sha 不符的包触发运行：恢复必须发生在校验之前，且目录复原
    zp = os.path.join(work, "new.zip")
    make_zip(zp)
    r = run_updater(zp, dir_, "0" * 64)
    check("退出码非 0（sha 不符）", r.returncode != 0)
    check("_internal 已从备份恢复", open(os.path.join(dir_, "_internal", "a.dll"),
                                    encoding="utf-8").read() == "OLD-A")
    check(".update-tmp 已清掉", not os.path.exists(os.path.join(dir_, ".update-tmp")))
    check("主程序未被动", open(os.path.join(dir_, "德育分自动化系统.exe"),
                           encoding="utf-8").read() == "OLD-MAIN-1.6.0")
    shutil.rmtree(work, ignore_errors=True)


def main():
    case_a_and_c()
    case_b()
    case_d()
    case_e()
    if FAILS:
        print(f"\n[失败] {len(FAILS)} 项：{FAILS}")
        return 1
    print("\n[通过] 更新器全部用例通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
