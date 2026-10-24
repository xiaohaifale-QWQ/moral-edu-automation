# -*- coding: utf-8 -*-
r"""德育分自动化系统 · 更新器（打包成 update.exe，单文件、无窗口）。

由主程序「检查更新 → 立即更新」调用：
    update.exe --job <update.json>
也可手动：
    update.exe --zip <包.zip> --dir <安装目录> [--pid 主程序pid]
               [--sha256 <摘要>] [--version <新版本号>] [--no-restart] [--silent]

流程（方案 4.4 节）：
  1. 互斥体防重复运行
  2. 若自身在安装目录里（打包模式），先复制到 %TEMP% 再执行，才能替换自己
  3. 等主程序退出（30 秒），超时强杀
  4. 校验 zip 的 sha256 与完整性
  5. 解压到 <安装目录>\.update-tmp
  6. 备份：_internal 整目录、其余逐文件 → <安装目录>\.backup\<旧版本>\
  7. 覆盖（用户数据白名单绝不覆盖，见 is_protected）
  8. 任一步失败 → 从备份回滚 → 报「已恢复原版本」
  9. 成功：更新 安装信息.json，重启主程序

日志：%LOCALAPPDATA%\德育分自动化系统\update.log
"""

import argparse
import ctypes
import datetime
import fnmatch
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

APP_NAME = "德育分自动化系统"
MUTEX_NAME = "Global\\deyufen_update"
LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), APP_NAME)
LOG_FILE = os.path.join(LOG_DIR, "update.log")

# 用户数据白名单：更新包里就算带了这些名字，也绝不覆盖（方案 4.5 节）
PROTECTED_FILES = {"名单模板.xlsx", "更新配置.json", "安装信息.json", "deepseek_config.json"}
PROTECTED_PATTERNS = ("名单模板.自动备份-*.xlsx", "*.log")
PROTECTED_DIR_RES = (re.compile(r"^20\d{2}年\d{1,2}月(/|$)"),   # 月度数据文件夹
                     re.compile(r"^_ai_bridge(/|$)"))            # AI 队列运行时

# 纯程序运行时目录：整目录替换（先搬走旧的再放新的），避免陈旧 DLL 残留
WIPE_DIRS = ("_internal",)

QUIET = [False]


def log(message):
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    if not getattr(sys, "frozen", False):
        print(line)


def alert(message):
    if QUIET[0]:
        return log(f"[弹窗被 --silent 吞掉] {message}")
    try:
        ctypes.windll.user32.MessageBoxW(0, message, APP_NAME + " · 更新器", 0x40)
    except Exception:
        pass


def is_protected(rel):
    rel = rel.replace("\\", "/")
    name = rel.rsplit("/", 1)[-1]
    if name in PROTECTED_FILES:
        return True
    for pat in PROTECTED_PATTERNS:
        if fnmatch.fnmatch(name, pat):
            return True
    for rx in PROTECTED_DIR_RES:
        if rx.search(rel):
            return True
    return False


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def wait_for_exit(pid, timeout=30):
    """等主程序退出；超时先礼后兵（TerminateProcess）。进程已不在返回 True。"""
    if not pid:
        return True
    SYNCHRONIZE, PROCESS_TERMINATE = 0x00100000, 0x0001
    k = ctypes.windll.kernel32
    h = k.OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE, False, int(pid))
    if not h:
        return True
    try:
        if k.WaitForSingleObject(h, timeout * 1000) == 0:
            return True
        log(f"等待进程 {pid} {timeout}s 超时，强制结束")
        k.TerminateProcess(h, 1)
        return k.WaitForSingleObject(h, 5000) == 0
    finally:
        k.CloseHandle(h)


def read_installed_version(install_dir):
    try:
        with open(os.path.join(install_dir, "安装信息.json"), encoding="utf-8") as f:
            return str(json.load(f).get("version") or "未知")
    except Exception:
        return "未知"


def recover_residue(install_dir):
    """上次更新中断电/被杀留下的现场：.update-tmp 还在 → 从最新备份恢复。

    正常失败路径当场就回滚了；这里兜的是「备份搬完、还没来得及应用」
    就断电的窗口期——那时安装目录缺文件，备份里有原件。
    """
    staging = os.path.join(install_dir, ".update-tmp")
    backup_root = os.path.join(install_dir, ".backup")
    if not os.path.isdir(staging):
        return
    snaps = sorted(os.listdir(backup_root)) if os.path.isdir(backup_root) else []
    if snaps:
        bdir = os.path.join(backup_root, snaps[-1])
        log(f"发现上次更新中断的残留，先从备份 {os.path.basename(bdir)} 恢复")
        try:
            for name in os.listdir(bdir):
                src = os.path.join(bdir, name)
                dst = os.path.join(install_dir, name)
                if os.path.isdir(src):
                    if os.path.isdir(dst):
                        shutil.rmtree(dst, ignore_errors=True)
                    elif os.path.lexists(dst):
                        os.remove(dst)
                    shutil.move(src, dst)
                else:
                    if os.path.lexists(dst):
                        os.remove(dst)
                    os.makedirs(os.path.dirname(dst) or install_dir, exist_ok=True)
                    shutil.move(src, dst)
        except Exception as exc:
            log(f"残留恢复不完整：{exc}（继续正常更新流程）")
        shutil.rmtree(bdir, ignore_errors=True)
    shutil.rmtree(staging, ignore_errors=True)


def do_update(zip_path, install_dir, pid, sha_expect, new_ver, restart, no_restart):
    log(f"=== 更新开始：{os.path.basename(zip_path)} → {install_dir}"
        f"（当前 v{read_installed_version(install_dir)}）===")

    recover_residue(install_dir)

    if pid and not wait_for_exit(pid, 30):
        alert("主程序迟迟没有退出，更新已中止。\n请手动关闭程序后重试。")
        return 1
    if not os.path.exists(zip_path):
        alert(f"找不到更新包：\n{zip_path}")
        return 1
    if sha_expect:
        digest = sha256_file(zip_path)
        if digest != sha_expect:
            alert(f"sha256 校验失败，拒绝安装。\n期望 {sha_expect}\n实际 {digest}")
            return 1
        log("sha256 校验通过")

    staging = os.path.join(install_dir, ".update-tmp")
    backup_root = os.path.join(install_dir, ".backup")
    old_ver = read_installed_version(install_dir)
    backup_dir = os.path.join(backup_root, f"{old_ver}-{time.strftime('%H%M%S')}")

    wipe_moved = []      # (目录名, 备份位置)  整目录搬走的
    file_moved = []      # (相对路径, 备份位置) 逐文件搬走的

    try:
        shutil.rmtree(staging, ignore_errors=True)
        os.makedirs(staging)
        with zipfile.ZipFile(zip_path) as zf:
            bad = zf.testzip()
            if bad:
                raise RuntimeError(f"压缩包损坏：{bad}")
            zf.extractall(staging)
            names = [n.replace("\\", "/") for n in zf.namelist() if not n.endswith("/")]

        skipped = [n for n in names if is_protected(n)]
        for n in skipped:
            log(f"跳过受保护文件（用户数据）：{n}")
        entries = [n for n in names if not is_protected(n)]
        if not entries:
            raise RuntimeError("更新包里没有可安装的程序文件")
        if os.path.commonpath([os.path.abspath(staging), os.path.abspath(install_dir)]) \
                != os.path.abspath(install_dir):
            raise RuntimeError("staging 目录异常")

        os.makedirs(backup_dir, exist_ok=True)
        in_wipe = lambda rel: any(rel == w or rel.startswith(w + "/") for w in WIPE_DIRS)

        # 1) 备份
        for wd in WIPE_DIRS:
            src = os.path.join(install_dir, wd)
            if any(n.startswith(wd + "/") for n in entries) and os.path.isdir(src):
                dst = os.path.join(backup_dir, wd)
                shutil.move(src, dst)
                wipe_moved.append((wd, dst))
                log(f"备份目录 {wd}\\ → .backup")
        for rel in entries:
            if in_wipe(rel):
                continue
            src = os.path.join(install_dir, rel)
            if os.path.lexists(src):
                dst = os.path.join(backup_dir, rel)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.move(src, dst)
                file_moved.append((rel, dst))

        # 2) 应用（_internal 整目录搬运，不依赖备份阶段是否搬过）
        done = 0
        for rel in entries:
            if in_wipe(rel):
                continue
            src = os.path.join(staging, rel)
            dst = os.path.join(install_dir, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            os.replace(src, dst)
            done += 1
        for wd in WIPE_DIRS:
            if not any(n.startswith(wd + "/") for n in entries):
                continue
            src = os.path.join(staging, wd)
            if not os.path.isdir(src):
                continue
            dst = os.path.join(install_dir, wd)
            if os.path.isdir(dst):                 # 没备份过（如旧目录是后出现的）→ 补备份
                bdir = os.path.join(backup_dir, wd)
                if not os.path.exists(bdir):
                    shutil.move(dst, bdir)
                    wipe_moved.append((wd, bdir))
                else:
                    shutil.rmtree(dst, ignore_errors=True)
            elif os.path.lexists(dst):             # 同名文件挡路 → 让开
                os.remove(dst)
            shutil.move(src, dst)
            done += 1
        log(f"已替换 {done} 个条目（包内 {len(entries)}，备份 {len(file_moved)} 文件"
            f"+{len(wipe_moved)} 目录）")

    except Exception as exc:
        log(f"更新失败：{exc}；开始回滚")
        ok = True
        for rel, bpath in reversed(file_moved):
            try:
                dst = os.path.join(install_dir, rel)
                if os.path.lexists(dst):
                    os.remove(dst)
                os.makedirs(os.path.dirname(dst) or install_dir, exist_ok=True)
                shutil.move(bpath, dst)
            except Exception as exc2:
                ok = False
                log(f"回滚文件 {rel} 失败：{exc2}")
        for wd, bpath in reversed(wipe_moved):
            try:
                dst = os.path.join(install_dir, wd)
                if os.path.lexists(dst):
                    shutil.rmtree(dst, ignore_errors=True)
                shutil.move(bpath, dst)
            except Exception as exc2:
                ok = False
                log(f"回滚目录 {wd} 失败：{exc2}")
        alert("更新失败，已恢复到原版本。" if ok else
              "更新失败，且部分文件未能自动恢复！\n请把安装目录下 .backup 里的内容手动拷回。")
        return 1

    # 3) 收尾
    shutil.rmtree(staging, ignore_errors=True)
    try:
        if os.path.isdir(backup_root):
            for old in sorted(os.listdir(backup_root))[:-1]:
                shutil.rmtree(os.path.join(backup_root, old), ignore_errors=True)
    except Exception as exc:
        log(f"清理旧备份失败（不影响结果）：{exc}")

    if not new_ver:
        m = re.search(r"(\d+\.){3}\d+|(\d+\.){2}\d+", os.path.basename(zip_path))
        new_ver = m.group(0) if m else ""
    info = {"version": new_ver or old_ver,
            "updated_at": datetime.datetime.now().isoformat(timespec="seconds")}
    try:
        with open(os.path.join(install_dir, "安装信息.json"), "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False, indent=2)
    except Exception as exc:
        log(f"写 安装信息.json 失败：{exc}")

    log(f"更新完成：v{old_ver} → v{info['version']}")
    if no_restart:
        return 0
    if restart and os.path.exists(restart):
        try:
            subprocess.Popen([restart], cwd=install_dir)
            log(f"已重启 {restart}")
        except Exception as exc:
            alert(f"更新成功，但自动重启失败：{exc}\n请手动打开程序。")
    else:
        alert(f"更新完成：v{info['version']}（未找到主程序，未自动重启）")
    return 0


def main():
    ap = argparse.ArgumentParser(description="德育分自动化系统更新器")
    ap.add_argument("--job", help="主程序写出的任务文件 update.json")
    ap.add_argument("--zip", dest="zip_path")
    ap.add_argument("--dir", dest="install_dir")
    ap.add_argument("--pid", type=int)
    ap.add_argument("--sha256")
    ap.add_argument("--version")
    ap.add_argument("--restart")
    ap.add_argument("--no-restart", action="store_true")
    ap.add_argument("--silent", action="store_true", help="不弹任何窗口（测试/静默用）")
    ap.add_argument("--relocated", action="store_true", help="内部参数：已在临时副本中运行")
    args = ap.parse_args()
    if args.silent:
        QUIET[0] = True

    k = ctypes.windll.kernel32
    k.CreateMutexW(None, False, MUTEX_NAME)
    if k.GetLastError() == 183:          # ERROR_ALREADY_EXISTS
        log("已有更新器在运行，本次退出")
        return 0

    job = {}
    if args.job:
        try:
            with open(args.job, encoding="utf-8") as f:
                job = json.load(f)
        except Exception as exc:
            log(f"任务文件读取失败：{exc}")
            alert(f"更新任务文件读取失败：\n{args.job}\n\n{exc}")
            return 1

    zip_path = args.zip_path or job.get("zip")
    install_dir = args.install_dir or job.get("dir")
    pid = args.pid or job.get("pid")
    sha_expect = str(args.sha256 or job.get("sha256") or "").lower()
    new_ver = args.version or job.get("version") or ""
    restart = args.restart or job.get("restart") or \
        (os.path.join(install_dir, APP_NAME + ".exe") if install_dir else None)

    if not zip_path or not install_dir:
        alert(f"参数不完整（--zip / --dir 必填）。")
        return 1
    if not os.path.isdir(install_dir):
        alert(f"安装目录不存在：\n{install_dir}")
        return 1

    # 打包模式下若自己就住在安装目录里，先复制到 %TEMP% 运行，才能替换 update.exe
    if getattr(sys, "frozen", False) and not args.relocated:
        me = os.path.abspath(sys.executable)
        if os.path.dirname(me) == os.path.abspath(install_dir):
            run_dir = os.path.join(tempfile.gettempdir(), "deyufen-update")
            os.makedirs(run_dir, exist_ok=True)
            copy = os.path.join(run_dir, "update-run.exe")
            try:
                shutil.copy2(me, copy)
                subprocess.Popen([copy] + sys.argv[1:] + ["--relocated"])
                log(f"自身在安装目录内，已搬迁到 {copy} 继续")
                return 0
            except Exception as exc:
                log(f"搬迁失败，原地继续：{exc}")

    return do_update(zip_path, install_dir, pid, sha_expect, new_ver,
                     restart, args.no_restart)


if __name__ == "__main__":
    sys.exit(main())
