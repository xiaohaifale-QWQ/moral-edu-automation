# -*- coding: utf-8 -*-
r"""德育分自动化系统 · 安装程序（方案 3.4 节 B 路线，自写安装器）。

打包：installer/installer.spec 把本文件与 payload.zip 一起打成
      「德育分自动化系统-安装程序.exe」（onedir payload 内嵌，双击即用）。

功能：
  · 向导式 GUI：欢迎 → 选目录（默认 %LOCALAPPDATA%\Programs，免管理员）→ 进度 → 完成
  · 静默安装：--silent --dir=<路径>（热修复/脚本用）
  · 卸载：安装后生成 卸载.exe，运行它就是卸载（保留名单模板与月度数据）
  · 用户数据白名单：名单模板.xlsx / 更新配置.json / 20??年*月\ / *.log 等，
    已存在的一律不覆盖；程序文件按升级语义覆盖
  · 安装完生成 AI对接\MCP配置.json（绝对路径随安装目录）与 安装信息.json（版本）

日志：%TEMP%\deyufen-install.log
"""

import argparse
import ctypes
import datetime
import fnmatch
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "德育分自动化系统"
APP_VERSION = "1.8.3.1"
MUTEX_NAME = "Global\\deyufen_install"
REG_UNINSTALL = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\\" + APP_NAME
DEFAULT_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
                           "Programs", APP_NAME)
LOG_FILE = os.path.join(tempfile.gettempdir(), "deyufen-install.log")
MAIN_EXE = APP_NAME + ".exe"

# 用户数据白名单（升级/重装都不覆盖）
KEEP_FILES = {"名单模板.xlsx", "更新配置.json", "deepseek_config.json"}
KEEP_PATTERNS = ("名单模板.自动备份-*.xlsx", "*.log")
KEEP_DIR_RES = ("^20\\d{2}年\\d{1,2}月(/|$)", "^_ai_bridge(/|$)")
WIPE_DIRS = ("_internal",)          # 程序运行时目录整体替换

MUTED = "#6B7280"
ACCENT = "#0D9488"
INK = "#1F2937"


def log(message):
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    if not getattr(sys, "frozen", False):
        print(line)


def keep_user_data(rel):
    rel = rel.replace("\\", "/")
    name = rel.rsplit("/", 1)[-1]
    if name in KEEP_FILES:
        return True
    for pat in KEEP_PATTERNS:
        if fnmatch.fnmatch(name, pat):
            return True
    for pat in KEEP_DIR_RES:
        import re
        if re.search(pat, rel):
            return True
    return False


def payload_zip_path():
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "payload.zip")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "payload.zip")


def make_shortcut(lnk_path, target, workdir):
    """用 PowerShell 生成 .lnk（不引第三方依赖）。"""
    if not lnk_path.lower().endswith(".lnk"):
        lnk_path += ".lnk"
    ps = ("$w = New-Object -ComObject WScript.Shell; "
          f"$s = $w.CreateShortcut('{lnk_path}'); "
          f"$s.TargetPath = '{target}'; "
          f"$s.WorkingDirectory = '{workdir}'; "
          f"$s.IconLocation = '{target},0'; "
          "$s.Save()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", ps], capture_output=True, timeout=30)
        return os.path.exists(lnk_path)
    except Exception as exc:
        log(f"快捷方式失败 {lnk_path}: {exc}")
        return False


def write_registry(install_dir):
    try:
        import winreg
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, REG_UNINSTALL, 0,
                                 winreg.KEY_SET_VALUE)
        uninstaller = os.path.join(install_dir, "卸载.exe")
        values = {
            "DisplayName": APP_NAME,
            "DisplayVersion": APP_VERSION,
            "Publisher": "moral-edu-automation",
            "InstallLocation": install_dir,
            "DisplayIcon": os.path.join(install_dir, MAIN_EXE) + ",0",
            "UninstallString": f'"{uninstaller}"',
            "NoModify": 1,
            "NoRepair": 1,
        }
        for k, v in values.items():
            winreg.SetValueEx(key, k, 0, winreg.REG_SZ, str(v))
        winreg.CloseKey(key)
        return True
    except Exception as exc:
        log(f"写注册表失败：{exc}")
        return False


def clear_registry():
    try:
        import winreg
        winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, REG_UNINSTALL)
    except FileNotFoundError:
        pass
    except Exception as exc:
        log(f"删注册表失败：{exc}")


def gen_mcp_config(install_dir):
    pack = os.path.join(install_dir, "AI对接")
    exe = os.path.join(pack, "mcp_server.exe")
    payload = {"mcpServers": {"deyufen": {"command": exe, "args": []}}}
    try:
        os.makedirs(pack, exist_ok=True)
        with open(os.path.join(pack, "MCP配置.json"), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as exc:
        log(f"MCP配置.json 生成失败：{exc}")


def do_install(install_dir, progress_cb=None, cancel_flag=None):
    """核心安装：解压到 .install-tmp → 合并（白名单跳过）→ 收尾。返回错误或 None。"""
    zp = payload_zip_path()
    if not os.path.exists(zp):
        return f"找不到内嵌的 payload.zip：{zp}"
    try:
        os.makedirs(install_dir, exist_ok=True)
    except Exception as exc:
        return f"无法创建安装目录：{exc}"

    staging = os.path.join(install_dir, ".install-tmp")
    shutil.rmtree(staging, ignore_errors=True)
    tmp_dir = os.path.join(install_dir, ".install-tmp")
    try:
        os.makedirs(staging)
        total = 0
        with zipfile.ZipFile(zp) as zf:
            infos = [i for i in zf.infolist() if not i.filename.endswith("/")]
            total = sum(i.file_size for i in infos) or 1
            done = 0
            with zf:
                for i in infos:
                    if cancel_flag and cancel_flag[0]:
                        raise RuntimeError("已取消")
                    zf.extract(i, staging)
                    done += i.file_size
                    if progress_cb:
                        progress_cb(done * 50 // total, f"解压 {i.filename}")

        names = [i.filename.replace("\\", "/") for i in infos]
        skipped = [n for n in names if keep_user_data(n)
                   and os.path.lexists(os.path.join(install_dir, n))]
        for n in skipped:
            log(f"保留用户文件（不覆盖）：{n}")

        in_wipe = lambda rel: any(rel == w or rel.startswith(w + "/") for w in WIPE_DIRS)
        k = 0
        for rel in names:
            if in_wipe(rel):
                continue
            if rel in skipped:
                continue
            src = os.path.join(staging, rel)
            dst = os.path.join(install_dir, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            os.replace(src, dst)
            k += 1
            if progress_cb and k % 5 == 0:
                progress_cb(50 + k * 45 // max(1, len(names)), f"安装 {rel}")
        for w in WIPE_DIRS:
            src = os.path.join(staging, w)
            if not os.path.isdir(src):
                continue
            dst = os.path.join(install_dir, w)
            if os.path.isdir(dst):
                shutil.rmtree(dst, ignore_errors=True)
            shutil.move(src, dst)
        if progress_cb:
            progress_cb(96, "生成配置…")
    except Exception as exc:
        shutil.rmtree(staging, ignore_errors=True)
        return f"安装失败：{exc}"
    finally:
        shutil.rmtree(staging, ignore_errors=True)

    gen_mcp_config(install_dir)
    try:
        with open(os.path.join(install_dir, "安装信息.json"), "w", encoding="utf-8") as f:
            json.dump({"version": APP_VERSION,
                       "install_dir": install_dir,
                       "installed_at": datetime.datetime.now().isoformat(timespec="seconds")},
                      f, ensure_ascii=False, indent=2)
    except Exception as exc:
        log(f"安装信息.json 写入失败：{exc}")

    # 卸载器 = 自己的一份副本
    me = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
    uninst = os.path.join(install_dir, "卸载.exe")
    try:
        if getattr(sys, "frozen", False):
            shutil.copy2(me, uninst)
        else:
            shutil.copy2(me, os.path.join(install_dir, "卸载-installer.py"))
    except Exception as exc:
        log(f"卸载器复制失败：{exc}")

    write_registry(install_dir)
    if progress_cb:
        progress_cb(100, "完成")
    log(f"安装完成：{install_dir}（v{APP_VERSION}）")
    return None


def desktop_dir():
    """真实桌面目录（处理 OneDrive/自定义重定向，不能写死 %USERPROFILE%\\Desktop）。"""
    try:
        buf = ctypes.create_unicode_buffer(260)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf) == 0:
            return buf.value
    except Exception:
        pass
    return os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")


def startmenu_dir():
    return os.path.join(os.environ.get("APPDATA", ""),
                        r"Microsoft\Windows\Start Menu\Programs")


def make_shortcuts(install_dir):
    exe = os.path.join(install_dir, MAIN_EXE)
    make_shortcut(os.path.join(desktop_dir(), APP_NAME), exe, install_dir)
    make_shortcut(os.path.join(startmenu_dir(), APP_NAME), exe, install_dir)


def remove_shortcuts():
    for base in (desktop_dir(), startmenu_dir()):
        p = os.path.join(base, APP_NAME + ".lnk")
        try:
            os.remove(p)
        except OSError:
            pass


# ───────────────────────── 卸载 ─────────────────────────
PROGRAM_ITEMS = [MAIN_EXE, "update.exe", "_internal", "AI对接", "软件说明",
                 ".install-tmp", ".update-tmp", "卸载.exe", "卸载-installer.py"]


def do_uninstall(keep_data=True):
    install_dir = os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, "frozen", False) else __file__))
    me = os.path.abspath(sys.executable) if getattr(sys, "frozen", False) else ""
    i_am_uninstaller = bool(me) and os.path.basename(me) == "卸载.exe"
    log(f"开始卸载：{install_dir}（保留用户数据={keep_data}）")
    for name in PROGRAM_ITEMS:
        if i_am_uninstaller and name == "卸载.exe":
            continue                     # 自己正在运行，删不掉；走末尾的延时自删
        p = os.path.join(install_dir, name)
        try:
            if os.path.isdir(p):
                shutil.rmtree(p, ignore_errors=True)
            elif os.path.lexists(p):
                os.remove(p)
        except Exception as exc:
            log(f"删除 {name} 失败：{exc}")
    if keep_data:
        log("按用户选择保留：名单模板.xlsx / 更新配置.json / 月度数据 等")
    else:
        for name in ("名单模板.xlsx", "更新配置.json"):
            try:
                os.remove(os.path.join(install_dir, name))
            except OSError:
                pass
        import glob
        for d in glob.glob(os.path.join(install_dir, "20*年*月")):
            shutil.rmtree(d, ignore_errors=True)
    clear_registry()
    remove_shortcuts()
    # 卸载器自删：ping 延时争取时间（timeout 在无控制台进程里会因 stdin 异常立刻失败）。
    # 注意必须用 args 形式：整串形式会被 cmd 的引号剥离规则弄坏重定向与路径。
    if i_am_uninstaller:
        subprocess.Popen(
            ["cmd", "/c", "ping", "-n", "4", "127.0.0.1", ">nul",
             "&", "del", "/f", "/q", me],
            creationflags=0x08000000)      # CREATE_NO_WINDOW
    log("卸载完成")


# ───────────────────────── GUI 向导 ─────────────────────────
def run_gui():
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title(APP_NAME + " · 安装向导")
    root.geometry("560x420")
    root.resizable(False, False)
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except Exception:
        pass

    pages = {}

    def show(name):
        for p in pages.values():
            p.pack_forget()
        pages[name].pack(fill="both", expand=True)

    # ── 页 1：欢迎 ──
    pg1 = ttk.Frame(root, padding=36)
    pages["welcome"] = pg1
    ttk.Label(pg1, text=f"欢迎使用「{APP_NAME}」", font=("Microsoft YaHei UI", 16, "bold"),
              foreground=INK).pack(anchor="w", pady=(0, 10))
    ttk.Label(pg1, foreground=MUTED, justify="left", wraplength=470, text=(
        f"版本 v{APP_VERSION}\n\n"
        "本向导将安装：\n"
        "  · 主程序（含 AI 对接组件，无需另装 Python）\n"
        "  · 空白名单模板（首次安装写入，之后永不覆盖）\n"
        "  · 桌面与开始菜单快捷方式\n\n"
        "默认安装到用户目录，不需要管理员权限。")).pack(anchor="w")
    ttk.Button(pg1, text="下一步 →", style="Accent.TButton",
               command=lambda: show("dir")).pack(side="right", anchor="se", pady=10)

    # ── 页 2：目录 ──
    pg2 = ttk.Frame(root, padding=36)
    pages["dir"] = pg2
    ttk.Label(pg2, text="安装位置", font=("Microsoft YaHei UI", 14, "bold")).pack(anchor="w")
    ttk.Label(pg2, text="建议保持默认（用户目录，免管理员权限）",
              foreground=MUTED).pack(anchor="w", pady=(4, 14))
    dir_var = tk.StringVar(value=DEFAULT_DIR)
    row = ttk.Frame(pg2)
    row.pack(fill="x")
    en_dir = ttk.Entry(row, textvariable=dir_var)
    en_dir.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def browse():
        d = filedialog.askdirectory(parent=root, initialdir=dir_var.get() or "/",
                                    title="选择安装目录")
        if d:
            dir_var.set(os.path.normpath(d))
    ttk.Button(row, text="浏览…", command=browse).pack(side="left")

    hint = ttk.Label(pg2, foreground=MUTED, justify="left", wraplength=470, text=(
        "注意：\n"
        "  · 若目标目录已有 名单模板.xlsx / 月度数据，安装与升级都不会动它们\n"
        "  · 请选择一个空目录或本软件的旧安装目录"))
    hint.pack(anchor="w", pady=(16, 0))

    def start_install():
        d = dir_var.get().strip()
        if not d or not os.path.isabs(d):
            messagebox.showwarning("安装位置", "请填写完整的安装目录路径", parent=root)
            return
        if not os.path.exists(d):
            try:
                os.makedirs(d, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("安装位置", f"目录创建失败：\n{exc}", parent=root)
                return
        show("progress")
        threading.Thread(target=install_thread, args=(os.path.normpath(d),), daemon=True).start()

    ttk.Button(pg2, text="开始安装", command=start_install).pack(side="right", pady=18)
    ttk.Button(pg2, text="← 上一步", command=lambda: show("welcome")).pack(
        side="right", padx=(0, 10), pady=18)

    # ── 页 3：进度 ──
    pg3 = ttk.Frame(root, padding=36)
    pages["progress"] = pg3
    ttk.Label(pg3, text="正在安装…", font=("Microsoft YaHei UI", 14, "bold")).pack(anchor="w")
    bar = ttk.Progressbar(pg3, maximum=100)
    bar.pack(fill="x", pady=(24, 8))
    lbl = ttk.Label(pg3, text="准备中…", foreground=MUTED)
    lbl.pack(anchor="w")

    def progress_cb(pct, text):
        root.after(0, lambda: (bar.configure(value=pct), lbl.config(text=text)))

    def install_thread(install_dir):
        err = do_install(install_dir, progress_cb)
        def done():
            if err:
                messagebox.showerror("安装失败", err + f"\n\n详见日志：\n{LOG_FILE}", parent=root)
                show("dir")
                return
            make_shortcuts(install_dir)
            final["dir"] = install_dir
            show("finish")
        root.after(0, done)

    # ── 页 4：完成 ──
    final = {"dir": None}
    pg4 = ttk.Frame(root, padding=36)
    pages["finish"] = pg4
    ttk.Label(pg4, text="安装完成！", font=("Microsoft YaHei UI", 16, "bold"),
              foreground="#0D9488").pack(anchor="w", pady=(0, 10))
    ttk.Label(pg4, foreground=MUTED, justify="left", wraplength=470, text=(
        "首次使用：\n"
        "  1. 双击桌面「德育分自动化系统」启动\n"
        "  2. 点工具栏「班级名单」导入学生名单（支持从 Excel 复制粘贴）\n"
        "  3. 以后有新版本，软件侧边栏「检查更新」会自动提示\n\n"
        "提示：首次运行 Windows 可能弹「已保护你的电脑」，点「更多信息 → 仍要运行」"
        "（未购买代码签名证书所致，属正常现象）。")).pack(anchor="w")

    def run_now():
        d = final.get("dir")
        if d:
            subprocess.Popen([os.path.join(d, MAIN_EXE)], cwd=d)
        root.destroy()

    ttk.Button(pg4, text="立即运行", command=run_now).pack(side="right", pady=18)
    ttk.Button(pg4, text="完成", command=root.destroy).pack(side="right", padx=(0, 10), pady=18)

    show("welcome")
    root.mainloop()


def run_silent(install_dir, no_shortcuts=False):
    err = do_install(install_dir)
    if err:
        log(f"[静默安装失败] {err}")
        ctypes.windll.user32.MessageBoxW(0, err, APP_NAME + " 安装", 0x10)
        return 1
    if not no_shortcuts:
        make_shortcuts(install_dir)
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--silent", action="store_true")
    ap.add_argument("--dir", dest="install_dir")
    ap.add_argument("--uninstall", action="store_true")
    ap.add_argument("--no-shortcuts", action="store_true")
    args = ap.parse_args()

    k = ctypes.windll.kernel32
    k.CreateMutexW(None, False, MUTEX_NAME)
    if k.GetLastError() == 183:
        log("已有安装器在运行")
        return 0

    if args.uninstall:
        if args.silent:
            do_uninstall(keep_data=True)
            return 0
        root = tk.Tk()
        root.withdraw()
        keep = messagebox.askyesno(
            "卸载", "确定卸载「德育分自动化系统」吗？\n\n"
            "「是」= 卸载程序，但保留名单模板与月度数据\n"
            "「否」= 全部删除（包括名单与数据）\n"
            "取消 = 放弃卸载", icon="question")
        cancel = messagebox.askokcancel("确认", "最后确认：开始卸载？")
        root.destroy()
        if cancel:
            do_uninstall(keep_data=keep)
        return 0

    if args.silent:
        d = args.install_dir or DEFAULT_DIR
        return run_silent(d, args.no_shortcuts)

    run_gui()
    return 0


if __name__ == "__main__":
    sys.exit(main())
