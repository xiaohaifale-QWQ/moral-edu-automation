# -*- coding: utf-8 -*-
"""v1.9.0.1 综合验证：
1. 空白名单启动不弹任何窗口（开局引导已删）
2. 设置对话框：左侧模块按钮（名单模板/版本与更新/AI 对接），切换正常
3. 名单模板面板：「添加模板…」选文件后直接替换（无确认弹窗）+「打开模板编辑」
4. 底部不裁屏：所有按钮都映射在对话框可视区内
"""
import importlib.util
import os
import sys
import tempfile
import shutil

import openpyxl

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, "软件说明"))
spec = importlib.util.spec_from_file_location(
    "moralapp", os.path.join(REPO, "软件说明", "德育分自动化系统.py"))
mod = importlib.util.module_from_spec(spec)
sys.modules["moralapp"] = mod
spec.loader.exec_module(mod)

FAILS = []


def check(label, ok, extra=""):
    print(f"  [{'OK ' if ok else 'FAIL'}] {label}" + (f" → {extra}" if extra else ""))
    if not ok:
        FAILS.append(label)


workdir = tempfile.mkdtemp(prefix="v1901-")
shutil.copy2(os.path.join(REPO, "_build_pkg", "payload", "名单模板.xlsx"),
             os.path.join(workdir, "名单模板.xlsx"))
rows = [{"sid": f"20250117{i:03d}", "name": f"同学{i}",
         "dorm": "901" if i < 3 else "902", "role": "室长" if i == 0 else "成员"}
        for i in range(6)]
app0 = mod.MoralScoreApp.__new__(mod.MoralScoreApp)
wb = openpyxl.load_workbook(os.path.join(workdir, "名单模板.xlsx"))
app0._write_month_roster(wb["月表模板"], rows, "九点零一班")
app0._write_week_roster(wb["周表模板"], rows, "九点零一班")
app0._write_dorm_roster(wb["寝室名单"], rows)
wb.save(os.path.join(workdir, "用户模板.xlsx"))
wb.close()

mod.MoralScoreApp._base_dir = lambda self: workdir
mod.messagebox.askyesno = lambda *a, **k: FAILS.append(f"意外的askyesno:{a}")
mod.messagebox.showwarning = lambda *a, **k: FAILS.append(f"意外的警告:{a}")
mod.messagebox.showerror = lambda *a, **k: FAILS.append(f"意外的错误:{a}")
opened = []
mod.MoralScoreApp._open_path = lambda self, p: opened.append(p)

toasts = []
mod.notify = lambda root, msg, **k: toasts.append(msg)

import ctypes as _ct
try:
    _ct.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
mod.enable_dpi_awareness()
mod.apply_theme(root, accent=mod.ACCENT, accent_hover=mod.ACCENT_HOVER)
mod.configure_app_styles(root)
root.geometry("1240x820+30+30")
app = mod.MoralScoreApp(root)
app._update_busy = True
for _ in range(12):
    root.update_idletasks(); root.update()

print("### 1. 空白名单启动零弹窗")
tops = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
check("没有出现任何弹窗", len(tops) == 0, str([t.title() for t in tops]))

print("### 2. 设置模块化")
app.show_settings_dialog()
for _ in range(8):
    root.update_idletasks(); root.update()
top = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel) and w.title() == "设置"][0]


def find_btn(w, text):
    for c in w.winfo_children():
        if isinstance(c, ttk.Button) and str(c.cget("text")) == text:
            return c
        r = find_btn(c, text)
        if r:
            return r
    return None


for label in ("名单模板", "版本与更新", "AI 对接"):
    check(f"左侧模块按钮「{label}」", find_btn(top, label) is not None)

# 名单模板面板默认显示：有添加模板 + 打开模板编辑
check("面板有「添加模板…」", find_btn(top, "添加模板…") is not None)
check("面板有「打开模板编辑」", find_btn(top, "打开模板编辑") is not None)
find_btn(top, "打开模板编辑").invoke()
for _ in range(4):
    root.update_idletasks(); root.update()
check("打开模板编辑 → 打开模板文件", any(n.endswith("名单模板.xlsx") for n in opened))

print("### 3. 添加模板 → 直接替换（无确认框）")
mod.filedialog.askopenfilename = lambda *a, **k: os.path.join(workdir, "用户模板.xlsx")
find_btn(top, "添加模板…").invoke()
for _ in range(12):
    root.update_idletasks(); root.update()
check("没有弹确认框", not [w for w in root.winfo_children()
                          if isinstance(w, tk.Toplevel) and w.title() == "拖入名单"])
check("直接导入 6 人", len(app.week_data) == 6, str(len(app.week_data)))
check("班级来自模板", app.class_name == "九点零一班", app.class_name)
check("有成功提示 toast", any("模板已替换" in t for t in toasts), str(toasts))
check("自动备份存在", any(n.startswith("名单模板.自动备份-") for n in os.listdir(workdir)))

print("### 4. 模块切换 + 不裁屏")
find_btn(top, "版本与更新").invoke()
for _ in range(6):
    root.update_idletasks(); root.update()
check("切到版本与更新：检查更新可见", find_btn(top, "检查更新") is not None
      and find_btn(top, "检查更新").winfo_ismapped())
check("保存按钮在可视区", find_btn(top, "保存") is not None
      and find_btn(top, "保存").winfo_ismapped())
bx = top.winfo_rootx() + top.winfo_width()
by = top.winfo_rooty() + top.winfo_height()
b = find_btn(top, "保存")
check("保存按钮未被裁掉", b.winfo_rooty() + b.winfo_reqheight() <= by
      and b.winfo_rootx() + b.winfo_reqwidth() <= bx)
find_btn(top, "AI 对接").invoke()
for _ in range(6):
    root.update_idletasks(); root.update()
check("切到 AI 对接：重新生成 MCP 可见", find_btn(top, "重新生成 MCP 配置") is not None
      and find_btn(top, "重新生成 MCP 配置").winfo_ismapped())
top.destroy()
root.destroy()
shutil.rmtree(workdir, ignore_errors=True)
print("\n[通过]" if not FAILS else f"\n[失败] {FAILS}")
sys.exit(1 if FAILS else 0)
