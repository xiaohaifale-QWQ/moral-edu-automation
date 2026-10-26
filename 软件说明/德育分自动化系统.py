import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ast
import glob
import hashlib
import json
import math
import operator
import os
import shutil
import subprocess
import sys
import datetime
import tempfile
import threading
import time
import urllib.request
import warnings
import re
import openpyxl
import pandas as pd
from copy import copy, deepcopy
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

from tkwebui import (enable_dpi_awareness, apply_theme, SectionCard,
                     Sidebar, SegmentedControl, notify)

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ── 版本与更新源（与 installer/setup.iss、仓库 version.json 保持一致）────
APP_VERSION = "2.1.1.0"
UPDATE_MANIFEST_DEFAULT = "https://raw.githubusercontent.com/xiaohaifale-QWQ/moral-edu-automation/main/version.json"
UPDATE_MANIFEST_JSDELIVR = "https://cdn.jsdelivr.net/gh/xiaohaifale-QWQ/moral-edu-automation@main/version.json"
UPDATE_GH_API_LATEST = "https://api.github.com/repos/xiaohaifale-QWQ/moral-edu-automation/releases/latest"
UPDATE_CONFIG_NAME = "更新配置.json"
UPDATE_DIR_NAME = "deyufen-update"      # %TEMP% 下的下载工作目录


def vtuple(s):
    """'1.8.3.1' → (1, 8, 3, 1)；'1.8.3' → (1, 8, 3, 0)。

    四位版本规则：大更新改第 1 位、小更新改第 2 位、功能修订改第 3 位、
    bug 修复改第 4 位。不足四位补 0，保证不同位数的版本能正确比较。
    带后缀（如 '2.0.1-beta'）取前导数字；解析不了按 (0,0,0,0)。
    """
    parts = []
    try:
        for x in str(s).strip().split(".")[:4]:
            m = re.match(r"\d+", x.strip())
            parts.append(int(m.group(0)) if m else 0)
    except (TypeError, ValueError):
        return (0, 0, 0, 0)
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts)


def newer_version(a, b):
    return vtuple(a) > vtuple(b)

# ── 外观设计令牌 ────────────────────────────────────────────────────────
ACCENT = "#0D9488"          # 主色（青绿）
ACCENT_HOVER = "#0B7A70"
INK = "#1F2937"             # 正文
MUTED = "#6B7280"           # 次要文字
OK_C = "#0D9488"
WARN_C = "#B45309"
ERR_C = "#DC2626"
HILITE = "#CCFBF1"          # 刚变动行的底色（青绿极浅）
TITLE_FONT = ("Microsoft YaHei UI", 15, "bold")
SUB_FONT = ("Microsoft YaHei UI", 11, "bold")
BODY_FONT = ("Microsoft YaHei UI", 10)

# ── 分数表达式安全求值（替代 eval）──────────────────────────────────────
_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv, ast.Pow: operator.pow,
        ast.USub: operator.neg, ast.UAdd: operator.pos}


def safe_eval_math(text, max_power=64, max_digits=12):
    """只允许数字与 + - * / ** () 的算术表达式，拒绝任意代码执行与指数炸弹。

    相比 eval()：'9**9**9' 这类会让界面卡死的输入会被直接拒绝。
    """
    s = text.strip()
    if not s:
        raise ValueError("空表达式")
    if len(s) > 120:
        raise ValueError("表达式过长")
    if not re.fullmatch(r'[\d+\-*/().\s]+', s):
        raise ValueError("只支持数字与 + - * / ( )")

    def walk(node):
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant):
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise ValueError("非法常量")
            if isinstance(node.value, int) and len(str(abs(node.value))) > max_digits:
                raise ValueError("数字过大")
            return node.value
        if isinstance(node, ast.BinOp):
            op = _OPS.get(type(node.op))
            if op is None:
                raise ValueError("不支持的运算符")
            left, right = walk(node.left), walk(node.right)
            if op is operator.pow:
                # 指数炸弹防护
                if abs(right) > max_power or (abs(left) > 1 and abs(right) > max_power):
                    raise ValueError("幂运算指数过大")
            result = op(left, right)
            if isinstance(result, complex) or abs(result) > 1e15:
                raise ValueError("结果超出范围")
            return result
        if isinstance(node, ast.UnaryOp):
            op = _OPS.get(type(node.op))
            if op is None:
                raise ValueError("不支持的一元运算符")
            return op(walk(node.operand))
        raise ValueError("非法表达式")

    tree = ast.parse(s, mode="eval")
    try:
        value = walk(tree)
    except ZeroDivisionError:
        raise ValueError("除数不能为 0")
    except (OverflowError, TypeError) as exc:
        raise ValueError(f"无法计算：{exc}")
    if isinstance(value, complex) or value != value:   # NaN 检查
        raise ValueError("结果不是有效数字")
    if value in (float("inf"), float("-inf")):
        raise ValueError("结果超出范围")
    return float(value)


def signed_text(value):
    """把分值格式化成带正负号的展示文本；0 显示为「不加减」。"""
    if abs(value) < 1e-9:
        return "不加减"
    txt = str(int(value)) if value == int(value) else str(value)
    return ("+" + txt if value > 0 else txt) + "分"


def is_ai_record(rec):
    """判断一条记录是不是 AI 写入的。

    记录本体是 (分数, 原因) 元组；AI 写入的会多挂一个来源字典，
    变成 [分数, 原因, {"src": "ai", "time": "HH:MM:SS"}]。
    其余逻辑仍然按 [0] / [1] 取值，不受影响。
    """
    return (len(rec) >= 3 and isinstance(rec[2], dict)
            and rec[2].get("src") == "ai")


# 寝室相关的明细：AI 写这类记录时软件会优先跳到「寝室统计」页
DORM_REASON_RE = re.compile(r"(寝室|宿管|[3-5]\s*星|三星|四星|五星)")


def is_dorm_reason(reason):
    """判断一条记录的明细是不是寝室星级相关。"""
    return bool(DORM_REASON_RE.search(str(reason or "")))


def configure_app_styles(root):
    """在 apply_theme 之后补齐本程序用到的具名样式。"""
    from tkwebui.core import PALETTE
    s = ttk.Style(root)
    sidebar_bg = PALETTE["sidebar"]
    s.configure("Brand.TLabel", background=sidebar_bg, foreground=INK, font=TITLE_FONT)
    s.configure("SidebarMuted.TLabel", background=sidebar_bg, foreground=MUTED,
                font=("Microsoft YaHei UI", 9))
    s.configure("PageTitle.TLabel", font=TITLE_FONT, foreground=INK)
    s.configure("CardTitle.TLabel", font=SUB_FONT, foreground=INK)
    s.configure("SidebarSuccess.TLabel", background=sidebar_bg, foreground=OK_C,
                font=("Microsoft YaHei UI", 9))
    s.configure("SidebarWarning.TLabel", background=sidebar_bg, foreground=WARN_C,
                font=("Microsoft YaHei UI", 9))
    s.configure("Success.TLabel", foreground=OK_C)
    s.configure("Warning.TLabel", foreground=WARN_C)
    s.configure("Danger.TLabel", foreground=ERR_C)


class MoralScoreApp:
    def center_window(self, width, height):
        """把窗口摆在主屏靠上居中位置并夹在屏内。

        tk::PlaceWindow . center 在多显示器/高 DPI 下会把窗口推到屏幕外，
        导致底部操作栏看不见，所以自己算坐标。
        """
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(width, screen_w - 40)
        height = min(height, screen_h - 80)

        # 留给窗口边框和任务栏的余量
        frame_w = width + 18
        frame_h = height + 47
        x = max(0, (screen_w - frame_w) // 2)
        y = max(0, min((screen_h - frame_h) // 3, screen_h - frame_h - 60))

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _place_dialog(self, top, width, height):
        """把弹窗摆到主窗口正中并夹在屏内。

        原来的 tk::PlaceWindow center 在多显示器下会把弹窗推到屏幕外。
        """
        top.update_idletasks()
        root_x, root_y = self.root.winfo_rootx(), self.root.winfo_rooty()
        root_w, root_h = self.root.winfo_width(), self.root.winfo_height()
        screen_w, screen_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = root_x + (root_w - width) // 2
        y = root_y + (root_h - height) // 3
        x = max(0, min(x, screen_w - width - 12))
        y = max(0, min(y, screen_h - height - 70))
        top.geometry(f"{width}x{height}+{x}+{y}")

    def _log_open_attempt(self, message):
        """把「打开文件/文件夹」的尝试记到 _ai_bridge\\open_path.log，便于排查。"""
        try:
            os.makedirs(self._bridge_dir(), exist_ok=True)
            stamp = datetime.datetime.now().strftime("%m-%d %H:%M:%S")
            with open(os.path.join(self._bridge_dir(), "open_path.log"), "a",
                      encoding="utf-8") as f:
                f.write(f"{stamp}  {message}\n")
        except Exception:
            pass

    def _open_path(self, path):
        """用资源管理器/默认程序打开文件或文件夹。

        注意：以前文件夹一定会失败（WinError 5 / explorer 0xC0000142），
        根因不是代码，而是程序所在目录被打了「低完整性」标签，
        从那里启动的进程无权指挥资源管理器。标签改回「中」之后就正常了。
        这里仍然保留兜底：万一环境又有问题，至少把内容列出来、路径复制走。
        """
        self._log_open_attempt(f"请求打开：{path}  存在={os.path.exists(path)} "
                               f"是目录={os.path.isdir(path)}")
        if not os.path.exists(path):
            self._log_open_attempt("  → 路径不存在")
            return messagebox.showwarning("找不到目标", f"路径不存在：\n{path}")

        if os.path.isdir(path):
            try:
                os.startfile(path)
                self._log_open_attempt("  → 文件夹打开成功")
                return True
            except Exception as exc:
                self._log_open_attempt(f"  → startfile(文件夹) 失败：{type(exc).__name__}: {exc}")
            try:
                subprocess.Popen(["explorer", path])
                self._log_open_attempt("  → explorer 子进程启动")
                return True
            except Exception as exc:
                self._log_open_attempt(f"  → explorer 失败：{type(exc).__name__}: {exc}")
            return self._reveal_folder(path)

        try:
            os.startfile(path)
            self._log_open_attempt("  → startfile 成功")
            return True
        except Exception as exc:
            self._log_open_attempt(f"  → startfile 失败：{type(exc).__name__}: {exc}，改用记事本")
            # 没有默认关联程序时退回记事本
            if self._open_with_notepad(path):
                return True
            return False

    def _open_folder(self, path):
        """打开文件夹。"""
        if not os.path.isdir(path):
            return messagebox.showwarning("文件夹不存在", path)
        return self._open_path(path)

    def _reveal_folder(self, path):
        """复制文件夹路径并列出其中文件，代替「打开资源管理器」。"""
        copied = False
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(path)
            copied = True
        except Exception:
            pass
        try:
            names = sorted(os.listdir(path))
        except Exception:
            names = []
        listing = "\n".join("  · " + n for n in names[:20]) or "  （空目录）"
        tail = ("\n\n路径已复制到剪贴板：按 Win+E 打开资源管理器，"
                "在地址栏按 Ctrl+V 粘贴即可进入。") if copied else ""
        self._log_open_attempt(f"  → 已复制路径到剪贴板，列出 {len(names)} 个文件")
        messagebox.showinfo("文件夹内容", f"{path}\n\n{listing}{tail}")
        return True

    # ====================== AI（MCP）对接 ======================
    # 外部 AI 通过 AI对接\mcp_server.py 把记录写成 _ai_bridge\inbox\*.json，
    # 本程序定时扫描并填入当前周次。改为文件队列是为了两边解耦：
    # AI 那边不用懂 Excel 布局，这边也不用懂 MCP 协议。
    AI_POLL_MS = 2000
    # AI 连接页的列宽：学号后两位 / 姓名 / 加减分 / 明细 / 操作
    AI_COL_WIDTHS = (110, 130, 90, 360, 96)

    def _bridge_dir(self):
        return os.path.join(self._base_dir(), "_ai_bridge")

    def _init_ai_bridge(self):
        try:
            os.makedirs(os.path.join(self._bridge_dir(), "inbox"), exist_ok=True)
            os.makedirs(os.path.join(self._bridge_dir(), "done"), exist_ok=True)
        except Exception:
            pass
        self._bridge_ready = True
        self._bridge_tick = 0
        self._write_bridge_state()
        self.root.after(self.AI_POLL_MS, self._poll_ai_inbox)

    def _write_bridge_state(self, running=True):
        """把自己的当前周期写出去，AI 才知道该往哪一周写。

        每几秒刷新一次，作为心跳；mcp_server 用 90 秒超时判断软件是否还在跑。
        """
        if not getattr(self, "_bridge_ready", False):
            return
        state = {
            "running": bool(running),
            "version": APP_VERSION,
            "year": self.current_year.get(),
            "month": self.current_month.get(),
            "week": self.current_week.get(),
            "class_name": self.class_name,
            "dnd": bool(getattr(self, "dnd_active", False)),
            "base_dir": self._base_dir(),
            "updated": datetime.datetime.now().isoformat(timespec="seconds"),
        }
        try:
            path = os.path.join(self._bridge_dir(), "state.json")
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
        except Exception:
            pass

    def _poll_ai_inbox(self):
        try:
            self._poll_ai_inbox_once()
        except Exception:
            pass
        finally:
            self._bridge_tick += 1
            if self._bridge_tick % 3 == 0:      # 约每 6 秒刷新一次心跳
                self._write_bridge_state()
            self.root.after(self.AI_POLL_MS, self._poll_ai_inbox)

    def _poll_ai_inbox_once(self):
        inbox = os.path.join(self._bridge_dir(), "inbox")
        if not os.path.isdir(inbox):
            return
        try:
            names = sorted(os.listdir(inbox))
        except Exception:
            return
        for name in names:
            if not name.endswith(".json"):
                continue
            path = os.path.join(inbox, name)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    batch = json.load(f)
            except Exception:
                self._archive_ai_file(path, name, ok=False)
                continue

            target = (str(batch.get("year", "")), str(batch.get("month", "")),
                      str(batch.get("week", "")))
            if target != self._cycle_key():
                continue        # 不是当前周期：先留着，等切到那一周再填
            batch_type = batch.get("type")
            if batch_type == "dorm_stars":
                ok = self._apply_ai_dorm_stars(batch)
                self._archive_ai_file(path, name, ok=ok)
                continue
            if batch_type == "undo":
                ok = self._apply_ai_undo(batch)
                self._archive_ai_file(path, name, ok=ok)
                continue
            if batch_type == "screenshot":
                ok = self._apply_ai_screenshot(batch)
                self._archive_ai_file(path, name, ok=ok, tag="shot")
                continue
            if self._is_batch_applied(batch):
                # 同一批被重发（AI 重试/网络重传）：归档留痕，绝不重复写入
                self._archive_ai_file(path, name, tag="dup")
                continue
            if self._apply_ai_batch(batch):
                self._mark_batch_applied(batch)
                self._archive_ai_file(path, name, ok=True)
            else:
                self._archive_ai_file(path, name, ok=False)

    def _apply_ai_undo(self, batch):
        """AI 要求撤回最近几批（只对当前周期的撤回栈有效）。"""
        try:
            times = int(batch.get("count") or 1)
        except (TypeError, ValueError):
            times = 1
        times = max(1, min(times, 5))
        if not getattr(self, "undo_stack", None):
            notify(self.root, "AI 想撤回，但撤回栈是空的（切过周期或已经撤回过了）",
                   title="撤回失败", kind="warn", duration=8000)
            return False
        done = 0
        for _ in range(times):
            if not self.undo_stack:
                break
            self.undo_last_action()
            done += 1
        notify(self.root, f"AI 已撤回 {done} 步", title="已撤回", kind="warn", duration=8000)
        return done > 0

    def _apply_ai_dorm_stars(self, batch):
        """AI 卫生检查结果 → 直接设置寝室星级滑块（不写记录）。

        按次数计分：同一寝室每收到一次结果累计一次，应用打分时
        分数 = 每次的星级分值之和；应用后清零。
        AI 设过的寝室号显示青色；一个结果都没有的显示红色（未发）。
        """
        stars = batch.get("stars")
        if not isinstance(stars, dict) or not stars:
            return False
        applied, unknown = [], []
        for dorm, star in stars.items():
            d = str(dorm).strip()
            v = self.STAR_TEXT_MAP.get(str(star).strip())
            if v is None and str(star).strip() in ("3", "4", "5"):
                v = int(str(star).strip())
            if v is None:
                unknown.append(f"{dorm}={star}")
                continue
            if d not in self.dorm_stars:
                unknown.append(f"{dorm}（名单里没有这个寝室）")
                continue
            self._set_dorm_star(d, v, origin="ai")
            applied.append(f"{d} {'三四五'[v - 3]}星")
        if not applied:
            return False
        self.show_dorm_view()
        try:
            msg = f"AI 已设置 {len(applied)} 个寝室星级（青色标记），核对后点「应用打分」"
            if unknown:
                msg += f"；未识别 {len(unknown)} 个"
            notify(self.root, msg, title="AI 卫生检查", kind="success", duration=10000)
        except Exception:
            pass
        return True

    def _apply_ai_batch(self, batch):
        """把一批记录填进 week_data。返回是否真的写入了。"""
        records = batch.get("records") or []
        if not records:
            return False

        pairs = []
        for rec in records:
            sid = str(rec.get("sid") or "")
            if sid not in self.week_data:
                continue
            try:
                score = float(rec.get("score"))
            except (TypeError, ValueError):
                continue
            if not math.isfinite(score):
                continue
            reason = str(rec.get("reason") or "").strip() or "AI 记录"
            match = rec.get("match") if isinstance(rec.get("match"), dict) else None
            auto = rec.get("auto_score") if isinstance(rec.get("auto_score"), dict) else None
            pairs.append((sid, score, reason, match, auto))

        if not pairs:
            return False

        # replace=true：先删掉这些学生「同一明细」的旧 AI 记录，再写新的。
        # 这样同一个奖项重发一次是「覆盖」，不会越改越多。
        if batch.get("replace"):
            keys = {(sid, reason) for sid, _score, reason, _m, _a in pairs}
            for sid, info in self.week_data.items():
                recs = info.get('records') or []
                if not recs:
                    continue
                keep = [r for r in recs
                        if not (is_ai_record(r) and (sid, str(r[1])) in keys)]
                if len(keep) != len(recs):
                    info['records'] = keep

        self._push_undo()          # 先快照，再改动
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        touched = []
        fuzzy_count = 0
        auto_count = 0
        for sid, score, reason, match, auto in pairs:
            # 挂上来源标记，AI 连接页据此列出「我跟 AI 说了什么」；
            # 近似匹配的再记下原文和相似度，界面用橙色提示人工核对
            meta = {"src": "ai", "time": stamp}
            if match and match.get("fuzzy"):
                meta["fuzzy"] = True
                meta["input"] = str(match.get("input") or "")
                try:
                    meta["rate"] = float(match.get("rate") or 0)
                except (TypeError, ValueError):
                    meta["rate"] = 0.0
                fuzzy_count += 1
            if auto:
                # 分数是 MCP 服务按模板「加分预览」自动查出来的，界面标青提示核对
                meta["auto"] = True
                meta["auto_from"] = str(auto.get("from") or "")
                auto_count += 1
            self.week_data[sid]['records'].append([score, reason, meta])
            touched.append(sid)

        self.last_edited_sids.update(touched)
        self.unsaved = True
        if not self.showing_inquiry:
            self.refresh_week_tree()

        names = "、".join(self.week_data[sid]['name'] for sid in touched[:3])
        more = " 等" if len(touched) > 3 else ""
        dorm_count = sum(1 for _, _, r, _, _ in pairs if is_dorm_reason(r))
        notes = []
        if fuzzy_count:
            notes.append(f"{fuzzy_count} 条姓名近似，见橙色行")
        if auto_count:
            notes.append(f"{auto_count} 条分值按模板自动取，见青色分值")
        if dorm_count:
            notes.append(f"{dorm_count} 条寝室记录，已切到「寝室统计」页")

        # 寝室相关的记录优先给寝室统计页看：直接跳过去
        if dorm_count:
            self.show_dorm_view()
        elif self.current_view == "ai":
            self.refresh_ai_view()

        if notes:
            notify(self.root,
                   f"AI 已填入 {len(touched)} 条：{names}{more}（{'；'.join(notes)}）",
                   title="AI 导入", kind="warn", duration=10000)
        else:
            notify(self.root, f"AI 已填入 {len(touched)} 条：{names}{more}",
                   title="AI 导入", kind="success", duration=8000)

        self.week_status.config(text=f"AI 导入了 {len(touched)} 条记录", foreground=OK_C)
        return True

    def _archive_ai_file(self, path, name, ok=True, tag=""):
        """处理过的批次移到 done 目录留痕，避免重复扫描。"""
        try:
            done = os.path.join(self._bridge_dir(), "done")
            os.makedirs(done, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            prefix = tag or ("ok" if ok else "bad")
            os.replace(path, os.path.join(done, f"{stamp}-{prefix}-{name}"))
        except Exception:
            pass

    # ---- 批次幂等：同一个 batch id 只应用一次，AI 重发不会翻倍 ----
    def _applied_batches_path(self):
        return os.path.join(self._bridge_dir(), "applied_batches.json")

    def _load_applied_batches(self):
        if not hasattr(self, "_applied_cache"):
            self._applied_cache = None
        if self._applied_cache is None:
            data = {}
            try:
                with open(self._applied_batches_path(), "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    data = loaded
            except Exception:
                data = {}
            self._applied_cache = data
        return self._applied_cache

    def _batch_token(self, batch):
        """批次身份：优先用 AI 给的 key，其次用批次 id。"""
        token = str(batch.get("batch_key") or batch.get("id") or "").strip()
        return token

    def _is_batch_applied(self, batch):
        token = self._batch_token(batch)
        if not token:
            return False
        return token in self._load_applied_batches()

    def _mark_batch_applied(self, batch):
        token = self._batch_token(batch)
        if not token:
            return
        cache = self._load_applied_batches()
        cache[token] = datetime.datetime.now().isoformat(timespec="seconds")
        if len(cache) > 500:          # 只留最近 500 个，别让文件无限长
            for old in sorted(cache, key=lambda k: cache[k])[:-500]:
                cache.pop(old, None)
        try:
            os.makedirs(self._bridge_dir(), exist_ok=True)
            with open(self._applied_batches_path(), "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def on_close(self):
        """关窗前拦截：有未保存录入时先问一句。"""
        action = self._ask_unsaved("退出", allow_carry=False)
        if action == "cancel":
            return
        if action == "save":
            self.save_week_data()
            if self.unsaved:
                return
        self._write_bridge_state(running=False)
        self.root.destroy()

    def __init__(self, root):
        self.root = root
        self.root.title(f"德育分自动化系统 v{APP_VERSION}")
        self.root.minsize(1040, 680)
        self.center_window(1240, 820)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.current_view = "week"

        # 周期变量
        now = datetime.datetime.now()
        self.current_year = tk.StringVar(value=str(now.year))
        self.current_month = tk.StringVar(value=str(now.month))
        self.current_week = tk.StringVar(value="1")

        # 核心数据
        self.class_name = "未设置班级"
        self.dorm_map = {}
        self.dorm_roles = {}  # 记录是室长还是成员
        self.week_data = {}
        self.base_layout = []
        self.inquiry_data = None  # 缓存加分标准
        self.showing_inquiry = False  # 切换状态
        self.dorm_period_var = tk.StringVar(value="周")  # 打分面板状态

        # 月表专用数据
        self.month_sum = {}
        self.last_month_scores = {}
        self.week_scanned = False
        self.last_month_loaded = False
        self.is_first_month = tk.BooleanVar(value=False)

        # 打分面板专用（dorm_ai_set：AI 已代设星级的寝室，青色标记，手动改动即清除）
        self.dorm_stars = {}
        self.dorm_controls = {}
        self.dorm_ai_set = set()
        self.dorm_applied = set()

        # 撤回 / 重做（限制深度，避免长期使用内存单调增长）
        self.undo_stack = []
        self.redo_stack = []
        self.last_edited_sid = None
        self.last_edited_sids = set()

        # 未保存标记：切周次/切月份前据此提示，避免静默丢失录入
        self.unsaved = False
        self._loaded_cycle = (self.current_year.get(), self.current_month.get(), self.current_week.get())

        self.setup_ui()
        self.load_templates()
        self.setup_dorm_view()
        self._sync_dorm_stars()   # 启动时把该周期的星级还原出来
        self._init_ai_bridge()    # 开始接收 AI（MCP）排入的记录

        # 主界面稳定后静默检查更新（有新版本才弹窗，失败静默）。
        # 名单为空不再弹窗打扰，侧边栏与状态栏的提示足以引导。
        self.root.after(3000, lambda: self._check_update(manual=False))

    def setup_ui(self):
        shell = ttk.Frame(self.root)
        shell.pack(fill="both", expand=True)

        # ========== 左侧导航栏 ==========
        side = ttk.Frame(shell, width=186, style="Sidebar.TFrame")
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        brand = ttk.Frame(side, style="Sidebar.TFrame", padding=(18, 22, 18, 16))
        brand.pack(fill="x")
        ttk.Label(brand, text="德育分助手", style="Brand.TLabel").pack(anchor="w")
        self.lbl_class = ttk.Label(brand, text="读取中...", style="SidebarMuted.TLabel")
        self.lbl_class.pack(anchor="w", pady=(3, 0))

        self.nav = Sidebar(side, [
            ("week", "周纪实表", "controls"),
            ("month", "月度汇总", "speed"),
            ("dorm", "寝室统计", "library"),
            ("ai", "AI 连接", "sources"),
        ], self._on_nav, width=186)

        # 侧边栏底部：版本号 + 设置入口。
        # 必须在 nav.pack 之前先 pack(side="bottom")，否则会被撑满的导航挤没。
        side_foot = ttk.Frame(side, style="Sidebar.TFrame", padding=(18, 6, 18, 14))
        side_foot.pack(side="bottom", fill="x")
        self.lbl_version = ttk.Label(side_foot, text=f"v{APP_VERSION}",
                                     style="SidebarMuted.TLabel")
        self.lbl_version.pack(anchor="w")
        # 设置按钮：带边框的长方形条（外层 1px 描边 + 悬停浅青）
        from tkwebui.core import PALETTE
        _sb = PALETTE["sidebar"]
        self.btn_open_dir = ttk.Button(side_foot, text="打开文件夹",
                                       style="SidebarRect.TButton",
                                       command=self.open_data_folder)
        self.btn_open_dir.pack(fill="x", pady=(6, 0))
        set_edge = tk.Frame(side_foot, bg="#C9D4D0", padx=1, pady=1)
        set_edge.pack(fill="x", pady=(6, 0))
        set_inner = tk.Frame(set_edge, bg=_sb)
        set_inner.pack(fill="both", expand=True)
        self.btn_settings = tk.Label(
            set_inner, text="设  置", bg=_sb, fg=INK,
            font=("Microsoft YaHei UI", 10), pady=7, cursor="hand2")
        self.btn_settings.pack(fill="x")

        def _settings_hover(e):
            self.btn_settings.config(bg="#E7F2F0")
            set_inner.config(bg="#E7F2F0")

        def _settings_leave(e):
            self.btn_settings.config(bg=_sb)
            set_inner.config(bg=_sb)

        self.btn_settings.bind("<Button-1>",
                               lambda e: self.show_settings_dialog())
        self.btn_settings.bind("<Enter>", _settings_hover)
        self.btn_settings.bind("<Leave>", _settings_leave)

        self.nav.pack(fill="both", expand=True)

        # ========== 右侧主区 ==========
        main = ttk.Frame(shell, padding=(24, 18, 24, 18))
        main.pack(side="left", fill="both", expand=True)

        # 页头：页面名 + 考评周期
        head = ttk.Frame(main)
        head.pack(fill="x", pady=(0, 16))

        titles = ttk.Frame(head)
        titles.pack(side="left")
        self.lbl_page_title = ttk.Label(titles, text="周纪实表", style="PageTitle.TLabel")
        self.lbl_page_title.pack(anchor="w")
        self.lbl_page_sub = ttk.Label(titles, text="", style="Muted.TLabel")
        self.lbl_page_sub.pack(anchor="w", pady=(3, 0))

        period = ttk.Frame(head)
        period.pack(side="right")
        ttk.Label(period, text="考评周期", style="Muted.TLabel").pack(side="left", padx=(0, 8))

        current_y = datetime.datetime.now().year
        year_cb = ttk.Combobox(period, textvariable=self.current_year, width=6, state="readonly",
                               values=[str(i) for i in range(current_y - 1, current_y + 4)])
        year_cb.pack(side="left")
        ttk.Label(period, text="年", style="Muted.TLabel").pack(side="left", padx=(4, 12))

        month_cb = ttk.Combobox(period, textvariable=self.current_month, width=4, state="readonly",
                                values=[str(i) for i in range(1, 13)])
        month_cb.pack(side="left")
        ttk.Label(period, text="月", style="Muted.TLabel").pack(side="left", padx=(4, 0))

        year_cb.bind("<<ComboboxSelected>>", lambda e: self.on_cycle_change())
        month_cb.bind("<<ComboboxSelected>>", lambda e: self.on_cycle_change())

        # 视图容器
        self.body = ttk.Frame(main)
        self.body.pack(fill="both", expand=True)
        self.frame_week = ttk.Frame(self.body)
        self.frame_month = ttk.Frame(self.body)
        self.frame_dorm = ttk.Frame(self.body)
        self.frame_ai = ttk.Frame(self.body)

        self.setup_week_view()
        self.setup_month_view()
        self.setup_ai_view()
        self.show_week_view()
        self._set_window_icon()
        self._enable_drag_drop()

    def _set_window_icon(self):
        """标题栏/任务栏用程序自己的图标（不设的话 Windows 会显示 Tk 默认羽毛）。"""
        if getattr(sys, "frozen", False):
            base = getattr(sys, "_MEIPASS", None) or \
                os.path.dirname(os.path.abspath(sys.executable))
            path = os.path.join(base, "doro.ico")
            if not os.path.exists(path):
                path = os.path.join(os.path.dirname(os.path.abspath(sys.executable)),
                                    "doro.ico")
        else:
            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "doro.ico")
        try:
            if os.path.exists(path):
                self.root.iconbitmap(path)
        except Exception:
            pass

    def _enable_drag_drop(self):
        """把名单模板 xlsx/csv 拖进窗口任意位置即可导入。

        - 依赖 tkinterdnd2（打包已内置）；不可用时状态栏一次性提示并记入
          state.json 的 "dnd" 字段，方便远程排查
        - 主窗口所有非输入类控件都注册成放置目标（含以后新建的对话框），
          避免拖到某个卡片上没反应
        """
        self.dnd_active = False
        try:
            from tkinterdnd2 import DND_FILES
            self.root.tk.call("package", "present", "tkdnd")
        except Exception:
            self.dnd_active = False
            try:
                self.week_status.config(text="拖拽导入不可用（缺少组件），请用 设置 → 添加模板 导入",
                                        foreground=WARN_C)
            except Exception:
                pass
            return

        def register(w):
            try:
                w.drop_target_register(DND_FILES)
                w.bind("<<Drop>>", self._handle_dropped_files)
            except Exception:
                pass

        def walk(w):
            for c in w.winfo_children():
                if isinstance(c, (ttk.Entry, tk.Entry, tk.Text, ttk.Combobox,
                                  ttk.Spinbox, ttk.Button, tk.Button)):
                    continue          # 输入件/按钮不抢拖放
                register(c)
                walk(c)

        register(self.root)
        walk(self.root)
        self.dnd_active = True

    def _handle_dropped_files(self, event):
        """拖入名单文件 → 确认 → 填班级名 → 走现成的写模板/重载流程。"""
        try:
            paths = [str(p) for p in self.root.tk.splitlist(event.data)]
        except Exception:
            return
        target = next((p for p in paths if str(p).lower().endswith((".xlsx", ".csv"))), None)
        if not target or not os.path.exists(target):
            # 拖了文件但不是名单格式：给一句轻提示（快捷方式/.xls 等常见误区）
            name = paths[0] if paths else ""
            hint = "只支持拖入 .xlsx / .csv 名单文件"
            if name.lower().endswith(".lnk"):
                hint = "拖的是快捷方式，请拖文件本身（.xlsx）"
            elif name.lower().endswith(".xls"):
                hint = "旧版 .xls 不支持，请先在 Excel/WPS 里另存为 .xlsx"
            try:
                notify(self.root, hint, title="拖入无效", kind="warning", duration=6000)
            except Exception:
                pass
            return
        try:
            parts_list, cls_hint = self._read_roster_source(target)
        except Exception as exc:
            return messagebox.showerror("读取失败", f"{target}\n\n{exc}")
        self._import_parsed_roster(target, parts_list, cls_hint)

    @staticmethod
    def _is_full_template(path):
        """是不是一份完整的《名单模板.xlsx》（周表/月表/寝室三表齐）。"""
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                names = set(wb.sheetnames)
            finally:
                wb.close()
            return {"周表模板", "月表模板", "寝室名单"} <= names
        except Exception:
            return False

    def _import_whole_template(self, path):
        """整文件替换当前名单模板（1:1 克隆，保排版/配对/寝室顺序/签名栏）。

        用于拖入/添加的是完整名单模板的场景——比解析重建更忠实：
        周表左右配对顺序、寝室排列、周表模板2 签名栏全部原样保留。
        """
        tpl = self._template_path()
        if os.path.abspath(path) == os.path.abspath(tpl):
            try:
                notify(self.root, "这就是当前正在使用的模板，无需导入。",
                       title="添加模板", kind="info", duration=6000)
            except Exception:
                pass
            return
        bak = os.path.join(os.path.dirname(tpl), "名单模板.自动备份-"
                           + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx")
        try:
            shutil.copy2(tpl, bak)
        except Exception as exc:
            if not messagebox.askyesno("备份失败",
                                       f"替换前自动备份失败：{exc}\n仍要继续吗？（有风险）"):
                return
        while True:
            try:
                shutil.copy2(path, tpl)
                break
            except PermissionError:
                if not self.ask_retry_save(tpl):
                    return
        # 整文件换过之后，类名等全部以新模板为准，直接重载
        self.class_name = "未设置班级"
        self.reload_template(ask=False)
        try:
            notify(self.root, f"模板已整体替换：{os.path.basename(path)}（原文件已备份）",
                   title="添加模板", kind="success", duration=8000)
        except Exception:
            pass

    def _import_parsed_roster(self, target, parts_list, cls_hint=""):
        """解析名单并弹确认框（拖入 / 设置按钮 / 对话框共用）。

        来源本身是完整名单模板时改走整文件替换（更忠实，保留原表全部信息）。
        """
        if self._is_full_template(target):
            return self._import_whole_template(target)
        rows, errors = self._parse_roster_parts(parts_list)
        if not rows:
            return messagebox.showwarning(
                "名单为空",
                "没能从文件里解析出名单。\n"
                + (errors[0] if errors else "请确认列序：学号、姓名、寝室号、角色"))
        self._show_drop_import_dialog(target, rows, errors, cls_hint)

    def pick_roster_file(self, parent=None):
        """设置里的「添加模板」：选完文件直接替换当前模板（自动备份，不弹确认）。

        - 选的是完整名单模板 → 整文件 1:1 替换（保留排版/配对/寝室顺序/签名栏）
        - 选的是四列名单文件 → 解析重建并沿用/要求班级名
        parent 必须用关键字传（挂到设置框上，文件选择框才弹在它上面）；
        默认打开程序目录，名单模板就在那里。
        """
        kwargs = dict(title="选择 名单模板.xlsx 或名单文件",
                      filetypes=[("名单文件", "*.xlsx *.csv"), ("所有文件", "*.*")],
                      initialdir=self._base_dir())
        if parent is not None:
            kwargs["parent"] = parent
        path = filedialog.askopenfilename(**kwargs)
        if not path:
            return
        if self._is_full_template(path):
            return self._import_whole_template(path)
        try:
            parts_list, cls_hint = self._read_roster_source(path)
        except Exception as exc:
            return messagebox.showerror("读取失败", f"{path}\n\n{exc}")
        rows, errors = self._parse_roster_parts(parts_list)
        if not rows:
            return messagebox.showwarning(
                "名单为空",
                "没能从文件里解析出名单。\n"
                + (errors[0] if errors else
                   "请确认是 名单模板.xlsx，或 学号/姓名/寝室号/角色 四列名单文件"))
        cls = cls_hint or (self.class_name if self.class_name != "未设置班级" else "")
        if not cls:
            return messagebox.showwarning(
                "缺少班级名",
                "文件里没有班级信息，当前也还没设置过班级。\n"
                "请选带班级名的 名单模板.xlsx。")
        self._apply_roster_to_template(rows, cls)
        try:
            notify(self.root, f"模板已替换：{len(rows)} 人"
                   + (f"（{len(errors)} 行跳过）" if errors else ""),
                   title="添加模板", kind="success", duration=8000)
        except Exception:
            pass

    def open_template_file(self):
        """用 Excel 打开 名单模板.xlsx 直接编辑；改完保存并关闭后点「刷新模板」生效。"""
        path = self._template_path()
        if not os.path.exists(path):
            return messagebox.showwarning("找不到模板", f"没有找到：\n{path}")
        self._open_path(path)

    def _show_drop_import_dialog(self, path, rows, errors, cls_hint=""):
        """拖入名单后的确认框：识别结果 + 班级名 + 「确定导入」一步完成。"""
        top = tk.Toplevel(self.root)
        top.title("拖入名单")
        top.grab_set()
        top.resizable(False, False)
        self._place_dialog(top, 500, 330)
        wrap = ttk.Frame(top, padding=24)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text=f"已识别 {len(rows)} 名学生",
                  font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        ttk.Label(wrap, text=f"文件：{os.path.basename(path)}",
                  style="Muted.TLabel").pack(anchor="w", pady=(4, 0))
        if errors:
            ttk.Label(wrap, wraplength=440, justify="left", foreground=WARN_C,
                      text=f"{len(errors)} 行有问题（导入时跳过）：{errors[0]}"
                           + (" 等" if len(errors) > 1 else "")
                      ).pack(anchor="w", pady=(6, 0))
        ttk.Label(wrap, text="班级全称", style="Muted.TLabel").pack(anchor="w", pady=(14, 2))
        en_cls = ttk.Entry(wrap)
        en_cls.pack(fill="x")
        prefill = cls_hint or ("" if self.class_name == "未设置班级" else self.class_name)
        en_cls.insert(0, prefill)
        en_cls.focus_set()
        en_cls.select_range(0, tk.END)
        ttk.Label(wrap, wraplength=440, style="Muted.TLabel", text=(
            "班级名写进周表表头与月表；原名单模板导入前会自动备份，"
            "当前未保存的加减分录入会被清除。")).pack(anchor="w", pady=(10, 0))

        def ok():
            cls = en_cls.get().strip()
            if not cls:
                messagebox.showwarning("拖入名单", "请填写班级全称", parent=top)
                return
            top.destroy()
            self._apply_roster_to_template(rows, cls)

        def cancel():
            top.destroy()

        btns = ttk.Frame(wrap)
        btns.pack(fill="x", pady=(16, 0))
        ttk.Button(btns, text="确定导入", style="Primary.TButton",
                   command=ok).pack(side="right")
        ttk.Button(btns, text="取消", command=cancel).pack(side="right", padx=(0, 8))
        top.bind("<Return>", lambda e: ok())
        top.bind("<Escape>", lambda e: cancel())
        top.protocol("WM_DELETE_WINDOW", cancel)

    @staticmethod
    def _read_roster_source(path):
        """读名单来源，返回 (分列行, 班级名提示)。

        两种文件都认：
        - 普通名单文件：首表按 学号/姓名/寝室号/角色 四列读（csv 同理）
        - 《名单模板.xlsx》本身（含「月表模板」表）：完整导入——名单取
          月表模板 A/B 列，班级名取 C2，寝室与室长取「寝室名单」表，
          这样把模板拖进来，寝室页的星级滑块立刻就有数据
        """
        if path.lower().endswith(".csv"):
            with open(path, "r", encoding="utf-8-sig") as f:
                return [re.split(r"[\t,，;；]+", ln) for ln in f.read().splitlines()], ""
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            if "月表模板" in wb.sheetnames:
                parts, cls_hint = [], ""
                for r, row in enumerate(wb["月表模板"].iter_rows(values_only=True), 1):
                    cells = ["" if v is None else str(v).strip() for v in row]
                    if r == 2 and len(cells) > 2 and cells[2]:
                        cls_hint = cells[2]
                    sid = cells[0] if cells else ""
                    name = cells[1] if len(cells) > 1 else ""
                    if sid and "签名" not in sid and "考评" not in sid and name:
                        parts.append([sid, name])
                if "寝室名单" in wb.sheetnames:
                    dorm_of, role_of = {}, {}
                    for row in wb["寝室名单"].iter_rows(values_only=True):
                        cells = ["" if v is None else str(v).strip() for v in row]
                        room = cells[0] if cells else ""
                        ppl = cells[2] if len(cells) > 2 else ""
                        leader = cells[6] if len(cells) > 6 else ""
                        if not room or "房" in room:
                            continue
                        for n in re.split(r"[，,\s]+", ppl):
                            if n:
                                dorm_of[n] = room
                                role_of.setdefault(n, "成员")
                        for n in re.split(r"[，,\s]+", leader):
                            if n:
                                dorm_of[n] = room
                                role_of[n] = "室长"
                    parts = [[sid, name, dorm_of.get(name, ""), role_of.get(name, "成员")]
                             for sid, name in parts]
                return parts, cls_hint
            ws = wb.worksheets[0]
            return [list(row[:4]) for row in ws.iter_rows(values_only=True)], ""
        finally:
            wb.close()

    @staticmethod
    def _parse_roster_parts(parts_list):
        """名单校验：学号 8~12 位唯一、姓名非空、角色只能是 成员/室长。"""
        rows, errors, seen = [], [], set()
        for parts in parts_list:
            parts = [str(p).strip() for p in parts if str(p).strip()]
            if not parts:
                continue
            if parts[0].lower() in ("学号", "sid", "id"):   # 表头行
                continue
            if len(parts) < 2:
                errors.append(f"行「{' '.join(parts)}」列数不足")
                continue
            sid, name = parts[0], parts[1]
            dorm = parts[2] if len(parts) > 2 else ""
            role = "室长" if "长" in (parts[3] if len(parts) > 3 else "") else "成员"
            if not re.fullmatch(r"\d{8,12}", sid):
                errors.append(f"学号「{sid}」不是 8~12 位数字")
                continue
            if sid in seen:
                errors.append(f"学号「{sid}」重复")
                continue
            seen.add(sid)
            rows.append({"sid": sid, "name": name, "dorm": dorm, "role": role})
        return rows, errors

    def _all_frames(self):
        return (self.frame_week, self.frame_month, self.frame_dorm, self.frame_ai)

    def _period_text(self):
        base = f"{self.current_year.get()}年{self.current_month.get()}月"
        # 月度汇总页展示「第N周」没有意义
        if self.current_view == "month":
            return base
        if self.current_view == "ai":
            return f"AI 写入目标：{base} · 第{self.current_week.get()}周"
        return f"{base} · 第{self.current_week.get()}周"

    def _on_nav(self, key):
        if key == "week":
            self.show_week_view()
        elif key == "month":
            self.show_month_view()
        elif key == "ai":
            self.show_ai_view()
        else:
            self.show_dorm_view()

    def _activate(self, key, frame, title):
        for f in self._all_frames():
            f.pack_forget()
        frame.pack(fill="both", expand=True)
        self.current_view = key
        if getattr(self, "nav", None) and self.nav.selected_key != key:
            self.nav.select(key, notify=False)
        self.lbl_page_title.config(text=title)
        self.lbl_page_sub.config(text=self._period_text())
        # 清掉按钮焦点，否则切视图后首个按钮显示为焦点态（看起来像禁用）
        self.root.focus_set()

    def show_week_view(self):
        self._activate("week", self.frame_week, "周纪实表")

    def show_month_view(self):
        self._activate("month", self.frame_month, "月度汇总")
        self.check_month_ready()

    def show_dorm_view(self):
        self._activate("dorm", self.frame_dorm, "寝室统计")
        self.lbl_dorm_week.config(text=f"当前周次：第{self.current_week.get()}周")
        self._rebuild_dorm_ai_rows()

    def show_ai_view(self):
        self._activate("ai", self.frame_ai, "AI 连接")
        self.refresh_ai_view()

    def setup_week_view(self):
        # ---------- 工具条 ----------
        bar = ttk.Frame(self.frame_week)
        bar.pack(fill="x", pady=(0, 12))

        ttk.Label(bar, text="编辑周次", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        week_cb = ttk.Combobox(bar, textvariable=self.current_week, values=["1", "2", "3", "4", "5"],
                               width=4, state="readonly")
        week_cb.pack(side="left")
        week_cb.bind("<<ComboboxSelected>>", lambda e: self.on_week_change())
        ttk.Label(bar, text="周", style="Muted.TLabel").pack(side="left", padx=(4, 18))

        self.btn_inquiry = ttk.Button(bar, text="查加分标准", command=self.toggle_score_inquiry)
        self.btn_inquiry.pack(side="left", padx=(0, 6))
        self.btn_undo = ttk.Button(bar, text="左撤回", state="disabled", command=self.undo_last_action)
        self.btn_undo.pack(side="left", padx=(0, 6))
        self.btn_redo = ttk.Button(bar, text="右撤回", state="disabled", command=self.redo_last_action)
        self.btn_redo.pack(side="left", padx=(0, 6))
        # 「刷新模板」收进设置；「班级名单」入口由 拖入 / 设置·添加模板 取代

        self.week_status = ttk.Label(bar, text="等待操作", style="Muted.TLabel")
        self.week_status.pack(side="right")

        # ---------- 录入区 ----------
        row = ttk.Frame(self.frame_week)
        row.pack(fill="x", pady=(0, 12))

        man = SectionCard(row, text="手动添加记录")
        man.pack(side="left", fill="both", expand=True)
        form = ttk.Frame(man)
        form.pack(fill="x", pady=(12, 4))

        ttk.Label(form, text="姓名").grid(row=0, column=0, sticky="w")
        self.en_name = ttk.Entry(form, width=12)
        self.en_name.grid(row=0, column=1, padx=(6, 16))
        ttk.Label(form, text="原因").grid(row=0, column=2, sticky="w")
        self.en_reason = ttk.Entry(form, width=16)
        self.en_reason.grid(row=0, column=3, padx=(6, 16))
        ttk.Label(form, text="分值").grid(row=0, column=4, sticky="w")
        self.en_score = ttk.Entry(form, width=7)
        self.en_score.grid(row=0, column=5, padx=(6, 16))

        self.en_score.bind("<Return>", lambda e: self.add_manual_score())
        ttk.Button(form, text="添加", style="Primary.TButton",
                   command=self.add_manual_score).grid(row=0, column=6)

        # ---------- 底部操作（先占位，避免被表格挤出可视区） ----------
        foot = ttk.Frame(self.frame_week)
        foot.pack(side="bottom", fill="x", pady=(14, 0))
        ttk.Button(foot, text="截图", command=lambda: self.screenshot_tree(
            self.tree_week, f"周纪实表截图_{self.class_name}{self.current_month.get()}月第{self.current_week.get()}周")).pack(side="left", padx=(12, 0))
        ttk.Button(foot, text="生成周表", style="Primary.TButton",
                   command=self.save_week_data).pack(side="left")
        ttk.Label(foot, text="按当前周次生成周纪实表（月度汇总需重新读取后才更新）",
                  style="Muted.TLabel").pack(side="left", padx=(14, 0))

        # ---------- 预览表 ----------
        card = SectionCard(self.frame_week, text="本周实时预览")
        card.pack(fill="both", expand=True)
        self.pv_frame = card
        self.pv_title = card.caption

        table = ttk.Frame(card)
        table.pack(fill="both", expand=True, pady=(12, 0))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        self.tree_week = ttk.Treeview(table, show="headings", height=12)
        self.setup_week_tree_columns()

        scr = ttk.Scrollbar(table, orient="vertical", command=self.tree_week.yview)
        self.tree_week.configure(yscrollcommand=scr.set)
        self.tree_week.grid(row=0, column=0, sticky="nsew")
        scr.grid(row=0, column=1, sticky="ns")

    def setup_week_tree_columns(self):
        """恢复/初始化学生预览列"""
        cols = ("学号", "姓名", "总加分", "项目明细", "寝室角色", "寝室号")
        self.tree_week["columns"] = cols
        for col in cols:
            self.tree_week.heading(col, text=col)
            self.tree_week.column(col, width=100 if col != "项目明细" else 250, anchor="center")
        # 只在首次绑定，否则每次从查询视图切回来都会叠加一个绑定点，
        # 双击时会同时弹出多个编辑框
        if not getattr(self, "_tree_dbl_bound", False):
            self.tree_week.bind("<Double-Button-1>", self._on_tree_double_click)
            self._tree_dbl_bound = True

    def toggle_score_inquiry(self):
        """切换查询/录入视图"""
        if not self.showing_inquiry:
            self.showing_inquiry = True
            self.btn_inquiry.config(text="返回录入", style="Primary.TButton")
            self.pv_title.config(text="加分标准查询")
            self.display_inquiry_view()
        else:
            self.showing_inquiry = False
            self.btn_inquiry.config(text="查加分标准", style="TButton")
            self.pv_title.config(text="本周实时预览")
            self.setup_week_tree_columns()
            self.refresh_week_tree()

    def display_inquiry_view(self):
        """显示加分预览表，彻底过滤 nan 字符"""
        if self.inquiry_data is None:
            try:
                df = pd.read_excel(self._template_path(), sheet_name="加分预览")
                df = df.fillna("")
                df = df.astype(str).replace('nan', '')
                self.inquiry_data = df
            except Exception as e:
                messagebox.showerror("读取标准失败", f"无法找到'加分预览'工作表：{e}")
                self.toggle_score_inquiry()
                return

        cols = list(self.inquiry_data.columns)
        self.tree_week["columns"] = cols
        for col in cols:
            display_text = "" if "Unnamed" in str(col) else col
            self.tree_week.heading(col, text=display_text)
            self.tree_week.column(col, width=80, anchor="center")

        for item in self.tree_week.get_children(): self.tree_week.delete(item)
        for _, row in self.inquiry_data.iterrows():
            self.tree_week.insert("", "end", values=list(row))

    def setup_month_view(self):
        step = SectionCard(self.frame_month, text="生成步骤")
        step.pack(fill="x")

        btn_frame = ttk.Frame(step)
        btn_frame.pack(fill="x", pady=(12, 4))
        btn_frame.columnconfigure(1, weight=1)

        self.btn_scan_weeks = ttk.Button(btn_frame, text="步骤 1 · 读取本月周纪实表", style="Primary.TButton",
                                        command=self.scan_weeks_for_month)
        self.btn_scan_weeks.grid(row=0, column=0, sticky="w")
        self.lbl_scan_status = ttk.Label(btn_frame, text="未读取", style="Danger.TLabel")
        self.lbl_scan_status.grid(row=0, column=1, sticky="w", padx=(14, 0))

        self.btn_load_last = ttk.Button(btn_frame, text="步骤 2 · 导入上月月度总表",
                                       command=self.load_last_month_file)
        self.btn_load_last.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.lbl_last_status = ttk.Label(btn_frame, text="未导入（按 0 分计算）", style="Danger.TLabel")
        self.lbl_last_status.grid(row=1, column=1, sticky="w", padx=(14, 0), pady=(10, 0))

        ttk.Checkbutton(btn_frame, text="首月模式（无上月数据，每人底分按 80 计）",
                        variable=self.is_first_month,
                        command=self.check_month_ready).grid(row=2, column=0, columnspan=2,
                                                             sticky="w", pady=(14, 2))

        foot = ttk.Frame(self.frame_month)
        foot.pack(side="bottom", fill="x", pady=(14, 0))
        self.btn_generate_month = ttk.Button(foot, text="请先完成上方两个步骤", style="Primary.TButton",
                                             state="disabled", command=self.save_month_data)
        self.btn_generate_month.pack(side="left")
        ttk.Button(foot, text="截图", command=lambda: self.screenshot_tree(
            self.tree_month, f"月度总表截图_{self.class_name}{self.current_month.get()}月")).pack(side="left", padx=(12, 0))
        ttk.Label(foot, text="按当前周期生成月度总表", style="Muted.TLabel").pack(side="left", padx=(14, 0))

        card = SectionCard(self.frame_month, text="本月分数累加预览")
        card.pack(fill="both", expand=True, pady=(12, 0))

        table = ttk.Frame(card)
        table.pack(fill="both", expand=True, pady=(12, 0))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        cols = ("学号", "姓名", "上月累计底分", "本月新增分", "本月最终总分")
        self.tree_month = ttk.Treeview(table, columns=cols, show="headings", height=12)
        for col in cols:
            self.tree_month.heading(col, text=col)
            self.tree_month.column(col, anchor="center")

        scr = ttk.Scrollbar(table, orient="vertical", command=self.tree_month.yview)
        self.tree_month.configure(yscrollcommand=scr.set)
        self.tree_month.grid(row=0, column=0, sticky="nsew")
        scr.grid(row=0, column=1, sticky="ns")

    # ---------------- 核心弹窗 ----------------
    def setup_dorm_view(self):
        """寝室统计：按寝室星级一键批量加减分"""

        # ========== 顶部信息条 ==========
        bar = ttk.Frame(self.frame_dorm)
        bar.pack(fill="x", pady=(0, 12))

        self.lbl_dorm_week = ttk.Label(bar, text="周次", style="Strong.TLabel")
        self.lbl_dorm_week.pack(side="left")
        dorm_week_cb = ttk.Combobox(bar, textvariable=self.current_week,
                                    values=["1", "2", "3", "4", "5"], width=4,
                                    state="readonly")
        dorm_week_cb.pack(side="left", padx=(6, 4))
        dorm_week_cb.bind("<<ComboboxSelected>>", lambda e: self.on_dorm_week_change())
        ttk.Label(bar, text="周", style="Muted.TLabel").pack(side="left")
        self.lbl_dorm_rule = ttk.Label(bar, text="", style="Muted.TLabel")
        self.lbl_dorm_rule.pack(side="left", padx=(14, 0))

        ttk.Button(bar, text="返回周表", command=self.show_week_view).pack(side="right")

        # AI 写入的寝室记录不再单独成区：直接内嵌在对应寝室卡片下方
        # （青色 = AI 写入，橙色 = AI 按相似度匹配，点一条可直接修改），
        # 由 _rebuild_dorm_ai_rows 就地刷新。
        self.dorm_stars = {}
        self.dorm_controls = {}
        self.dorm_ai_boxes = {}      # 寝室号 → AI 记录容器

        # ========== 主体：寝室星级评定 ==========
        left_card = SectionCard(self.frame_dorm, text="寝室星级评定（青=已报 红=未发；按次数计分，应用打分后清零）")
        self.dorm_left_card = left_card
        left_card.pack(fill="both", expand=True)

        # 周 / 月 规则切换
        period_row = ttk.Frame(left_card)
        period_row.pack(fill="x", pady=(12, 8))
        ttk.Label(period_row, text="评定周期", style="Muted.TLabel").pack(side="left", padx=(0, 10))
        SegmentedControl(period_row, [("周", "按周"), ("月", "按月")],
                         self.dorm_period_var, command=lambda v: self._sync_dorm_rule_label()).pack(side="left")

        # 操作按钮先按底部占位，避免被寝室列表挤出可视区
        act = ttk.Frame(left_card)
        act.pack(side="bottom", fill="x", pady=(12, 0))
        ttk.Label(act, text="应用后自动生成星级评定表和本周周纪实表",
                  style="Muted.TLabel").pack(side="right")
        ttk.Button(act, text="读取星级", command=self.load_dorm_stars).pack(side="left")
        ttk.Button(act, text="应用打分", style="Primary.TButton",
                   command=self.apply_dorm_scores).pack(side="left", padx=(8, 0))

        list_wrap = ttk.Frame(left_card)
        list_wrap.pack(fill="both", expand=True)

        cv = tk.Canvas(list_wrap, highlightthickness=0, borderwidth=0)
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=cv.yview)
        sf = ttk.Frame(cv)
        win = cv.create_window((0, 0), window=sf, anchor="nw")
        sf.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
        cv.configure(yscrollcommand=sb.set)
        cv.bind("<Enter>", lambda e: cv.bind_all("<MouseWheel>",
                 lambda ev: cv.yview_scroll(int(-1 * (ev.delta / 120)), "units")))
        cv.bind("<Leave>", lambda e: cv.unbind_all("<MouseWheel>"))
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        from tkwebui.core import PALETTE
        cv.configure(bg=PALETTE["bg"])

        self._sync_dorm_rule_label()

        columns = 3
        for col in range(columns):
            sf.columnconfigure(col, weight=1, uniform="dorm")

        self.dorm_ai_boxes.clear()
        self.dorm_labels = {}
        self.dorm_inspections = {}      # 寝室号 → [星级,...] 本周期收到的检查结果（按次数计分）
        self._paint_dorm_states()
        for idx, dorm in enumerate(sorted(set(self.dorm_map.values()))):
            row, col = divmod(idx, columns)
            cell = ttk.Frame(sf)
            cell.grid(row=row, column=col, sticky="w", padx=(0, 18), pady=6)
            top_row = ttk.Frame(cell)
            top_row.pack(fill="x")
            lbl_dorm = ttk.Label(top_row, text=dorm, width=6, anchor="w")
            lbl_dorm.pack(side="left")
            self.dorm_labels[dorm] = lbl_dorm
            var = tk.IntVar(value=4)
            self.dorm_stars[dorm] = 4
            seg = SegmentedControl(top_row, [(3, "3星"), (4, "4星"), (5, "5星")], var,
                                   command=lambda star, d=dorm: self._set_dorm_star(d, star))
            seg.pack(side="left")
            self.dorm_controls[dorm] = seg
            ai_box = ttk.Frame(cell)
            ai_box.pack(fill="x")
            self.dorm_ai_boxes[dorm] = ai_box

        if not self.dorm_stars:
            ttk.Label(sf, wraplength=560, style="Muted.TLabel", text=(
                "名单里还没有寝室信息——用 设置 → 添加模板 导入含寝室号的名单模板后，\n"
                "这里会显示每个寝室的星级卡片；AI 写入的寝室记录也会直接挂在自己寝室下面。"
            )).grid(row=0, column=0, columnspan=columns, sticky="w", padx=12, pady=28)

    # ---------------- 第四页：AI 连接（MCP） ----------------
    def setup_ai_view(self):
        # ========== 顶部：状态 + 三个按钮 ==========
        bar = ttk.Frame(self.frame_ai)
        bar.pack(fill="x", pady=(0, 12))
        self.ai_status = ttk.Label(bar, text="", style="Muted.TLabel")
        self.ai_status.pack(side="left")
        # 队列里有别的周期的批次时，这里给个一键跳过去的按钮
        self.btn_ai_jump = ttk.Button(bar, text="跳到队列周期",
                                      command=self._jump_to_pending_cycle)
        # MCP 配置/接入说明/打开文件夹 收进了左下角「设置」对话框，这里只留状态
        ttk.Label(bar, text="接入与配置入口在左下角「设置」", style="Muted.TLabel").pack(side="right")

        # ========== AI 写入的记录（逐条，可改、可驳回） ==========
        log_card = SectionCard(self.frame_ai, text="AI 写入的记录（当前周期，点一行可修改）")
        log_card.pack(fill="both", expand=True)

        # 表头（列宽固定，末尾留一个弹性空白列：
        # 既保证表头和数据对齐，也避免「驳回」按钮被推到窗口最右边）
        head = ttk.Frame(log_card)
        head.pack(fill="x", pady=(10, 0))
        for col, (text, anchor) in enumerate((
                ("学号后两位", "w"), ("姓名", "w"),
                ("加减分", "center"), ("明细", "w"), ("", "center"))):
            ttk.Label(head, text=text, style="Muted.TLabel", anchor=anchor).grid(
                row=0, column=col, sticky="ew", padx=(12 if col == 0 else 6, 6))
            head.grid_columnconfigure(col, minsize=self.AI_COL_WIDTHS[col])
        head.grid_columnconfigure(len(self.AI_COL_WIDTHS), weight=1)

        ttk.Separator(log_card, orient="horizontal").pack(fill="x", pady=(6, 0))

        # 近似匹配图例：平时不显示，出现橙色行时才插入到列表上方
        self.ai_legend = ttk.Label(
            log_card, foreground=WARN_C,
            text="橙色行 = 姓名是按相似度猜的，请核对「原文」那几个人对不对")
        self.ai_legend_auto = ttk.Label(
            log_card, foreground=ACCENT,
            text="青色分值 = 你没说分数，程序按模板「加分预览」自动算的，点该行可改")

        # 可滚动的行容器
        list_wrap = ttk.Frame(log_card)
        self.ai_list_wrap = list_wrap
        list_wrap.pack(fill="both", expand=True, pady=(4, 0))
        from tkwebui.core import PALETTE
        self.ai_canvas = tk.Canvas(list_wrap, highlightthickness=0, borderwidth=0,
                                   bg=PALETTE["bg"])
        ai_bar = ttk.Scrollbar(list_wrap, orient="vertical", command=self.ai_canvas.yview)
        self.ai_rows = ttk.Frame(self.ai_canvas)
        ai_win = self.ai_canvas.create_window((0, 0), window=self.ai_rows, anchor="nw")
        self.ai_rows.bind("<Configure>", lambda e: self.ai_canvas.configure(
            scrollregion=self.ai_canvas.bbox("all")))
        self.ai_canvas.bind("<Configure>", lambda e: self.ai_canvas.itemconfigure(
            ai_win, width=e.width))
        self.ai_canvas.configure(yscrollcommand=ai_bar.set)
        self.ai_canvas.pack(side="left", fill="both", expand=True)
        ai_bar.pack(side="right", fill="y")
        self.ai_rows.grid_columnconfigure(0, weight=1)

        self.ai_empty = ttk.Label(log_card, style="Muted.TLabel",
                                  text="还没有收到 AI 记录。接入后，AI 写入的每条记录都会出现在这里。")
        self.ai_empty.pack(side="bottom", anchor="w", pady=(6, 0))

        self.refresh_ai_view()

    def _ai_pack_dir(self):
        return os.path.join(self._base_dir(), "AI对接")

    def _ai_config_path(self):
        return os.path.join(self._ai_pack_dir(), "MCP配置.json")

    def _open_ai_folder(self):
        folder = self._ai_pack_dir()
        if not os.path.isdir(folder):
            return messagebox.showwarning("找不到目录", f"没有找到 AI 对接目录：\n{folder}")
        self._open_folder(folder)

    def _open_ai_config_file(self):
        """在软件内直接查看配置内容。

        不能用 os.startfile 打开 .json：文件关联可能被别的程序抢走
        （这台机器上 .json 关联到了 VIA.exe），点了只会弹那个程序，
        看起来就像按钮没反应。改成内置查看窗，不依赖任何外部程序。
        """
        path = self._ai_config_path()
        if not os.path.exists(path):
            return messagebox.showwarning("找不到配置", f"没有找到配置文件：\n{path}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as exc:
            return messagebox.showerror("读取失败", f"{path}\n\n{exc}")
        self._show_text_dialog("MCP 配置内容", content, path)

    def _regen_mcp_config(self):
        r"""按当前运行方式重新生成 AI对接\MCP配置.json。

        挪动安装目录 / 换机器 / 重装后，配置里的绝对路径会失效，
        点这个按钮就能按实际路径重写，不用手改 JSON。
        """
        pack_dir = self._ai_pack_dir()
        if getattr(sys, "frozen", False):
            cmd = os.path.join(pack_dir, "mcp_server.exe")
            if not os.path.exists(cmd):
                return messagebox.showwarning(
                    "缺少组件", f"没有找到：\n{cmd}\n\n请用安装包重新安装（修复）一次。")
            payload = {"mcpServers": {"deyufen": {"command": cmd, "args": []}}}
        else:
            script = os.path.join(pack_dir, "mcp_server.py")
            if not os.path.exists(script):
                return messagebox.showwarning("缺少组件", f"没有找到：\n{script}")
            payload = {"mcpServers": {"deyufen": {
                "command": sys.executable, "args": [script]}}}
        path = self._ai_config_path()
        try:
            os.makedirs(pack_dir, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            return messagebox.showerror("写入失败", f"{path}\n\n{exc}")
        messagebox.showinfo(
            "完成", f"MCP 配置已重新生成：\n{path}\n\n"
            f"在 AI 客户端里重新添加（或刷新）该 MCP 服务即可。")

    def _show_text_dialog(self, title, content, path=None):
        """只读文本查看窗，带「复制全部」和「用记事本打开」。"""
        from tkwebui.core import PALETTE

        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.configure(bg=PALETTE["bg"])
        self._place_dialog(win, 720, 430)

        outer = ttk.Frame(win, padding=14)
        outer.pack(fill="both", expand=True)

        # 按钮行先按底部占位，否则会被 Text 的默认高度挤出窗口
        btns = ttk.Frame(outer)
        btns.pack(side="bottom", fill="x", pady=(10, 0))

        head = ttk.Frame(outer)
        head.pack(side="top", fill="x", pady=(0, 8))
        ttk.Label(head, text=title, style="Strong.TLabel").pack(side="left")
        if path:
            ttk.Label(head, text=os.path.dirname(path), style="Muted.TLabel").pack(side="right")

        text_area = ttk.Frame(outer)
        text_area.pack(fill="both", expand=True)

        body = tk.Text(text_area, wrap="none", height=1, font=("Consolas", 10),
                       relief="flat", borderwidth=1, highlightthickness=0,
                       padx=10, pady=8)
        body.insert("1.0", content)
        body.configure(state="disabled")

        vbar = ttk.Scrollbar(text_area, orient="vertical", command=body.yview)
        hbar = ttk.Scrollbar(text_area, orient="horizontal", command=body.xview)
        body.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        body.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        text_area.rowconfigure(0, weight=1)
        text_area.columnconfigure(0, weight=1)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 0))

        def copy_all():
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                notify(self.root, "已复制到剪贴板", title="已复制", kind="success")
            except Exception as exc:
                messagebox.showerror("复制失败", str(exc))

        ttk.Button(btns, text="复制全部", style="Primary.TButton",
                   command=copy_all).pack(side="right")
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right", padx=(0, 8))
        if path:
            ttk.Button(btns, text="用记事本打开", command=lambda p=path: self._open_with_notepad(p)
                       ).pack(side="left")

        win.bind("<Escape>", lambda e: win.destroy())
        win.focus_set()

    def _open_with_notepad(self, path):
        """用记事本打开（不依赖文件关联，任何 Windows 都有）。"""
        try:
            subprocess.Popen(["notepad.exe", path])
            return True
        except Exception as exc:
            messagebox.showerror("打不开", f"{path}\n\n{exc}")
            return False

    def _copy_ai_folder(self):
        folder = self._ai_pack_dir()
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(folder)
            notify(self.root, "文件夹路径已复制到剪贴板", title="已复制", kind="success")
        except Exception as exc:
            messagebox.showerror("复制失败", str(exc))

    def _open_ai_doc(self):
        """接入说明：在软件内查看（不依赖文件关联，.md 可能被别的程序抢走）。"""
        path = os.path.join(self._ai_pack_dir(), "使用说明.md")
        if not os.path.exists(path):
            return messagebox.showwarning("找不到文档", f"没有找到使用说明：\n{path}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as exc:
            return messagebox.showerror("读取失败", f"{path}\n\n{exc}")
        self._show_text_dialog("接入说明", content, path)

    def _copy_ai_config(self):
        path = self._ai_config_path()
        if not os.path.exists(path):
            return messagebox.showwarning("找不到配置", f"没有找到配置文件：\n{path}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            notify(self.root, "MCP 配置已复制到剪贴板", title="已复制", kind="success")
        except Exception as exc:
            messagebox.showerror("复制失败", str(exc))

    def refresh_ai_view(self):
        """刷新 AI 连接页：顶部状态 + 逐条记录列表。"""
        if not hasattr(self, "ai_status"):
            return
        self._refresh_ai_status_only()
        self._rebuild_ai_rows()

    def _collect_ai_rows(self, only_dorm=False):
        """把当前周期里 AI 写入的记录逐条取出来，按写入时间排。

        only_dorm=True 时只取寝室相关的记录（寝室统计页用）。
        """
        rows = []
        for sid, info in self.week_data.items():
            for idx, rec in enumerate(info.get('records', [])):
                if not is_ai_record(rec):
                    continue
                if only_dorm and not is_dorm_reason(rec[1]):
                    continue
                meta = rec[2]
                fuzzy = bool(meta.get("fuzzy"))
                auto = bool(meta.get("auto"))
                rows.append({
                    "sid": sid, "idx": idx,
                    "tail": sid[-2:],
                    "name": info.get('name', ''),
                    "score": rec[0],
                    "reason": rec[1],
                    "time": str(meta.get("time", "")),
                    "fuzzy": fuzzy,
                    "input": str(meta.get("input", "")),
                    "rate": meta.get("rate", 0),
                    "auto": auto,
                    "auto_from": str(meta.get("auto_from", "")),
                })
        rows.sort(key=lambda r: (r["time"], r["sid"], r["idx"]))
        return rows

    def _render_ai_rows(self, container, wrap, legends, empty, rows, empty_text=""):
        """按统一格式画 AI 记录行（AI 连接页和寝室统计页共用）。

        container：放行的 Frame；wrap：可滚动的行容器（图例插在它上面）；
        legends：(橙色图例, 青色图例)，没有该类记录时隐藏。
        近似匹配的行整行橙色，按模板自动取分的分值用青色。
        """
        for child in container.winfo_children():
            child.destroy()

        fuzzy_rows = [r for r in rows if r["fuzzy"]]
        auto_rows = [r for r in rows if r["auto"]]
        # 图例只在确实有该类记录时出现。
        # 每次都先收起再按顺序放回：winfo_ismapped 在卡片刚 pack 时不可靠，
        # 靠它判断会导致第二条图例有时不显示。
        for legend, shows in zip(legends, (fuzzy_rows, auto_rows)):
            if legend is None:
                continue
            legend.pack_forget()
            if shows:
                legend.pack(fill="x", pady=(6, 0), before=wrap)

        if not rows:
            if empty_text:
                empty.config(text=empty_text)
            empty.pack(side="bottom", anchor="w", pady=(6, 0))
            return
        empty.pack_forget()

        for line, row in enumerate(rows):
            item = ttk.Frame(container)
            item.grid(row=line, column=0, sticky="ew", pady=1)
            for col in range(len(self.AI_COL_WIDTHS)):
                item.grid_columnconfigure(col, minsize=self.AI_COL_WIDTHS[col])
            item.grid_columnconfigure(len(self.AI_COL_WIDTHS), weight=1)

            score_text = self._fmt_score(row["score"])
            reason_text = str(row["reason"])
            tags = []
            if row["fuzzy"]:
                pct = int(round(float(row["rate"] or 0) * 100))
                tags.append(f"原文 {row['input']} {pct}%")
            if row["auto"]:
                source = str(row["auto_from"])
                head = source.split("·")[0]
                # 规则名已经写在原因里就不重复（如「五星寝室」+「五星寝室·周·成员」）
                tags.append("模板取分" if head and head in str(row["reason"])
                            else f"模板 {source}")
            if tags:
                reason_text = f"{reason_text}（{' · '.join(tags)}）"
            labels = [
                (0, row["tail"], "Muted.TLabel", "w"),
                (1, row["name"], "TLabel", "w"),
                (2, score_text, "TLabel", "center"),
                (3, reason_text, "Muted.TLabel", "w"),
            ]
            for col, text, style, anchor in labels:
                opts = {}
                if row["fuzzy"]:
                    opts["foreground"] = WARN_C
                elif col == 2 and row["auto"]:
                    # 分值不是用户说的，是程序按模板查的，用青色提示核对
                    opts["foreground"] = ACCENT
                lab = ttk.Label(item, text=text, style=style, anchor=anchor, **opts)
                lab.grid(row=0, column=col, sticky="ew",
                         padx=(12 if col == 0 else 6, 6), pady=4)
                # 点任意一列都能进入修改
                lab.bind("<Button-1>",
                         lambda e, s=row["sid"], i=row["idx"]: self._edit_ai_record(s, i))
            ttk.Button(item, text="驳回", width=8,
                       command=lambda s=row["sid"], i=row["idx"]: self._reject_ai_record(s, i)
                       ).grid(row=0, column=4, padx=(6, 12), pady=2)

    def _rebuild_ai_rows(self):
        """重画 AI 连接页的记录列表。"""
        if not hasattr(self, "ai_rows"):
            return
        self._render_ai_rows(self.ai_rows, self.ai_list_wrap,
                             (self.ai_legend, self.ai_legend_auto), self.ai_empty,
                             self._collect_ai_rows())

    def _paint_dorm_state(self, dorm):
        """寝室号颜色 = 状态：青=本周期已报过结果；红=未发；黑=手动/已应用清零。"""
        lbl = getattr(self, "dorm_labels", {}).get(dorm)
        if lbl is None or not lbl.winfo_exists():
            return
        has_report = bool(getattr(self, "dorm_inspections", {}).get(dorm))
        ai_set = dorm in getattr(self, "dorm_ai_set", set())
        applied = dorm in getattr(self, "dorm_applied", set())
        if has_report or ai_set:
            lbl.config(foreground=ACCENT)          # 青 = 已报，待应用
        elif applied:
            lbl.config(foreground=INK)             # 黑 = 已应用清零
        else:
            lbl.config(foreground="#DC2626")       # 红 = 未发

    def _paint_dorm_states(self):
        for dorm in getattr(self, "dorm_labels", {}):
            self._paint_dorm_state(dorm)

    def on_dorm_week_change(self):
        """寝室页切周次：未保存先问，然后加载该周数据并恢复星级/状态。"""
        if not self._confirm_discard("切换周次"):
            self._restore_cycle()
            return
        self.load_week_data()
        self._sync_dorm_stars()
        self.dorm_inspections = {}
        self.dorm_ai_set = set()
        self.dorm_applied = set()
        self.lbl_dorm_week.config(text="周次")
        self._rebuild_dorm_ai_rows()
        self._sync_dorm_rule_label()

    def _rebuild_dorm_ai_rows(self):
        """把 AI 写入的寝室记录就地合并进各寝室卡片（不再单独成区）。

        每条记录挂在自己寝室的星级卡片下方：
          青色 = AI 写入；橙色 = AI 按姓名相似度匹配（请核对）。
        点一条可直接进入该记录的修改框。
        本周期没收到检查结果的寝室，滑块后面挂红色「未发」标提醒。
        """
        if not hasattr(self, "dorm_ai_boxes"):
            return
        per_dorm = {}
        for r in self._collect_ai_rows(only_dorm=True):
            dorm = self.week_data.get(r["sid"], {}).get("dorm", "")
            per_dorm.setdefault(dorm, []).append(r)

        for dorm, box in self.dorm_ai_boxes.items():
            for w in box.winfo_children():
                w.destroy()
            recs = per_dorm.get(dorm, [])
            self._paint_dorm_state(dorm)
            for r in recs:
                color = WARN_C if r["fuzzy"] else ACCENT
                txt = f"AI·{r['name']}({r['tail']})  {r['reason']}  {signed_text(r['score'])}"
                item = ttk.Label(box, text=txt, foreground=color,
                                 font=("Microsoft YaHei UI", 9), cursor="hand2")
                item.pack(anchor="w")
                item.bind("<Button-1>",
                          lambda e, sid=r["sid"], idx=r["idx"]: self._edit_ai_record(sid, idx))

    def _refresh_after_ai_change(self):
        """AI 记录被改/被驳回后，同步刷新所有界面。"""
        self.unsaved = True
        self._mark_month_stale()
        if not self.showing_inquiry:
            self.refresh_week_tree()
        self._rebuild_ai_rows()
        self._rebuild_dorm_ai_rows()
        self._refresh_ai_status_only()

    def _pending_batches(self):
        """读 inbox 里排队批次的目标周期。

        返回 [{"cycle": (年,月,周), "count": 条数, "name": 文件名}]，按文件名（时间）排序。
        """
        inbox = os.path.join(self._bridge_dir(), "inbox")
        batches = []
        if not os.path.isdir(inbox):
            return batches
        try:
            names = sorted(os.listdir(inbox))
        except Exception:
            return batches
        for name in names:
            if not name.endswith(".json"):
                continue
            try:
                with open(os.path.join(inbox, name), "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                continue
            cycle = (str(data.get("year", "")), str(data.get("month", "")),
                     str(data.get("week", "")))
            if data.get("type") == "dorm_stars":
                count = len(data.get("stars") or {})
                label = "寝室检查"
            else:
                count = len(data.get("records") or [])
                label = "记录"
            batches.append({"cycle": cycle, "count": count,
                            "label": label, "name": name})
        return batches

    def _refresh_ai_status_only(self):
        """只刷新顶部状态行，不动列表。"""
        if not hasattr(self, "ai_status"):
            return
        config_ok = os.path.exists(self._ai_config_path())
        batches = self._pending_batches()
        now_count = len(self._collect_ai_rows())
        parts = [f"配置文件：{'已就绪' if config_ok else '缺失'}",
                 f"当前周期 AI 记录 {now_count} 条"]
        elsewhere = [b for b in batches if b["cycle"] != self._cycle_key()]
        if batches:
            detail = "、".join(
                f"{b['cycle'][0]}年{b['cycle'][1]}月第{b['cycle'][2]}周"
                f"（{b['count']} 条{b['label']}）" for b in batches)
            if elsewhere:
                parts.append(f"队列待导入：{detail}，切到对应周期会自动填入")
            else:
                parts.append(f"队列待导入：{detail}")
        else:
            parts.append("队列为空")
        self.ai_status.config(text="  ·  ".join(parts),
                              foreground=WARN_C if elsewhere else
                              (OK_C if config_ok else WARN_C))

        # 「跳到队列周期」按钮：只在队列里有别的周期时出现
        btn = getattr(self, "btn_ai_jump", None)
        if btn is None:
            return
        if elsewhere:
            first = elsewhere[0]["cycle"]
            btn.config(text=f"切到 {first[0]}年{first[1]}月第{first[2]}周")
            if not btn.winfo_ismapped():
                btn.pack(side="left", padx=(10, 0))
        elif btn.winfo_ismapped():
            btn.pack_forget()

    def _jump_to_pending_cycle(self):
        """把软件切到队列批次指向的周期（走正常的未保存询问流程）。"""
        batches = self._pending_batches()
        targets = [b["cycle"] for b in batches if b["cycle"] != self._cycle_key()]
        if not targets:
            return
        year, month, week = targets[0]
        if not (year and month and week):
            messagebox.showwarning("队列里这批没写周期",
                                   "批次文件缺少 year/month/week，不会自动填入。\n"
                                   "让 AI 重新发一次带周期的批次。")
            return
        self._suspend_cycle_cb = True
        try:
            self.current_year.set(year)
            self.current_month.set(month)
            self.current_week.set(week)
        finally:
            self._suspend_cycle_cb = False
        self._switch_cycle(f"切到 {year}年{month}月第{week}周")

    def _edit_ai_record(self, sid, idx):
        """点一行 → 改这条记录的分值和明细。"""
        records = self.week_data.get(sid, {}).get('records', [])
        if idx >= len(records):
            return
        rec = records[idx]
        info = self.week_data[sid]

        win = tk.Toplevel(self.root)
        win.title(f"修改记录 · {info.get('name', '')}")
        win.transient(self.root)
        self._place_dialog(win, 430, 210)

        outer = ttk.Frame(win, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=f"{sid}  {info.get('name', '')}",
                  style="Strong.TLabel").pack(anchor="w")

        score_var = tk.StringVar(value=self._fmt_score(rec[0]))
        reason_var = tk.StringVar(value=str(rec[1]))

        form = ttk.Frame(outer)
        form.pack(fill="x", pady=(12, 0))
        ttk.Label(form, text="分值", style="Muted.TLabel", width=6).grid(row=0, column=0, sticky="w")
        score_entry = ttk.Entry(form, textvariable=score_var, width=12)
        score_entry.grid(row=0, column=1, sticky="w")
        ttk.Label(form, text="（正数加分，负数扣分）", style="Muted.TLabel").grid(
            row=0, column=2, sticky="w", padx=(8, 0))
        ttk.Label(form, text="明细", style="Muted.TLabel", width=6).grid(
            row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(form, textvariable=reason_var, width=34).grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        btns = ttk.Frame(outer)
        btns.pack(side="bottom", fill="x", pady=(16, 0))

        def save():
            try:
                new_score = self._eval_score(score_var.get())
            except (ValueError, SyntaxError):
                messagebox.showwarning("分值不对", "请输入数字，例如 1、-0.5、2", parent=win)
                return
            new_reason = reason_var.get().strip() or "AI 记录"
            self._push_undo()
            cur = self.week_data.get(sid, {}).get('records', [])
            if idx >= len(cur):
                win.destroy()
                return
            old = cur[idx]
            cur[idx] = [new_score, new_reason] + list(old[2:])
            self.last_edited_sid = sid
            win.destroy()
            self._refresh_after_ai_change()
            notify(self.root, f"{info.get('name', '')} 已改为 {self._fmt_score(new_score)}",
                   title="已修改", kind="success")

        ttk.Button(btns, text="保存", style="Primary.TButton",
                   command=save).pack(side="right")
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=(0, 8))
        win.bind("<Return>", lambda e: save())
        win.bind("<Escape>", lambda e: win.destroy())
        score_entry.focus_set()

    def _reject_ai_record(self, sid, idx):
        """驳回：把这条记录从本周删掉（可撤回）。"""
        records = self.week_data.get(sid, {}).get('records', [])
        if idx >= len(records):
            return
        rec = records[idx]
        name = self.week_data[sid].get('name', '')
        preview = f"{name} {self._fmt_score(rec[0])} {rec[1]}"
        if not messagebox.askyesno("驳回这条记录",
                                   f"从本周删除这条记录？\n\n{preview}\n\n"
                                   "删除后可以点「撤回」恢复。",
                                   parent=self.root):
            return
        self._push_undo()
        del records[idx]
        self.last_edited_sids.discard(sid)
        self._refresh_after_ai_change()
        notify(self.root, f"已驳回：{preview}", title="已删除", kind="warn")

    def _sync_dorm_rule_label(self):
        """规则从模板的「寝室加分规则」表读，模板改了这里跟着变。"""
        rules = self._load_dorm_rules()
        period = "月" if self.dorm_period_var.get() == "月" else "周"
        star_names = {5: "5星", 4: "4星", 3: "3星"}
        parts = []
        for star in (5, 4, 3):
            values = rules.get(period, {}).get(star)
            if not values:
                continue
            member, leader = values.get('成员', 0.0), values.get('室长', 0.0)
            if abs(member) < 1e-9 and abs(leader) < 1e-9:
                continue      # 四星不加减，不必显示
            name = star_names.get(star, f"{star}星")
            if abs(member - leader) < 1e-9:
                parts.append(f"{name} {signed_text(member)}")
            else:
                parts.append(f"{name} 成员 {signed_text(member)} / 室长 {signed_text(leader)}")
        label = "月度" if period == "月" else "周"
        self.lbl_dorm_rule.config(text=f"{label}规则：" + "  |  ".join(parts))

    def _set_dorm_star(self, dorm, star, origin="user"):
        """设置寝室星级。既可由分段控件回调，也可由代码直接调用。

        origin="ai" 时把该寝室记入 dorm_ai_set（青色标记）；
        用户手动改滑块会清除该寝室的 AI 标记。
        """
        self.dorm_stars[dorm] = star
        if origin == "ai":
            self.dorm_ai_set.add(dorm)
            self.dorm_applied.discard(dorm)
            # 按次数计分：每收到一次结果累计一次（应用打分后清零）
            self.dorm_inspections.setdefault(dorm, []).append(star)
        else:
            self.dorm_ai_set.discard(dorm)
            self.dorm_inspections.pop(dorm, None)   # 手动改滑块 = 覆盖 AI 计数
        seg = getattr(self, "dorm_controls", {}).get(dorm)
        if seg is not None and int(seg.variable.get()) != star:
            seg.variable.set(star)   # 触发 SegmentedControl 自行同步高亮

    STAR_TEXT_MAP = {"三": 3, "四": 4, "五": 5, "三星": 3, "四星": 4, "五星": 5,
                     "3星": 3, "4星": 4, "5星": 5}

    def _load_dorm_rules(self):
        """从名单模板的「寝室加分规则」表读周/月两套规则。

        表结构：A 列星级，B/C 周（成员/室长），D/E 月（成员/室长）。
        读不到或格式不符时退回内置默认值，保证功能可用。
        """
        cached = getattr(self, "_dorm_rules_cache", None)
        if cached is not None:
            return cached

        rules = {
            '周': {5: {'成员': 1.0, '室长': 1.0}, 4: {'成员': 0.0, '室长': 0.0},
                   3: {'成员': -1.0, '室长': -1.5}},
            '月': {5: {'成员': 2.0, '室长': 2.5}, 4: {'成员': 0.0, '室长': 0.0},
                   3: {'成员': -2.0, '室长': -4.0}},
        }
        wb = None
        try:
            wb = openpyxl.load_workbook(self._template_path(), data_only=True)
            if "寝室加分规则" in wb.sheetnames:
                ws = wb["寝室加分规则"]
                for r in range(2, ws.max_row + 1):
                    label = ws.cell(row=r, column=1).value
                    if label is None:
                        continue
                    star = self.STAR_TEXT_MAP.get(str(label).strip())
                    if star is None:
                        continue
                    for period, col_member, col_leader in (('周', 2, 3), ('月', 4, 5)):
                        member = self._to_number(ws.cell(row=r, column=col_member).value)
                        leader = self._to_number(ws.cell(row=r, column=col_leader).value)
                        if member is not None:
                            rules[period][star]['成员'] = member
                        if leader is not None:
                            rules[period][star]['室长'] = leader
        except Exception:
            pass
        finally:
            if wb:
                wb.close()

        self._dorm_rules_cache = rules
        return rules

    def _pick_student(self, keyword, matches):
        """匹配到多人时的选择弹窗：表格列出，双击或回车即选。返回 sid 或 None。"""
        result = [None]
        top = tk.Toplevel(self.root)
        top.title("选择学生")
        top.grab_set()
        self._place_dialog(top, 460, 420)

        wrap = ttk.Frame(top, padding=22)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text=f"「{keyword}」匹配到 {len(matches)} 人", style="Strong.TLabel").pack(anchor="w")
        ttk.Label(wrap, text="双击或选中后按确定", style="Muted.TLabel").pack(anchor="w", pady=(4, 12))

        holder = ttk.Frame(wrap)
        holder.pack(fill="both", expand=True)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)

        tree = ttk.Treeview(holder, columns=("name", "sid", "dorm"), show="headings", height=8)
        for cid, title, w in (("name", "姓名", 110), ("sid", "学号", 150), ("dorm", "寝室", 100)):
            tree.heading(cid, text=title)
            tree.column(cid, width=w, anchor="center")
        for sid, info in matches:
            tree.insert("", "end", iid=sid, values=(info['name'], sid, info['dorm']))

        scr = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scr.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scr.grid(row=0, column=1, sticky="ns")

        def choose(event=None):
            sel = tree.selection()
            if sel:
                result[0] = sel[0]
                top.destroy()

        tree.bind("<Double-Button-1>", choose)
        tree.bind("<Return>", choose)

        btns = ttk.Frame(wrap)
        btns.pack(fill="x", pady=(16, 0))
        ttk.Button(btns, text="确定", style="Primary.TButton", command=choose).pack(side="right")
        ttk.Button(btns, text="取消", command=top.destroy).pack(side="right", padx=(0, 8))

        if matches:
            tree.selection_set(matches[0][0])
            tree.focus(matches[0][0])

        self.root.wait_window(top)
        return result[0]

    def ask_retry_save(self, file_path):
        top = tk.Toplevel(self.root)
        top.title("文件被占用")
        top.grab_set()
        self._place_dialog(top, 460, 240)

        wrap = ttk.Frame(top, padding=26)
        wrap.pack(fill="both", expand=True)

        ttk.Label(wrap, text="写入被拒绝", style="Danger.TLabel",
                  font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        ttk.Label(wrap, justify="left", wraplength=390, style="Muted.TLabel",
                  text=f"目标文件「{os.path.basename(file_path)}」正被 Excel 打开。\n"
                       f"请在任务栏彻底关闭该 Excel，然后点击下方按钮重试保存。").pack(anchor="w", pady=(10, 0))

        result = [False]

        def on_retry():
            result[0] = True
            top.destroy()

        def on_cancel():
            top.destroy()

        btns = ttk.Frame(wrap)
        btns.pack(side="bottom", fill="x", pady=(20, 0))
        ttk.Button(btns, text="重试保存", style="Primary.TButton", command=on_retry).pack(side="right")
        ttk.Button(btns, text="取消", command=on_cancel).pack(side="right", padx=(0, 8))

        top.protocol("WM_DELETE_WINDOW", on_cancel)
        self.root.wait_window(top)
        return result[0]

    # ---------------- 数据清洗 ----------------
    def clean_val(self, val):
        if pd.isna(val) or val is None: return ""
        s = str(val).strip()
        if s.endswith(".0"): s = s[:-2]
        if s.lower() == 'nan': return ""
        return s

    def clean_name(self, name):
        n = re.sub(r'\s+', '', str(name))
        return n.replace("奕", "弈")

    def _fmt_score(self, val):
        """整数不显示小数点。对 nan/inf/非数字做兜底，避免 int() 抛异常搞崩整表刷新。"""
        try:
            f = float(val)
        except (TypeError, ValueError):
            return str(val)
        if not math.isfinite(f):
            return "0"
        return str(int(f)) if f == int(f) else str(f)

    def _to_number(self, value):
        """把单元格值转成有限数字；转不了返回 None，由调用方跳过而不是中断整表读取。"""
        if value is None or value == "":
            return None
        try:
            f = float(str(value).strip())
        except (TypeError, ValueError):
            return None
        return f if math.isfinite(f) else None

    def _eval_score(self, text):
        """解析分数，支持 '10+2' 这类简单表达式（安全求值，不用 eval）"""
        return safe_eval_math(text)

    def scrub_ghost_data(self, text):
        if pd.isna(text) or text is None or str(text).lower() == 'nan': return ""
        s = str(text)
        s = re.sub(r'\(\+[0-9.\-]+\)', '', s)
        return s.strip()

    # ---------------- 模板加载 ----------------
    def load_templates(self):
        tpl_path = self._template_path()
        if not os.path.exists(tpl_path):
            messagebox.showerror("致命错误", f"找不到名单模板：\n{tpl_path}")
            self.root.quit()
            return
        try:
            # 重新装载前先清空，避免重复调用时数据叠加
            self.week_data.clear()
            self.base_layout.clear()
            self.month_sum.clear()
            self.last_month_scores.clear()
            self.dorm_map.clear()
            self.dorm_roles.clear()
            self._dorm_rules_cache = None   # 模板可能改了规则，缓存作废

            df_m = pd.read_excel(tpl_path, sheet_name="月表模板", header=None)
            cls_raw = str(df_m.iloc[1, 2]).strip()
            if cls_raw and cls_raw.lower() != 'nan':
                self.class_name = cls_raw if '班' in cls_raw else cls_raw + '班'
            if self.class_name == "未设置班级":
                self.lbl_class.config(text="未设置班级 · 在设置里添加模板",
                                      style="SidebarWarning.TLabel")
            else:
                self.lbl_class.config(text=self.class_name, style="SidebarSuccess.TLabel")

            # 智能提取室长和成员
            self.dorm_roles.clear()
            self.dorm_map.clear()
            df_d = pd.read_excel(tpl_path, sheet_name="寝室名单", header=None)
            for r in range(df_d.shape[0]):
                room = self.clean_val(df_d.iloc[r, 0])
                if '房间' in room or not room: continue

                ppl = str(df_d.iloc[r, 2]).strip() if df_d.shape[1] > 2 else ""
                leader = str(df_d.iloc[r, 6]).strip() if df_d.shape[1] > 6 else ""

                if leader and leader.lower() != 'nan':
                    for l_name in re.split(r'[，, \s]+', leader):
                        cl = self.clean_name(l_name)
                        if cl:
                            self.dorm_map[cl] = room
                            self.dorm_roles[cl] = '室长'

                if ppl and ppl.lower() != 'nan':
                    for n in re.split(r'[，, \s]+', ppl):
                        cn = self.clean_name(n)
                        if cn:
                            self.dorm_map[cn] = room
                            if cn not in self.dorm_roles:
                                self.dorm_roles[cn] = '成员'

            for row in range(1, df_m.shape[0]):
                sid = self.clean_val(df_m.iloc[row, 0])
                name = self.clean_val(df_m.iloc[row, 1])
                if sid and "签名" not in sid and "考评" not in sid:
                    c_name = self.clean_name(name)
                    self.week_data[sid] = {
                        "name": name,
                        "dorm": self.dorm_map.get(c_name, "未知"),
                        "role": self.dorm_roles.get(c_name, "成员"),
                        "records": []
                    }
                    self.month_sum[sid] = 0.0
                    self.last_month_scores[sid] = 0.0

            df_w = pd.read_excel(tpl_path, sheet_name="周表模板", header=None)
            for r in range(3, df_w.shape[0]):
                l_id = self.clean_val(df_w.iloc[r, 0])
                r_id = self.clean_val(df_w.iloc[r, 4]) if df_w.shape[1] > 4 else ""
                valid_l = l_id if l_id in self.week_data else ""
                valid_r = r_id if r_id in self.week_data else ""
                if valid_l or valid_r:
                    self.base_layout.append((valid_l, valid_r))

            # 兜底：周表模板左右栏是空的（空白模板 / 用户手改坏了）时，
            # 按月表名单对半拆分布局，保证还能生成左右两栏的周表
            if not self.base_layout and self.week_data:
                sids = list(self.week_data.keys())
                half = math.ceil(len(sids) / 2)
                for i in range(half):
                    self.base_layout.append(
                        (sids[i], sids[half + i] if half + i < len(sids) else ""))

            self.load_week_data()

            if not self.week_data:
                self.week_status.config(text="名单为空：左下角 设置 → 添加模板，或把名单文件拖进窗口",
                                        foreground=WARN_C)
        except Exception as e:
            messagebox.showerror("模板错误", f"读取失败：{e}")

    def reload_template(self, ask=True):
        if ask and not messagebox.askyesno("确认", "重新加载模板将清除当前所有未保存的加分记录（已生成的文件不受影响），\n确定继续吗？"):
            return
        if not os.path.exists(self._template_path()):
            messagebox.showerror("错误", "找不到名单模板 名单模板.xlsx！")
            return
        # 清除所有旧数据
        self.dorm_map.clear()
        self.dorm_roles.clear()
        self.week_data.clear()
        self.base_layout.clear()
        self.month_sum.clear()
        self.last_month_scores.clear()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.btn_undo.config(state="disabled")
        self.btn_redo.config(state="disabled")
        self.inquiry_data = None
        # 月度汇总的中间状态也一起作废，否则会拿着空数据生成月表
        self.week_scanned = False
        self.last_month_loaded = False
        self.unsaved = False
        # 重新加载
        self.load_templates()
        # 重建寝室打分视图
        for w in self.frame_dorm.winfo_children():
            w.destroy()
        self.setup_dorm_view()
        self._sync_dorm_stars()
        # 刷新预览
        if not self.showing_inquiry:
            self.refresh_week_tree()
        self.refresh_month_tree()
        self.week_status.config(text="模板已重新加载", foreground=OK_C)

    def _template_path(self):
        return os.path.join(self._base_dir(), "名单模板.xlsx")

    # ---------------- 班级名单导入 ----------------
    def show_roster_dialog(self):
        """班级名单导入：粘贴 / 选文件 → 校验预览 → 写回名单模板三张表。

        空白模板没有任何学生，这是新用户录入名单的唯一入口。
        """
        top = tk.Toplevel(self.root)
        top.title("班级名单")
        top.grab_set()
        self._place_dialog(top, 740, 600)
        rows = []

        wrap = ttk.Frame(top, padding=20)
        wrap.pack(fill="both", expand=True)

        head = ttk.Frame(wrap)
        head.pack(fill="x")
        ttk.Label(head, text="班级全称").pack(side="left")
        en_cls = ttk.Entry(head, width=26)
        en_cls.pack(side="left", padx=(8, 0))
        prefill = cls_hint or ("" if self.class_name == "未设置班级" else self.class_name)
        en_cls.insert(0, prefill)
        ttk.Label(head, text="（如：计算机2301，会写进周表表头与月表）",
                  style="Muted.TLabel").pack(side="left", padx=(8, 0))

        ttk.Label(wrap, justify="left", style="Muted.TLabel", text=(
            "① 从 Excel 复制名单（列序：学号、姓名、寝室号、角色，后两列可省），粘到下面点「解析」；\n"
            "② 或点「从文件导入」选 .xlsx / .csv（列序相同）；③ 核对预览无误后「保存到模板」。")).pack(
            anchor="w", pady=(10, 6))

        txt = tk.Text(wrap, height=6, undo=False)
        txt.pack(fill="x")

        holder = ttk.Frame(wrap)
        holder.pack(fill="both", expand=True, pady=(10, 0))
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        tree = ttk.Treeview(holder, columns=("sid", "name", "dorm", "role"),
                            show="headings", height=9, selectmode="extended")
        for cid, title, w in (("sid", "学号", 140), ("name", "姓名", 110),
                              ("dorm", "寝室号", 90), ("role", "角色", 80)):
            tree.heading(cid, text=title)
            tree.column(cid, width=w, anchor="center")
        scr = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scr.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scr.grid(row=0, column=1, sticky="ns")

        status = ttk.Label(wrap, text="等待解析", style="Muted.TLabel")
        status.pack(anchor="w", pady=(6, 0))

        def render():
            for iid in tree.get_children():
                tree.delete(iid)
            for r in rows:
                tree.insert("", "end", values=(r["sid"], r["name"], r["dorm"], r["role"]))

        def delete_selected(event=None):
            for iid in tree.selection():
                idx = tree.index(iid)
                if 0 <= idx < len(rows):
                    rows.pop(idx)
            render()
            status.config(text=f"已删除选中行，剩余 {len(rows)} 人", foreground=MUTED)

        tree.bind("<Delete>", delete_selected)

        def validate(parts_list, source):
            rows.clear()
            parsed, errors = self._parse_roster_parts(parts_list)
            rows.extend(parsed)
            render()
            msg = f"解析到 {len(rows)} 人"
            if errors:
                msg += f"；{len(errors)} 行有问题：{errors[0]}" + (" 等" if len(errors) > 1 else "")
            status.config(text=msg, foreground=(WARN_C if errors else OK_C))
            return rows and not errors

        def parse_text():
            lines = txt.get("1.0", "end").splitlines()
            validate([re.split(r"[\t,，;；\s]+", ln) for ln in lines], "")

        def import_file():
            path = filedialog.askopenfilename(
                parent=top, title="选择名单文件",
                filetypes=[("名单文件", "*.xlsx *.csv"), ("所有文件", "*.*")])
            if not path:
                return
            try:
                parts_list, _hint = self._read_roster_source(path)
            except Exception as exc:
                messagebox.showerror("读取失败", f"{path}\n\n{exc}", parent=top)
                return
            txt.delete("1.0", "end")
            txt.insert("1.0", "\n".join(
                "\t".join(str(c).strip() for c in p if str(c).strip()) for p in parts_list))
            validate(parts_list, "文件")

        def save():
            if not rows:
                status.config(text="先解析出名单再保存", foreground=WARN_C)
                return
            cls_name = en_cls.get().strip()
            if not cls_name:
                messagebox.showwarning("班级名单", "请填写班级全称（用于周表表头与月表）",
                                       parent=top)
                return
            if not messagebox.askyesno(
                    "确认写入",
                    f"将把 {len(rows)} 名学生写入名单模板（写入前会自动备份原文件）。\n"
                    f"当前未保存的加减分录入会被清除，继续吗？", parent=top):
                return
            top.destroy()
            self._apply_roster_to_template(rows, cls_name)

        btns = ttk.Frame(wrap)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="保存到模板", style="Primary.TButton",
                   command=save).pack(side="right")
        ttk.Button(btns, text="删除选中行（Del）", command=delete_selected).pack(
            side="right", padx=(0, 8))
        ttk.Button(btns, text="解析", command=parse_text).pack(side="left")
        ttk.Button(btns, text="从文件导入…", command=import_file).pack(side="left", padx=(8, 0))

        # 对话框本身也支持拖入名单文件（拖到弹窗上 = 直接读入并预览）
        try:
            from tkinterdnd2 import DND_FILES
            top.drop_target_register(DND_FILES)

            def drop_into_dialog(event):
                try:
                    ps = [str(x) for x in self.root.tk.splitlist(event.data)]
                except Exception:
                    return
                path = next((x for x in ps if str(x).lower().endswith((".xlsx", ".csv"))), None)
                if not path or not os.path.exists(path):
                    return
                try:
                    parts, _hint = self._read_roster_source(path)
                except Exception as exc:
                    messagebox.showerror("读取失败", f"{path}\n\n{exc}", parent=top)
                    return
                txt.delete("1.0", "end")
                txt.insert("1.0", "\n".join(
                    "\t".join(str(c).strip() for c in row if str(c).strip()) for row in parts))
                validate(parts, "拖入文件")

            top.bind("<<Drop>>", drop_into_dialog)
        except Exception:
            pass

    @staticmethod
    def _copy_cell_style(src, dst):
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format

    def _apply_roster_to_template(self, rows, cls_raw):
        """把导入的名单写回 名单模板.xlsx 的三张表，然后重载。

        月表模板：第 2 行起 A=学号 B=姓名 C=班级全称，多余旧行清空
        周表模板：A1 写班级名；对半分左右两组，第 4 行起 A/B 与 E/F
        寝室名单：按寝室号升序，每寝占 2 行（A=房号 C=成员空格分隔 G=室长）
        """
        tpl = self._template_path()
        if not os.path.exists(tpl):
            return messagebox.showerror("错误", f"找不到名单模板：\n{tpl}")
        bak = os.path.join(os.path.dirname(tpl), "名单模板.自动备份-"
                           + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx")
        try:
            shutil.copy2(tpl, bak)
        except Exception as exc:
            if not messagebox.askyesno("备份失败",
                                       f"写模板前自动备份失败：{exc}\n仍要继续吗？（有丢数据风险）"):
                return

        wb = None
        while True:
            try:
                wb = openpyxl.load_workbook(tpl)
                break
            except PermissionError:
                if not self.ask_retry_save(tpl):
                    return
        try:
            class_full = cls_raw if "班" in cls_raw else cls_raw + "班"
            self._write_month_roster(wb["月表模板"], rows, cls_raw)
            self._write_week_roster(wb["周表模板"], rows, class_full)
            self._write_dorm_roster(wb["寝室名单"], rows)
            while True:
                try:
                    wb.save(tpl)
                    break
                except PermissionError:
                    if not self.ask_retry_save(tpl):
                        return
        except Exception as exc:
            return messagebox.showerror(
                "写入失败", f"名单写入模板时出错：\n{exc}\n\n原文件未动，备份在：{bak}")
        finally:
            if wb:
                wb.close()

        self.class_name = class_full
        self.reload_template(ask=False)
        self.week_status.config(
            text=f"名单已导入（{len(rows)} 人，原文件备份：{os.path.basename(bak)}）",
            foreground=OK_C)

    def _write_month_roster(self, ws, rows, cls_raw):
        n = len(rows)
        if ws.max_row < 1 + n:            # 超出 47 行骨架时补行（复制第 2 行样式）
            for r in range(ws.max_row + 1, 2 + n):
                for c in range(1, 6):
                    self._copy_cell_style(ws.cell(2, c), ws.cell(r, c))
                d2 = str(ws.cell(2, 4).value or "")
                if "#" in d2:
                    ws.cell(r, 4).value = d2
        for i, row in enumerate(rows):
            r = 2 + i
            ws.cell(r, 1).value = row["sid"]
            ws.cell(r, 2).value = row["name"]
            ws.cell(r, 3).value = cls_raw
        for r in range(2 + n, ws.max_row + 1):
            for c in (1, 2, 3, 5):
                cell = ws.cell(r, c)
                if type(cell).__name__ != "MergedCell":
                    cell.value = None

    def _write_week_roster(self, ws, rows, class_full):
        n = len(rows)
        half = math.ceil(n / 2)
        left, right = rows[:half], rows[half:]
        ws["A1"] = f"{class_full}#月#周记实"
        need_last = 3 + max(len(left), len(right))
        if ws.max_row < need_last:        # 超出 31 行对骨架时补行（复制第 4 行样式）
            for r in range(ws.max_row + 1, need_last + 1):
                for c in range(1, 9):
                    self._copy_cell_style(ws.cell(4, c), ws.cell(r, c))
        for k in range(max(len(left), len(right))):
            r = 4 + k
            l = left[k] if k < len(left) else None
            rr = right[k] if k < len(right) else None
            for col, val in ((1, l and l["sid"]), (2, l and l["name"]),
                             (5, rr and rr["sid"]), (6, rr and rr["name"])):
                cell = ws.cell(r, col)
                if type(cell).__name__ != "MergedCell":
                    cell.value = val or None
            for col in (3, 4, 7, 8):      # 清残留的加减分/原因
                cell = ws.cell(r, col)
                if type(cell).__name__ != "MergedCell":
                    cell.value = None
        for r in range(4 + max(len(left), len(right)), ws.max_row + 1):
            for c in range(1, 9):
                cell = ws.cell(r, c)
                if type(cell).__name__ != "MergedCell":
                    cell.value = None

    def _write_dorm_roster(self, ws, rows):
        dorms = {}
        for row in rows:
            d = (row["dorm"] or "").strip()
            if not d:
                continue
            info = dorms.setdefault(d, {"members": [], "leaders": []})
            info["members"].append(row["name"])
            if row["role"] == "室长":
                info["leaders"].append(row["name"])
        keys = sorted(dorms.keys(),
                      key=lambda s: (0, int(s)) if s.isdigit() else (1, s))
        # 先清空现有寝室区（1~2 行是表头）
        for r in range(3, ws.max_row + 1):
            for c in (1, 3, 7):
                cell = ws.cell(r, c)
                if type(cell).__name__ != "MergedCell":
                    cell.value = None
        # 寝室数超过骨架（12 寝）时成对补行：复制第 3/4 行样式 + 同款合并
        pairs_have = (ws.max_row - 1) // 2
        for p in range(pairs_have, len(keys)):
            r0 = 3 + 2 * p
            for rr, src_r in ((r0, 3), (r0 + 1, 4)):
                for c in range(1, 8):
                    self._copy_cell_style(ws.cell(src_r, c), ws.cell(rr, c))
            for rng in (f"A{r0}:B{r0 + 1}", f"C{r0}:F{r0 + 1}", f"G{r0}:G{r0 + 1}"):
                ws.merge_cells(rng)
        for i, d in enumerate(keys):
            r = 3 + 2 * i
            info = dorms[d]
            ws.cell(r, 1).value = int(d) if d.isdigit() else d
            ws.cell(r, 3).value = " ".join(info["members"])
            ws.cell(r, 7).value = " ".join(info["leaders"]) or None

    # ---------------- 检查更新（GitHub 热更新） ----------------
    def _update_config_path(self):
        return os.path.join(self._base_dir(), UPDATE_CONFIG_NAME)

    def _update_config(self):
        cfg = {"manifest_url": UPDATE_MANIFEST_DEFAULT,
               "mirror_url": UPDATE_MANIFEST_JSDELIVR,
               "last_check": 0.0,
               "skipped_version": ""}
        try:
            with open(self._update_config_path(), "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                for k in cfg:
                    if k in data:
                        cfg[k] = data[k]
        except Exception:
            pass
        return cfg

    def _save_update_config(self, cfg):
        try:
            with open(self._update_config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @staticmethod
    def _http_json(url, timeout=6):
        """拉一个 JSON；任何异常都吞掉返回 None，网络问题绝不打扰界面。"""
        try:
            req = urllib.request.Request(url, headers={"User-Agent": f"deyufen/{APP_VERSION}"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception:
            return None

    def _check_update(self, manual=False):
        """入口：侧边栏按钮（manual=True）或启动 3 秒后自动。同一时间只跑一个。"""
        if getattr(self, "_update_busy", False):
            return
        self._update_busy = True
        if manual and getattr(self, "btn_check_update", None):
            self.btn_check_update.config(state="disabled", text="检查中…")
        threading.Thread(target=self._check_update_worker, args=(manual,), daemon=True).start()

    def _check_update_worker(self, manual):
        # 每次启动都检查（无 24h 节流）；last_check 仅作记录。
        # 自动模式全程静默：无新版 / 网络不通都不弹窗，只有手动检查才给反馈。
        cfg = self._update_config()
        info = None
        seen = []
        for url in (cfg.get("manifest_url"), cfg.get("mirror_url"),
                    UPDATE_MANIFEST_JSDELIVR, UPDATE_MANIFEST_DEFAULT):
            if not url or url in seen:
                continue
            seen.append(url)
            data = self._http_json(url)
            if isinstance(data, dict) and data.get("version"):
                info = data
                break
        if info is None:
            # 最后的兜底：GitHub API 的最新 Release（tag 即版本号）
            rel = self._http_json(UPDATE_GH_API_LATEST)
            if isinstance(rel, dict) and rel.get("tag_name"):
                entry = {"version": str(rel["tag_name"]).lstrip("vV"),
                         "notes": str(rel.get("body") or "")[:600]}
                assets = [a for a in (rel.get("assets") or [])
                          if str(a.get("name", "")).endswith(".zip")]
                if assets:
                    entry["download"] = {"win64": {
                        "url": assets[0].get("browser_download_url"),
                        "size": assets[0].get("size")}}
                info = entry
        self.root.after(0, lambda: self._on_manifest(manual, cfg, info))

    def _update_done(self):
        self._update_busy = False
        if getattr(self, "btn_check_update", None):
            try:
                self.btn_check_update.config(state="normal", text="检查更新")
            except Exception:
                pass

    def _on_manifest(self, manual, cfg, info):
        if info is None:
            if manual:
                messagebox.showinfo("检查更新",
                                    "网络不可用或更新源暂时无法访问。\n"
                                    "可稍后重试，或在「更新设置」里换成镜像地址。")
            self._update_done()
            return
        cfg["last_check"] = time.time()
        self._save_update_config(cfg)
        remote = str(info.get("version") or "")
        if not remote or not newer_version(remote, APP_VERSION):
            if manual:
                messagebox.showinfo("检查更新", f"已是最新版本（v{APP_VERSION}）。")
            self._update_done()
            return
        if not manual and str(cfg.get("skipped_version") or "") == remote:
            self._update_done()
            return
        self._update_done()
        forced = bool(info.get("min_supported")) and \
            newer_version(str(info["min_supported"]), APP_VERSION)
        self._prompt_update(info, forced)

    def _prompt_update(self, info, forced=False):
        remote = str(info.get("version"))
        notes = str(info.get("notes") or "").strip() or "（本次更新没有填写说明）"
        dlg = tk.Toplevel(self.root)
        dlg.title("发现新版本")
        dlg.grab_set()
        dlg.resizable(False, False)
        self._place_dialog(dlg, 500, 400)
        wrap = ttk.Frame(dlg, padding=24)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text=f"新版本 v{remote}（当前 v{APP_VERSION}）",
                  font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        body = ttk.Frame(wrap)
        body.pack(fill="both", expand=True, pady=(12, 4))
        txt = tk.Text(body, width=54, height=10, wrap="word", relief="flat")
        scr = ttk.Scrollbar(body, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scr.set)
        txt.pack(side="left", fill="both", expand=True)
        scr.pack(side="right", fill="y")
        txt.insert("1.0", notes)
        txt.config(state="disabled")

        def go():
            dlg.destroy()
            if self.unsaved:
                return messagebox.showwarning(
                    "先保存再更新",
                    "当前周还有未保存的加减分录入。\n请先点「生成周表」保存，再点「检查更新」。")
            self._download_update(info)

        btns = ttk.Frame(wrap)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="立即更新", style="Primary.TButton", command=go).pack(side="right")
        if forced:
            ttk.Label(wrap, text="当前版本过旧，必须更新后才能继续正常使用",
                      style="Danger.TLabel").pack(anchor="w", pady=(8, 0))
            dlg.protocol("WM_DELETE_WINDOW", lambda: messagebox.showinfo(
                "必须更新", "这个版本必须更新（数据结构不兼容），请点「立即更新」。", parent=dlg))
        else:
            def skip():
                cfg = self._update_config()
                cfg["skipped_version"] = remote
                self._save_update_config(cfg)
                dlg.destroy()
            ttk.Button(btns, text="跳过此版本", command=skip).pack(side="right", padx=(0, 8))
            ttk.Button(btns, text="稍后", command=dlg.destroy).pack(side="right", padx=(0, 8))
            dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    def _download_update(self, info):
        dl = (info.get("download") or {}).get("win64") or {}
        url = dl.get("url")
        if not url:
            return messagebox.showerror(
                "检查更新", "更新源没有提供下载地址。\n请到 GitHub Releases 手动下载安装包。")
        mirror = (info.get("download") or {}).get("mirror")
        version = str(info.get("version"))
        workdir = os.path.join(tempfile.gettempdir(), UPDATE_DIR_NAME)
        zpath = os.path.join(workdir, f"deyufen-{version}.zip")
        try:
            os.makedirs(workdir, exist_ok=True)
        except Exception:
            pass

        win = tk.Toplevel(self.root)
        win.title("正在下载更新")
        win.grab_set()
        win.resizable(False, False)
        self._place_dialog(win, 460, 190)
        wrap = ttk.Frame(win, padding=24)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text=f"正在下载 v{version} 更新包…", style="Strong.TLabel").pack(anchor="w")
        bar = ttk.Progressbar(wrap, maximum=100)
        bar.pack(fill="x", pady=(14, 4))
        lbl = ttk.Label(wrap, text="连接中…", style="Muted.TLabel")
        lbl.pack(anchor="w")
        state = {"cancel": False}

        def cancel():
            state["cancel"] = True
            win.destroy()

        ttk.Button(wrap, text="取消", command=cancel).pack(anchor="e", pady=(12, 0))
        win.protocol("WM_DELETE_WINDOW", cancel)

        def progress(pct, text):
            try:
                bar["value"] = pct
                lbl.config(text=text)
            except Exception:
                pass

        def cb(p, t):
            self.root.after(0, lambda: progress(p, t))

        def worker():
            ok, digest = self._fetch_zip(url, zpath, dl.get("size"),
                                         str(dl.get("sha256") or "").lower(), cb)
            if not ok and mirror and not state["cancel"]:
                ok, digest = self._fetch_zip(mirror, zpath, dl.get("size"),
                                             str(dl.get("sha256") or "").lower(), cb)
            self.root.after(0, lambda: self._download_done(
                win, state, ok, digest, zpath, version))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _fetch_zip(url, path, size_expect, sha_expect, progress_cb):
        """流式下载 + 边下边算 sha256，先写 .part 再转正，不留半截 zip。

        返回 (ok, sha256hex)；失败时 (False, 错误说明)。
        """
        tmp = path + ".part"
        h = hashlib.sha256()
        got = 0
        try:
            req = urllib.request.Request(url, headers={"User-Agent": f"deyufen/{APP_VERSION}"})
            with urllib.request.urlopen(req, timeout=20) as resp, open(tmp, "wb") as f:
                total = int(resp.headers.get("Content-Length") or 0)
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)
                    h.update(chunk)
                    got += len(chunk)
                    pct = (got * 100 / total) if total else 0
                    tail = f" / {total / 1048576:.1f} MB" if total else ""
                    progress_cb(pct, f"已下载 {got / 1048576:.1f} MB{tail}")
            if size_expect and got != int(size_expect):
                raise ValueError("文件大小与更新清单不符")
            if sha_expect and h.hexdigest() != sha_expect:
                raise ValueError("sha256 校验失败，文件可能不完整")
            os.replace(tmp, path)
            return True, h.hexdigest()
        except Exception as exc:
            try:
                os.remove(tmp)
            except OSError:
                pass
            return False, str(exc) or type(exc).__name__

    @staticmethod
    def _sha256_file(path):
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()

    def _download_done(self, win, state, ok, digest, zpath, version):
        try:
            win.destroy()
        except Exception:
            pass
        if state["cancel"]:
            return
        if not ok:
            return messagebox.showerror(
                "更新下载失败", f"{digest}\n\n已保留当前版本，可稍后重试。\n"
                f"国内网络打不开 GitHub 时，可到 Releases 页面手动下载。")
        if not messagebox.askyesno(
                "下载完成",
                f"v{version} 更新包已就绪并通过校验。\n\n"
                f"点「是」立即安装：程序会退出，由更新器替换文件后自动重启。"):
            return
        self._start_updater(zpath, digest if ok else "", version)

    def _start_updater(self, zip_path, sha256_value, version):
        """写 job 文件并启动 update.exe，然后正常退出（由更新器负责重启）。"""
        update_exe = os.path.join(self._base_dir(), "update.exe")
        if not os.path.exists(update_exe):
            return messagebox.showerror(
                "缺少更新器",
                f"找不到更新器：\n{update_exe}\n\n"
                f"请到 GitHub Releases 下载完整安装包手动覆盖安装。")
        job = {"zip": zip_path,
               "sha256": sha256_value or "",
               "dir": self._base_dir(),
               "pid": os.getpid(),
               "version": version,
               "restart": os.path.join(self._base_dir(), "德育分自动化系统.exe")}
        try:
            job_path = os.path.join(os.path.dirname(zip_path), "update.json")
            with open(job_path, "w", encoding="utf-8") as f:
                json.dump(job, f, ensure_ascii=False, indent=2)
            subprocess.Popen([update_exe, "--job", job_path])
        except Exception as exc:
            return messagebox.showerror("更新", f"启动更新器失败：{exc}")
        self._write_bridge_state(running=False)
        self.root.destroy()

    def show_settings_dialog(self):
        """左下角「设置」：左侧模块导航，右侧内容卡片。"""
        # ── 外观常量（主题色）──
        from tkwebui.core import PALETTE
        PAGE_BG = PALETTE["bg"]          # 白
        EDGE = PALETTE["border"]         # 卡片描边
        HOVER = "#EDF6F3"                # 悬停浅青
        SEL_BG = "#DCEEEA"               # 选中底色
        SEL_FG = "#0B7A70"               # 选中文字（主题青绿）

        top = tk.Toplevel(self.root)
        top.title("设置")
        top.grab_set()
        top.resizable(False, False)
        body = ttk.Frame(top, padding=20)
        body.pack(fill="both", expand=True)

        # ── 左：模块导航（直角条目 + 选中指示条）──
        nav_edge = tk.Frame(body, bg=EDGE, padx=1, pady=1)
        nav_edge.pack(side="left", fill="y")
        nav = tk.Frame(nav_edge, bg=PAGE_BG)
        nav.pack(fill="both", expand=True)
        nav_items = {}

        # ── 右：内容卡片（白底描边）──
        card_edge = tk.Frame(body, bg=EDGE, padx=1, pady=1)
        card_edge.pack(side="left", fill="both", expand=True, padx=(16, 0))
        card = tk.Frame(card_edge, bg=PAGE_BG)
        card.pack(fill="both", expand=True)

        panels = {}
        current = {"key": None}

        def paint(key, state):
            it = nav_items[key]
            if state == "sel":
                bg, fg, bold = SEL_BG, SEL_FG, True
            elif state == "hover":
                bg, fg, bold = HOVER, INK, False
            else:
                bg, fg, bold = PAGE_BG, INK, False
            font = ("Microsoft YaHei UI", 10, "bold") if bold else ("Microsoft YaHei UI", 10)
            it["row"].config(bg=bg)
            it["lab"].config(bg=bg, fg=fg, font=font)
            it["ind"].config(bg=SEL_FG if state == "sel" else bg)

        def show(key):
            current["key"] = key
            for k, p in panels.items():
                p.pack_forget()
                paint(k, "sel" if k == key else "base")
            panels[key].pack(fill="both", expand=True, padx=16, pady=14)
            # 窗口高度跟随当前面板（不拉长、不裁切），并夹在屏幕内居中
            top.update_idletasks()
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            w = min(top.winfo_reqwidth() + 12, sw - 40)
            h = min(top.winfo_reqheight() + 12, sh - 80)
            self._place_dialog(top, w, h)

        def add_nav(key, label):
            row = tk.Frame(nav, bg=PAGE_BG, cursor="hand2")
            row.pack(fill="x")
            ind = tk.Frame(row, bg=PAGE_BG, width=3)
            ind.pack(side="left", fill="y")
            ind.pack_propagate(False)
            lab = tk.Label(row, text="  " + label + "   ", bg=PAGE_BG, fg=INK,
                           font=("Microsoft YaHei UI", 10), pady=10, anchor="w")
            lab.pack(side="left", fill="x", expand=True)

            def on_click(e=None):
                show(key)

            def on_enter(e):
                if current["key"] != key:
                    paint(key, "hover")

            def on_leave(e):
                if current["key"] != key:
                    paint(key, "base")

            for w in (row, lab):
                w.bind("<Button-1>", on_click)
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
            nav_items[key] = {"row": row, "lab": lab, "ind": ind}

        # ── 面板：名单模板 ──
        p1 = tk.Frame(card, bg=PAGE_BG)
        panels["roster"] = p1
        ttk.Label(p1, text="名单模板", style="CardTitle.TLabel").pack(anchor="w")
        # 高 DPI/小屏下四个按钮一行会溢出：固定两行两列，任何缩放都放得下
        row1 = ttk.Frame(p1)
        row1.pack(fill="x", pady=(14, 0))
        ttk.Button(row1, text="添加模板…", style="Primary.TButton",
                   command=lambda: self.pick_roster_file(parent=top)).pack(side="left")
        ttk.Button(row1, text="打开模板编辑",
                   command=self.open_template_file).pack(side="left", padx=(8, 0))
        row2 = ttk.Frame(p1)
        row2.pack(fill="x", pady=(8, 0))
        ttk.Button(row2, text="刷新模板",
                   command=self.reload_template).pack(side="left")
        ttk.Button(row2, text="打开文件夹",
                   command=lambda: self._open_folder(self._base_dir())).pack(
            side="left", padx=(8, 0))
        ttk.Label(p1, wraplength=470, justify="left", style="Muted.TLabel", text=(
            "「添加模板」选 名单模板.xlsx（自动带出班级/寝室/室长）或四列名单文件，\\n"
            "选好即替换当前模板（原文件自动备份）；也可以把文件直接拖进主窗口。\\n\\n"
            "「打开模板编辑」用 Excel 打开 名单模板.xlsx，改完保存并关闭 Excel，\\n"
            "再点上面的「刷新模板」生效。")).pack(anchor="w", pady=(14, 0))

        # ── 面板：版本与更新 ──
        p2 = tk.Frame(card, bg=PAGE_BG)
        panels["update"] = p2
        cfg = self._update_config()
        ttk.Label(p2, text=f"当前版本 v{APP_VERSION}",
                  style="CardTitle.TLabel").pack(anchor="w")
        row_u = ttk.Frame(p2)
        row_u.pack(fill="x", pady=(14, 0))
        self.btn_check_update = ttk.Button(row_u, text="检查更新",
                                           command=lambda: self._check_update(manual=True))
        self.btn_check_update.pack(side="left")
        # 提示放按钮正下方整行（并排在高 DPI 下会把按钮挤出屏幕）
        ttk.Label(p2, wraplength=470, style="Muted.TLabel", text=(
            "每次启动也会自动检查，无新版或网络不通时不弹窗")).pack(
            anchor="w", pady=(6, 0))
        ttk.Label(p2, text="更新清单地址（主）", style="Muted.TLabel").pack(
            anchor="w", pady=(18, 2))
        e1 = ttk.Entry(p2)
        e1.pack(fill="x")
        e1.insert(0, cfg.get("manifest_url") or "")
        ttk.Label(p2, text="更新清单地址（镜像，主地址失败时使用）",
                  style="Muted.TLabel").pack(anchor="w", pady=(10, 2))
        e2 = ttk.Entry(p2)
        e2.pack(fill="x")
        e2.insert(0, cfg.get("mirror_url") or "")
        ttk.Label(p2, wraplength=470, style="Muted.TLabel", text=(
            "国内网络打不开 GitHub 时，可把镜像换成加速地址，例如\\n"
            "https://ghfast.top/https://raw.githubusercontent.com/…/version.json")).pack(
            anchor="w", pady=(10, 0))

        def save_cfg():
            cfg["manifest_url"] = e1.get().strip() or UPDATE_MANIFEST_DEFAULT
            cfg["mirror_url"] = e2.get().strip()
            self._save_update_config(cfg)
            messagebox.showinfo("设置", "更新源已保存。", parent=top)

        def reset_cfg():
            e1.delete(0, tk.END)
            e1.insert(0, UPDATE_MANIFEST_DEFAULT)
            e2.delete(0, tk.END)
            e2.insert(0, UPDATE_MANIFEST_JSDELIVR)

        row_s = ttk.Frame(p2)
        row_s.pack(fill="x", pady=(10, 0))
        ttk.Button(row_s, text="保存", style="Primary.TButton",
                   command=save_cfg).pack(side="right")
        ttk.Button(row_s, text="恢复默认", command=reset_cfg).pack(
            side="right", padx=(0, 8))

        # ── 面板：AI 对接 ──
        p3 = tk.Frame(card, bg=PAGE_BG)
        panels["ai"] = p3
        ttk.Label(p3, text="AI 对接（MCP）", style="CardTitle.TLabel").pack(anchor="w")
        # 宽按钮单独一行，窄的两个一行（防高 DPI 溢出）
        row_ai = ttk.Frame(p3)
        row_ai.pack(fill="x", pady=(14, 0))
        ttk.Button(row_ai, text="重新生成 MCP 配置",
                   command=self._regen_mcp_config).pack(side="left")
        row_ai2 = ttk.Frame(p3)
        row_ai2.pack(fill="x", pady=(8, 0))
        ttk.Button(row_ai2, text="接入说明",
                   command=self._open_ai_doc).pack(side="left")
        ttk.Button(row_ai2, text="打开文件夹",
                   command=self._open_ai_folder).pack(side="left", padx=(8, 0))
        ttk.Label(p3, wraplength=470, justify="left", style="Muted.TLabel", text=(
            "在 AI 客户端（Cherry Studio / Claude Desktop 等）里添加\\n"
            "AI对接\\\\MCP配置.json 的内容即可。挪动安装目录或重装后，\\n"
            "点「重新生成 MCP 配置」按实际路径重写。")).pack(anchor="w", pady=(14, 0))

        # ── 面板：联系作者 ──
        p4 = tk.Frame(card, bg=PAGE_BG)
        panels["contact"] = p4
        ttk.Label(p4, text="联系作者", style="CardTitle.TLabel").pack(anchor="w")
        row_mail = ttk.Frame(p4)
        row_mail.pack(fill="x", pady=(14, 0))
        ttk.Label(row_mail, text="xiaohai@fale52.top",
                  style="CardTitle.TLabel").pack(side="left")

        def copy_mail():
            self.root.clipboard_clear()
            self.root.clipboard_append("xiaohai@fale52.top")
            try:
                notify(self.root, "邮箱已复制：xiaohai@fale52.top",
                       title="联系作者", kind="success", duration=5000)
            except Exception:
                pass

        ttk.Button(row_mail, text="复制邮箱", command=copy_mail).pack(
            side="left", padx=(12, 0))
        ttk.Label(p4, wraplength=470, justify="left", style="Muted.TLabel", text=(
            "反馈问题时请附上：\n"
            "  1. 软件版本（设置左上角可见）\n"
            "  2. 问题截图（周表/月表页的「截图」按钮可直接生成）\n"
            "  3. 当月文件夹里的报错或现象说明")).pack(anchor="w", pady=(14, 0))

        add_nav("roster", "名单模板")
        add_nav("update", "版本与更新")
        add_nav("ai", "AI 对接")
        add_nav("contact", "联系作者")

        show("roster")

    def get_folder_path(self):
        """月份数据文件夹，锚定在程序根目录而不是当前工作目录。"""
        return os.path.join(self._base_dir(), f"{self.current_year.get()}年{self.current_month.get()}月")

    def get_week_file_name(self):
        return f"{self.class_name}{self.current_month.get()}月德育考评周纪实表.xlsx".replace("班班", "班")

    def get_month_file_name(self):
        return f"{self.class_name}{self.current_month.get()}月德育考评月表.xlsx".replace("班班", "班")

    # ---------------- 撤回 / 未保存保护 ----------------
    UNDO_LIMIT = 50

    def _push_undo(self):
        """记录一次可撤回状态。栈有上限，避免长时间使用内存单调增长。"""
        self.redo_stack.clear()
        self.undo_stack.append(deepcopy(self.week_data))
        if len(self.undo_stack) > self.UNDO_LIMIT:
            del self.undo_stack[:len(self.undo_stack) - self.UNDO_LIMIT]
        self.btn_undo.config(state="normal")
        self.btn_redo.config(state="disabled")
        self.unsaved = True
        # 分数变了，之前扫描出来的月度累加已经过期，必须重扫
        self.week_scanned = False
        self._mark_month_stale()

    def _mark_month_stale(self):
        """数据变了：把月度累加标记为过期，并同步界面状态。"""
        if not hasattr(self, "lbl_scan_status"):
            return
        self.lbl_scan_status.config(text="数据已变化，需重新读取", foreground=WARN_C)
        if self.current_view == "month":
            self.check_month_ready()

    def _cycle_key(self):
        return (self.current_year.get(), self.current_month.get(), self.current_week.get())

    def _week_file(self):
        return os.path.join(self.get_folder_path(), self.get_week_file_name())

    def _remember_week_file_state(self):
        """记下「载入时这个周表文件在不在」，用来发现被外部删掉。"""
        path = self._week_file()
        self.week_file_state = (path, os.path.exists(path))

    def _on_window_focus(self, event=None):
        self._check_disk_drift()

    def _check_disk_drift(self):
        """周表文件被软件外部删掉了：提醒一句，别让内存里的记录静默写回去。"""
        state = getattr(self, "week_file_state", None)
        if not state:
            return
        path, existed = state
        if not existed or path != self._week_file():
            return
        if os.path.exists(path):
            return
        self.week_file_state = (path, False)      # 只提醒一次
        try:
            count = sum(len(info.get('records') or []) for info in self.week_data.values())
        except Exception:
            count = 0
        name = os.path.basename(path)
        if not count:
            self.week_status.config(text=f"{name} 已在软件外被删除", foreground=WARN_C)
            return
        self.week_status.config(
            text=f"{name} 已被删除，内存里仍有 {count} 条记录", foreground=WARN_C)
        notify(self.root,
               f"{name} 被删掉了，软件内存里还留着 {count} 条记录。\n"
               f"直接点「生成周表」会把它重新写回来；不想留就先点「撤回」或清空。",
               title="周表文件不见了", kind="warn", duration=25000)

    def _ask_unsaved(self, what, allow_carry=True):
        """有未保存录入时问一句。返回 save / carry / discard / cancel。

        carry = 把这一周还没生成的录入（含寝室检查次数）整体搬到要切过去的那一周。
        allow_carry=False 时不给这个选项（比如关窗退出）。
        """
        if not self.unsaved:
            return "discard"

        dlg = tk.Toplevel(self.root)
        dlg.title("有未保存的录入")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        self._place_dialog(dlg, 470, 260)
        result = {"value": "cancel"}

        outer = ttk.Frame(dlg, padding=20)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="当前周的加减分还没生成表格。").pack(anchor="w")
        ttk.Label(outer, text=f"要{what}，这些改动怎么处理？").pack(anchor="w", pady=(4, 12))

        try:
            n_rec = sum(len(info.get('records') or []) for info in self.week_data.values())
        except Exception:
            n_rec = 0
        try:
            n_dorm = sum(len(v) for v in getattr(self, "dorm_inspections", {}).values())
        except Exception:
            n_dorm = 0
        ttk.Label(outer,
                  text=f"没生成的：{n_rec} 条加减分记录、{n_dorm} 次寝室检查").pack(anchor="w")

        btns = ttk.Frame(outer)
        btns.pack(side="bottom", fill="x", pady=(16, 0))

        def pick(value):
            result["value"] = value
            dlg.destroy()

        ttk.Button(btns, text="先生成表格", command=lambda: pick("save")).pack(side="left")
        if allow_carry:
            ttk.Button(btns, text="整个迁移过去",
                       command=lambda: pick("carry")).pack(side="left", padx=6)
        ttk.Button(btns, text="放弃改动", command=lambda: pick("discard")).pack(side="left")
        ttk.Button(btns, text="取消", command=lambda: pick("cancel")).pack(side="right")
        dlg.protocol("WM_DELETE_WINDOW", lambda: pick("cancel"))
        dlg.bind("<Escape>", lambda e: pick("cancel"))
        dlg.grab_set()
        dlg.wait_window()
        return result["value"]

    def _restore_cycle(self):
        y, m, w = self._loaded_cycle
        self._suspend_cycle_cb = True
        try:
            self.current_year.set(y)
            self.current_month.set(m)
            self.current_week.set(w)
        finally:
            self._suspend_cycle_cb = False

    def _switch_cycle(self, what):
        """切周次/切考评周期的统一入口。"""
        action = self._ask_unsaved(what)
        if action == "cancel":
            self._restore_cycle()
            return
        snap = None
        if action == "carry":
            snap = self._snapshot_unsaved()
        elif action == "save":
            self.save_week_data()
            if self.unsaved:            # 保存被取消或失败
                self._restore_cycle()
                return
        self.load_week_data()
        if snap is not None:
            self._carry_unsaved(snap)

    def on_cycle_change(self):
        if getattr(self, "_suspend_cycle_cb", False):
            return
        self._switch_cycle("切换考评周期")

    def on_week_change(self):
        if getattr(self, "_suspend_cycle_cb", False):
            return
        self._switch_cycle("切换周次")

    def _snapshot_unsaved(self):
        """把没落盘的录入和寝室检查次数拷一份，供「整个迁移」用。"""
        return {
            "week": deepcopy(self.week_data),
            "dorm": deepcopy(getattr(self, "dorm_inspections", {})),
            "dorm_ai": set(getattr(self, "dorm_ai_set", set())),
        }

    def _carry_unsaved(self, snap):
        """把上一份快照整体搬到当前周期（记录直接追加，不重算分数）。"""
        moved = 0
        for sid, info in snap["week"].items():
            if sid not in self.week_data:
                continue
            recs = list(info.get('records') or [])
            if not recs:
                continue
            self.week_data[sid]['records'].extend(deepcopy(recs))
            moved += len(recs)

        dorm_moved = 0
        for dorm, values in (snap["dorm"] or {}).items():
            if not values:
                continue
            self.dorm_inspections.setdefault(dorm, []).extend(deepcopy(values))
            dorm_moved += len(values)
        if hasattr(self, "dorm_ai_set"):
            self.dorm_ai_set |= snap["dorm_ai"]

        if not moved and not dorm_moved:
            self.week_status.config(text="这一周本来就没有可迁移的改动", foreground=WARN_C)
            return

        self.unsaved = True
        self.week_scanned = False
        self._mark_month_stale()
        if not self.showing_inquiry:
            self.refresh_week_tree()
        self._restore_dorm_sliders_from_inspections()

        parts = []
        if moved:
            parts.append(f"{moved} 条记录")
        if dorm_moved:
            parts.append(f"{dorm_moved} 次寝室检查")
        target = (f"{self.current_year.get()}年{self.current_month.get()}月"
                  f"第{self.current_week.get()}周")
        self.week_status.config(
            text=f"已把 {'和 '.join(parts)} 迁移到 {target}，记得生成周表",
            foreground=WARN_C)

    def _restore_dorm_sliders_from_inspections(self):
        """迁移后把滑块拨回检查结果，别出现「看着四星、实际按五星算」。"""
        controls = getattr(self, "dorm_controls", {}) or {}
        for dorm, values in getattr(self, "dorm_inspections", {}).items():
            if not values:
                continue
            star = values[-1]
            if dorm in self.dorm_stars:
                self.dorm_stars[dorm] = star
            seg = controls.get(dorm)
            if seg is None:
                continue
            try:
                if int(seg.variable.get()) != int(star):
                    seg.variable.set(star)
            except Exception:
                pass
        for name in ("_rebuild_dorm_ai_rows", "_paint_dorm_states"):
            fn = getattr(self, name, None)
            if fn is None:
                continue
            try:
                fn()
            except Exception:
                pass

    # ---------------- 第一页：周表逻辑 ----------------
    def load_week_data(self):
        # 正在看「加分标准」时切周期：先退出查看态，否则表格还停在标准表上
        if self.showing_inquiry:
            self.showing_inquiry = False
            if hasattr(self, "btn_inquiry"):
                self.btn_inquiry.config(text="查加分标准", style="TButton")
                self.pv_title.config(text="本周实时预览")
                self.setup_week_tree_columns()

        old_cycle = self._loaded_cycle
        new_cycle = self._cycle_key()

        for sid in self.week_data:
            self.week_data[sid]['records'] = []

        # 换了周期，上一周期的撤回快照不再适用（撤回会把别的周的数据搬过来）
        had_undo = bool(self.undo_stack)
        self.undo_stack.clear()
        self.redo_stack.clear()
        if hasattr(self, "btn_undo"):
            self.btn_undo.config(state="disabled")
            self.btn_redo.config(state="disabled")

        # 换了年月，月度汇总相关状态全部失效
        if (old_cycle[0], old_cycle[1]) != (new_cycle[0], new_cycle[1]):
            self.week_scanned = False
            self.last_month_loaded = False
            for sid in self.last_month_scores:
                self.last_month_scores[sid] = 0.0
            for sid in self.month_sum:
                self.month_sum[sid] = 0.0

        self.unsaved = False
        self._loaded_cycle = new_cycle

        folder = self.get_folder_path()
        file_path = os.path.join(folder, self.get_week_file_name())
        w = self.current_week.get()
        s_name = f"第{w}周"

        if os.path.exists(file_path):
            wb = None
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                if s_name in wb.sheetnames:
                    ws = wb[s_name]

                    # 签名栏边界按栏独立：左栏看 A/B（班主任签名在 A），
                    # 右栏看 E/F（考评小组签名合并格在 E）。全局截断会漏掉
                    # 与签名区同行的最后一行学生（王小领 bug）。
                    def _sig_row(cols):
                        for r in range(4, ws.max_row + 1):
                            for c in cols:
                                v = str(ws.cell(row=r, column=c).value or "")
                                if ("签名" in v) or ("考评" in v):
                                    return r
                        return ws.max_row + 1
                    sig_row_left = _sig_row((1, 2))
                    sig_row_right = _sig_row((5, 6))

                    active_l_sid = None
                    active_r_sid = None
                    for r in range(4, ws.max_row + 1):
                        if r < sig_row_left:
                            l_id_cell = ws.cell(row=r, column=1).value
                            if l_id_cell is not None and str(l_id_cell).strip() != "":
                                active_l_sid = self.clean_val(l_id_cell)
                            if active_l_sid and active_l_sid in self.week_data:
                                sc = ws.cell(row=r, column=3).value
                                rn = ws.cell(row=r, column=4).value
                                num = self._to_number(sc)
                                if num is not None:
                                    self.week_data[active_l_sid]['records'].append(
                                        (num, self.scrub_ghost_data(rn)))

                        if ws.max_column >= 7 and r < sig_row_right:
                            r_id_cell = ws.cell(row=r, column=5).value
                            if r_id_cell is not None and str(r_id_cell).strip() != "":
                                active_r_sid = self.clean_val(r_id_cell)
                            if active_r_sid and active_r_sid in self.week_data:
                                sc = ws.cell(row=r, column=7).value
                                rn = ws.cell(row=r, column=8).value
                                num = self._to_number(sc)
                                if num is not None:
                                    self.week_data[active_r_sid]['records'].append(
                                        (num, self.scrub_ghost_data(rn)))

                    self.week_status.config(text=f"已读取 {s_name} 的历史记录", foreground=OK_C)
                    self._week_file_existed = True
                else:
                    self.week_status.config(text=f"{s_name} 暂无记录，将新建", foreground=WARN_C)
            except Exception as e:
                self.week_status.config(text="读取异常", foreground=ERR_C)
            finally:
                if wb: wb.close()
        else:
            self._week_file_existed = False
            self.week_status.config(text="周表尚未创建", foreground=WARN_C)

        # 上一周期收到的寝室检查次数不能跟着走，否则上周的五星会被算进这周
        if getattr(self, "dorm_inspections", None):
            self.dorm_inspections = {}

        # 星级选择跟着周次走，避免把上一周选的星级打到这一周
        self._sync_dorm_stars()
        if had_undo:
            self.week_status.config(text="已切换周期（原撤回栈已清空）", foreground=WARN_C)
        if not self.showing_inquiry:
            self.refresh_week_tree()
        if self.current_view == "month":
            self.check_month_ready()
        if self.current_view == "ai":
            self.refresh_ai_view()
        # 周期变了要马上告诉 AI，别等心跳
        self._write_bridge_state()

        # 记下这一步磁盘上有没有这个周表；窗口回到前台时检查有没有被外部删掉
        self._remember_week_file_state()
        if not getattr(self, "_drift_bound", False):
            try:
                self.root.bind("<FocusIn>", self._on_window_focus, add="+")
                self._drift_bound = True
            except Exception:
                pass

    def add_manual_score(self):
        kw = self.clean_name(self.en_name.get())
        rs = self.en_reason.get().strip()
        sc_str = self.en_score.get().strip()

        if not kw:
            self.en_name.focus()
            return messagebox.showwarning("还差一项", "请填写姓名，或学号的后两位。")
        if not sc_str:
            self.en_score.focus()
            return messagebox.showwarning("还差一项", "请填写分值，例如 1、-0.5 或 10+2。")

        # 统一走安全求值：'nan' / 'inf' / '1e400' 这类会被挡下来，
        # 否则会写进 week_data 并在刷新预览表时把程序搞崩
        try:
            sc = self._eval_score(sc_str)
        except (ValueError, SyntaxError) as exc:
            return messagebox.showerror("分值无效",
                                        f"{exc}\n\n支持数字或表达式，例如 1、-0.5、10+2。")
        if not math.isfinite(sc):
            return messagebox.showerror("分值无效", "分值必须是有限的数字。")

        # 匹配学生：优先精确匹配，再模糊匹配
        exact_matches = []
        fuzzy_matches = []
        for sid, info in self.week_data.items():
            clean_n = self.clean_name(info['name'])
            if kw == clean_n or kw == sid:
                exact_matches.append((sid, info))
            elif kw in clean_n:
                fuzzy_matches.append((sid, info))
            elif kw.isdigit():
                # 数字匹配：2→尾号02，22→尾号22
                if sid.endswith(kw.zfill(2)):
                    fuzzy_matches.append((sid, info))

        if exact_matches:
            # 精确匹配（全名或完整学号）→ 直接使用，不弹窗
            t_sid = exact_matches[0][0]
        elif len(fuzzy_matches) == 1:
            # 模糊匹配只有一个人 → 直接使用
            t_sid = fuzzy_matches[0][0]
        elif len(fuzzy_matches) >= 2:
            # 模糊匹配多人 → 弹窗选择
            t_sid = self._pick_student(kw, fuzzy_matches)
            if t_sid is None:
                return
        else:
            messagebox.showerror("未找到", "找不到该学生！")
            return

        if t_sid:
            self.last_edited_sid = t_sid
            self._push_undo()
            reason_text = rs if rs else "手动加分"
            self.week_data[t_sid]['records'].append((sc, reason_text))
            if not self.showing_inquiry:
                self.refresh_week_tree()
            self.en_name.delete(0, tk.END)
            self.en_reason.delete(0, tk.END)
            self.en_score.delete(0, tk.END)
            self.en_name.focus()

    def apply_dorm_scores(self):
        period = self.dorm_period_var.get()

        # 按月规则给的是整月的加减分，只能在某一个月里应用一次，
        # 否则月度汇总会把同一条记录累加多遍
        if period == "月":
            if not messagebox.askyesno(
                    "按月打分只应用一次",
                    "按月规则给的是整个月的寝室加减分。\n\n"
                    "请在每个月只应用一次，并选一个周次保存即可。\n"
                    "如果在同一个月的不同周次各应用一次，\n"
                    "月度汇总里这部分会被重复计分。\n\n确定继续吗？"):
                return

        if not self.dorm_stars:
            return messagebox.showwarning("还没选星级", "请先给寝室选择星级。")

        self._push_undo()
        prefix = "周" if period == "周" else "月度"
        rules = self._load_dorm_rules()

        star_zh_map = {3: "三", 4: "四", 5: "五"}

        # 清除同类型的老旧星级记录，防叠加
        for info in self.week_data.values():
            info['records'] = [r for r in info['records'] if '星寝室' not in str(r[1])]

        inspections = getattr(self, "dorm_inspections", {})
        applied = 0
        for sid, info in self.week_data.items():
            dorm = info.get('dorm', '未知')
            role = info.get('role', '成员')
            if dorm not in self.dorm_stars:
                continue
            stars_list = inspections.get(dorm) or [self.dorm_stars[dorm]]
            # 按次数计分：本周期每收到一次检查结果，各按当次星级取分求和
            score = 0.0
            counts = {}
            for star in stars_list:
                if star not in rules[period]:
                    continue
                score += rules[period][star].get(role, 0.0)
                counts[star] = counts.get(star, 0) + 1
            if not counts:
                continue
            if abs(score) < 1e-9 and len(stars_list) < 2:
                continue          # 单次四星这类 0 分不写；多次净 0 仍写，便于追溯
            reason = prefix + "+".join(
                f"{star_zh_map[star]}星寝室×{n}" if n > 1 else f"{star_zh_map[star]}星寝室"
                for star, n in sorted(counts.items(), reverse=True))
            if role == "室长":
                reason += "(寝室长)"
            info["records"].append((round(score, 2), reason))
            applied += 1

        if not self.showing_inquiry:
            self.refresh_week_tree()
        self.generate_dorm_excel(period)
        self.save_week_data()

        # 按次数计的分数已写入周表：清零计数与青标，寝室号转黑（已应用）防重复计分
        self.dorm_applied = set(inspections.keys())
        self.dorm_inspections = {}
        self.dorm_ai_set = set()
        self._paint_dorm_states()

        self.week_status.config(
            text=f"寝室星级已打分并保存（{period}规则，{applied} 条记录）", foreground=OK_C)

    def _dorm_star_file(self):
        """当前周期对应的星级评定表路径（周/月两种命名）。"""
        folder = self.get_folder_path()
        y, m, w = self.current_year.get(), self.current_month.get(), self.current_week.get()
        if self.dorm_period_var.get() == "月":
            return os.path.join(folder, f"{y}年{m}月_月度寝室星级评定表.xlsx")
        return os.path.join(folder, f"{y}年{m}月第{w}周_寝室星级评定表.xlsx")

    def _sync_dorm_stars(self):
        """把星级面板重置为该周期已保存的值；没有记录则全部回到 4 星。

        返回恢复成功的寝室数。界面还没建好时直接返回 0（__init__ 阶段会调到）。
        """
        if not getattr(self, "dorm_controls", None):
            return 0
        self.dorm_ai_set = getattr(self, "dorm_ai_set", set())
        self.dorm_ai_set.clear()      # 换周期后 AI 代设的星级作废
        for dorm, seg in self.dorm_controls.items():
            self.dorm_stars[dorm] = 4
            if int(seg.variable.get()) != 4:
                seg.variable.set(4)

        path = self._dorm_star_file()
        if not os.path.exists(path):
            return 0

        loaded = 0
        wb = None
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]
            for r in range(2, ws.max_row + 1):
                dorm_val = ws.cell(row=r, column=1).value
                star_val = ws.cell(row=r, column=2).value
                if dorm_val is None or star_val is None:
                    continue
                dorm = str(dorm_val).strip()
                star = self.STAR_TEXT_MAP.get(str(star_val).strip())
                if not dorm or star is None or dorm not in self.dorm_controls:
                    continue
                self.dorm_stars[dorm] = star
                self.dorm_controls[dorm].variable.set(star)
                loaded += 1
        except Exception:
            pass
        finally:
            if wb:
                wb.close()
        return loaded

    def load_dorm_stars(self):
        """「读取周表」按钮：从已保存的星级评定表恢复打分状态。"""
        path = self._dorm_star_file()
        if not os.path.exists(path):
            messagebox.showinfo("还没有星级记录",
                                f"这个周期还没保存过星级评定：\n{path}\n\n"
                                f"先选好星级，点「应用打分」就会生成这张表。")
            return
        loaded = self._sync_dorm_stars()
        self.frame_dorm.update_idletasks()
        messagebox.showinfo("完成", f"已从星级评定表恢复 {loaded} 个寝室的打分状态")

    def _base_dir(self):
        """程序的工作根目录：模板、数据文件夹、生成结果都以它为基准。

        打包后 = exe 所在目录；源码运行时 = 脚本所在目录，
        若脚本放在子目录（如 软件说明/）而模板在其上一级，则自动上溯一级。
        这样无论从哪个工作目录启动，都不会找不到 名单模板.xlsx。
        """
        cached = getattr(self, "_base_dir_cache", None)
        if cached:
            return cached

        if getattr(sys, 'frozen', False):
            here = os.path.dirname(os.path.abspath(sys.executable))
        else:
            here = os.path.dirname(os.path.abspath(__file__))

        base = here
        for cand in (here, os.path.dirname(here)):
            if cand and os.path.exists(os.path.join(cand, "名单模板.xlsx")):
                base = cand
                break
        self._base_dir_cache = base
        return base

    def generate_dorm_excel(self, period):
        """生成独立的五星打分存档，标题自动适应周或月"""
        folder = self.get_folder_path()
        if not os.path.exists(folder): os.makedirs(folder, exist_ok=True)

        y, m, w = self.current_year.get(), self.current_month.get(), self.current_week.get()

        if period == "周":
            file_path = os.path.join(folder, f"{y}年{m}月第{w}周_寝室星级评定表.xlsx")
            title = f"第{w}周评定"
        else:
            file_path = os.path.join(folder, f"{y}年{m}月_月度寝室星级评定表.xlsx")
            title = f"{m}月评定"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = title

        headers = ["寝室号", "评定星级", "室长分数变动", "成员分数变动"]
        ws.append(headers)

        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        align = Alignment(horizontal="center", vertical="center")

        for col, width in enumerate([15, 15, 18, 18], 1):
            ws.column_dimensions[get_column_letter(col)].width = width
            cell = ws.cell(row=1, column=col)
            cell.font = Font(name="微软雅黑", size=11, bold=True)
            cell.border = border
            cell.alignment = align

        row_idx = 2

        rules = self._load_dorm_rules()
        star_zh_map = {3: "三", 4: "四", 5: "五"}

        for dorm in sorted(self.dorm_stars.keys()):
            star = self.dorm_stars[dorm]
            if star not in rules[period]:
                continue
            score_m = signed_text(rules[period][star]['成员'])
            score_l = signed_text(rules[period][star]['室长'])
            star_zh = star_zh_map[star]

            ws.cell(row=row_idx, column=1, value=dorm)
            ws.cell(row=row_idx, column=2, value=f"{star_zh}星")
            ws.cell(row=row_idx, column=3, value=score_l)
            ws.cell(row=row_idx, column=4, value=score_m)

            for c in range(1, 5):
                cell = ws.cell(row=row_idx, column=c)
                cell.border = border
                cell.alignment = align
                cell.font = Font(name="微软雅黑", size=11)
            row_idx += 1

        try:
            wb.save(file_path)
        except Exception:
            pass

    def refresh_week_tree(self):
        """重建预览表。刚被改动的行用高亮色标出并滚动到可见位置。"""
        for item in self.tree_week.get_children():
            self.tree_week.delete(item)
        self.tree_week.tag_configure("changed", background=HILITE)

        highlight_sids = set(self.last_edited_sids) if self.last_edited_sids else set()
        if self.last_edited_sid:
            highlight_sids.add(self.last_edited_sid)

        highlight_item = None
        for sid, info in self.week_data.items():
            records = info.get('records') or []
            total = sum(r[0] for r in records)
            tags = ("changed",) if sid in highlight_sids else ()
            if not records:
                item = self.tree_week.insert(
                    "", "end", iid=f"{sid}@-1",
                    values=(sid, info['name'], self._fmt_score(total), "",
                            info['role'], info['dorm']),
                    tags=tags)
                if sid in highlight_sids and highlight_item is None:
                    highlight_item = item
                continue
            # 一人多条：逐条向下排，每条独占一行（以前用「|」挤在一格里，看不清也好改错）
            for idx, rec in enumerate(records):
                first = (idx == 0)
                values = (sid if first else "",
                          info['name'] if first else "",
                          self._fmt_score(total) if first else "",
                          str(rec[1]),
                          info['role'] if first else "",
                          info['dorm'] if first else "")
                item = self.tree_week.insert("", "end", iid=f"{sid}@{idx}",
                                             values=values, tags=tags)
                if first and sid in highlight_sids and highlight_item is None:
                    highlight_item = item
        if highlight_item:
            self.tree_week.see(highlight_item)
            self.tree_week.selection_set(highlight_item)

        # 高亮只保留一轮，下次刷新自动褪去
        self.last_edited_sids.clear()
        self.last_edited_sid = None

        # 寝室统计页的「AI 写入的寝室记录」跟着周次/数据一起刷新
        self._rebuild_dorm_ai_rows()

    @staticmethod
    def _parse_tree_iid(iid):
        """行标识 sid@记录序号（-1 = 这个学生本周还没有记录）。"""
        sid, _, idx = str(iid).partition("@")
        try:
            return sid, int(idx)
        except ValueError:
            return sid, -1

    def _on_tree_double_click(self, event):
        """双击编辑 总加分(#3) 或 项目明细(#4)。一人多条时只改被双击的那一条。"""
        if self.showing_inquiry:
            return
        region = self.tree_week.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree_week.identify_column(event.x)
        row = self.tree_week.identify_row(event.y)
        if not row or not column:
            return
        col_idx = int(column.replace("#", ""))
        if col_idx not in (3, 4):
            return
        sid, rec_idx = self._parse_tree_iid(row)
        if sid not in self.week_data:
            return
        records = self.week_data[sid]['records']
        if col_idx == 4:
            if not records:
                # 以前这里只是悄悄改状态栏，看着就像「双击没反应」
                messagebox.showinfo(
                    "这个学生还没有记录",
                    f"{self.week_data[sid]['name']} 这一周还没有加减分记录。\n\n"
                    "先在「总加分」列双击填个分值，再回来改明细。")
                return
            if not (0 <= rec_idx < len(records)):
                return
        values = self.tree_week.item(row, "values")
        if not values:
            return
        old_text = values[col_idx - 1] if col_idx - 1 < len(values) else ""
        bbox = self.tree_week.bbox(row, column)
        if not bbox:
            return
        x, y, w, h = bbox
        edit_entry = ttk.Entry(self.tree_week)
        edit_entry.place(x=x, y=y, width=w, height=h)
        edit_entry.insert(0, old_text)
        edit_entry.select_range(0, tk.END)
        edit_entry.focus_set()

        state = {"done": False}

        def cancel_edit(event=None):
            if state["done"]:
                return
            state["done"] = True
            edit_entry.destroy()

        def save_edit(event=None):
            if state["done"]:
                return
            state["done"] = True
            new_text = edit_entry.get().strip()
            edit_entry.destroy()
            if sid not in self.week_data:
                return
            records_now = self.week_data[sid]['records']
            if new_text == str(old_text).strip():
                return                      # 没改就不入撤回栈

            if col_idx == 3:
                # 编辑总加分：把差额记到被双击的那条记录上
                try:
                    new_total = self._eval_score(new_text)
                except (ValueError, SyntaxError):
                    messagebox.showwarning(
                        "分值看不懂", f"「{new_text}」不是分数。\n可以写 1、-0.5、10+2 这种。")
                    return
                old_total = sum(r[0] for r in records_now)
                diff = new_total - old_total
                if abs(diff) <= 0.001:
                    return
                self._push_undo()
                if 0 <= rec_idx < len(records_now):
                    rec = records_now[rec_idx]
                    records_now[rec_idx] = [rec[0] + diff, rec[1]] + list(rec[2:])
                elif records_now:
                    rec = records_now[-1]
                    records_now[-1] = [rec[0] + diff, rec[1]] + list(rec[2:])
                else:
                    # 原本没有任何记录：直接补一条，否则改了没反应
                    records_now.append([new_total, "手动调整", {"src": "manual"}])
            else:
                # 编辑项目明细：只改这一条的原因，不动分数
                if not (0 <= rec_idx < len(records_now)):
                    return
                if str(records_now[rec_idx][1]) == new_text:
                    return
                self._push_undo()
                rec = records_now[rec_idx]
                records_now[rec_idx] = [rec[0], new_text] + list(rec[2:])

            self.last_edited_sid = sid
            if not self.showing_inquiry:
                self.refresh_week_tree()

        edit_entry.bind("<Return>", save_edit)
        edit_entry.bind("<FocusOut>", save_edit)
        edit_entry.bind("<Escape>", cancel_edit)

    def undo_last_action(self):
        if not self.undo_stack:
            return
        self.redo_stack.append(deepcopy(self.week_data))
        self.week_data = self.undo_stack.pop()
        self.unsaved = True
        self.week_scanned = False
        self._mark_month_stale()
        if not self.showing_inquiry:
            self.refresh_week_tree()
        if self.current_view == "ai":
            self.refresh_ai_view()
        self.btn_undo.config(state="disabled" if not self.undo_stack else "normal")
        self.btn_redo.config(state="normal")
        self.week_status.config(text="已撤回上一步操作", foreground=WARN_C)

    def redo_last_action(self):
        if not self.redo_stack:
            return
        self.undo_stack.append(deepcopy(self.week_data))
        if len(self.undo_stack) > self.UNDO_LIMIT:
            del self.undo_stack[:len(self.undo_stack) - self.UNDO_LIMIT]
        self.week_data = self.redo_stack.pop()
        self.unsaved = True
        self.week_scanned = False
        self._mark_month_stale()
        if not self.showing_inquiry:
            self.refresh_week_tree()
        if self.current_view == "ai":
            self.refresh_ai_view()
        self.btn_redo.config(state="disabled" if not self.redo_stack else "normal")
        self.btn_undo.config(state="normal")
        self.week_status.config(text="已重做上一步操作", foreground=ACCENT)

    def save_week_data(self):
        folder = self.get_folder_path()
        if not os.path.exists(folder): os.makedirs(folder, exist_ok=True)
        file_path = os.path.join(folder, self.get_week_file_name())
        # 外部删除检测：加载时表还在、现在没了 → 明确提示，避免"记录复活"的困惑
        if getattr(self, "_week_file_existed", False) and not os.path.exists(file_path):
            notify(self.root, "磁盘上的本周周表已被外部删除，本次生成将按当前内容重建。",
                   title="检测到外部删除", kind="warn", duration=10000)
        m, w = self.current_month.get(), self.current_week.get()
        s_name = f"第{w}周"

        wb_target = None
        wb_tpl = None
        try:
            wb_tpl = openpyxl.load_workbook(self._template_path())
            ws_tpl = wb_tpl["周表模板"]

            if not os.path.exists(file_path):
                wb_target = openpyxl.Workbook()
                if wb_target.sheetnames: del wb_target[wb_target.sheetnames[0]]
            else:
                wb_target = openpyxl.load_workbook(file_path)
                if s_name in wb_target.sheetnames:
                    del wb_target[s_name]

            ws_new = wb_target.create_sheet(s_name)

            for col in ws_tpl.column_dimensions: ws_new.column_dimensions[col].width = ws_tpl.column_dimensions[
                col].width

            for r in range(1, 4):
                if ws_tpl.row_dimensions[r].height: ws_new.row_dimensions[r].height = ws_tpl.row_dimensions[r].height
                for c in range(1, 9):
                    c_tpl = ws_tpl.cell(row=r, column=c)
                    c_new = ws_new.cell(row=r, column=c)
                    val = c_tpl.value
                    if val and isinstance(val, str) and ('#' in val or '*' in val):
                        val = val.replace('#月', f'{m}月').replace('*月', f'{m}月').replace('#周', f'{w}周').replace(
                            '*周', f'{w}周').replace('#', w).replace('*', w)
                    c_new.value = val
                    if c_tpl.has_style:
                        c_new.font, c_new.border, c_new.fill, c_new.alignment = copy(c_tpl.font), copy(
                            c_tpl.border), copy(c_tpl.fill), copy(c_tpl.alignment)

            for m_range in ws_tpl.merged_cells.ranges:
                if getattr(m_range, 'max_row', 0) <= 3: ws_new.merge_cells(str(m_range))

            thin = Side(border_style="thin", color="000000")
            border = Border(top=thin, left=thin, right=thin, bottom=thin)
            align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            font = Font(name='宋体', size=11)

            cur_r = 4
            total_extra_rows = 0

            for l_id, r_id in self.base_layout:
                l_recs = self.week_data.get(l_id, {}).get('records', [])
                r_recs = self.week_data.get(r_id, {}).get('records', [])

                # 空占位用 None（真空单元格），写 "" 会让 Excel 把它当文本
                if not l_recs and l_id: l_recs = [(None, "")]
                if not r_recs and r_id: r_recs = [(None, "")]

                block_rows = max(1, len(l_recs), len(r_recs))
                total_extra_rows += (block_rows - 1)

                for r_idx in range(cur_r, cur_r + block_rows):
                    for c_idx in range(1, 9):
                        cell = ws_new.cell(row=r_idx, column=c_idx)
                        cell.value = None
                        cell.border = border
                        cell.alignment = align
                        cell.font = font

                if l_id:
                    ws_new.cell(row=cur_r, column=1, value=l_id)
                    ws_new.cell(row=cur_r, column=2, value=self.week_data[l_id]['name'])
                    if block_rows > 1:
                        ws_new.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r + block_rows - 1,
                                           end_column=1)
                        ws_new.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r + block_rows - 1,
                                           end_column=2)

                    for i in range(len(l_recs)):
                        ws_new.cell(row=cur_r + i, column=3, value=l_recs[i][0])
                        ws_new.cell(row=cur_r + i, column=4, value=l_recs[i][1])

                    if len(l_recs) < block_rows:
                        last_idx = max(0, len(l_recs) - 1)
                        start_merge = cur_r + last_idx
                        end_merge = cur_r + block_rows - 1
                        if start_merge < end_merge:
                            ws_new.merge_cells(start_row=start_merge, start_column=3, end_row=end_merge, end_column=3)
                            ws_new.merge_cells(start_row=start_merge, start_column=4, end_row=end_merge, end_column=4)

                if r_id:
                    ws_new.cell(row=cur_r, column=5, value=r_id)
                    ws_new.cell(row=cur_r, column=6, value=self.week_data[r_id]['name'])
                    if block_rows > 1:
                        ws_new.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r + block_rows - 1,
                                           end_column=5)
                        ws_new.merge_cells(start_row=cur_r, start_column=6, end_row=cur_r + block_rows - 1,
                                           end_column=6)

                    for i in range(len(r_recs)):
                        ws_new.cell(row=cur_r + i, column=7, value=r_recs[i][0])
                        ws_new.cell(row=cur_r + i, column=8, value=r_recs[i][1])

                    if len(r_recs) < block_rows:
                        last_idx = max(0, len(r_recs) - 1)
                        start_merge = cur_r + last_idx
                        end_merge = cur_r + block_rows - 1
                        if start_merge < end_merge:
                            ws_new.merge_cells(start_row=start_merge, start_column=7, end_row=end_merge, end_column=7)
                            ws_new.merge_cells(start_row=start_merge, start_column=8, end_row=end_merge, end_column=8)

                cur_r += block_rows

            if "周表模板2" in wb_tpl.sheetnames:
                ws_sig = wb_tpl["周表模板2"]
                sig_start = 1
                for r in range(4, ws_sig.max_row + 1):
                    if any(str(ws_sig.cell(r, c).value).strip() not in ['', 'None'] for c in range(1, 15)):
                        sig_start = r
                        break

                for r_src in range(sig_start, ws_sig.max_row + 1):
                    r_tgt = r_src + total_extra_rows
                    if ws_sig.row_dimensions[r_src].height:
                        ws_new.row_dimensions[r_tgt].height = ws_sig.row_dimensions[r_src].height

                    for c in range(1, 15):
                        c_src = ws_sig.cell(row=r_src, column=c)
                        if c_src.value is not None and str(c_src.value).strip() != "":
                            c_tgt = ws_new.cell(row=r_tgt, column=c)
                            c_tgt.value = c_src.value
                            if c_src.has_style:
                                c_tgt.font = copy(c_src.font)
                                c_tgt.border = copy(c_src.border)
                                c_tgt.alignment = copy(c_src.alignment)

                for m_range in ws_sig.merged_cells.ranges:
                    if m_range.min_row >= sig_start:
                        try:
                            ws_new.merge_cells(
                                start_row=m_range.min_row + total_extra_rows,
                                start_column=m_range.min_col,
                                end_row=m_range.max_row + total_extra_rows,
                                end_column=m_range.max_col
                            )
                        except Exception:
                            pass

            while True:
                try:
                    wb_target.save(file_path)
                    self.unsaved = False
                    # duration 给足 15 秒：鼠标移到提示上会暂停倒计时，够点按钮
                    notify(self.root, f"周纪实表已生成：{os.path.basename(file_path)}",
                           title="生成完成", kind="success", duration=15000,
                           action=("打开文件", lambda p=file_path: self._open_path(p)))
                    self.week_status.config(text=f"已生成第{m}月第{w}周周表", foreground=OK_C)
                    self._remember_week_file_state()
                    self.week_scanned = False
                    self._mark_month_stale()
                    self.check_month_ready()
                    break
                except PermissionError:
                    if not self.ask_retry_save(file_path): break
                except Exception as e:
                    messagebox.showerror("写入失败", f"发生错误：\n{e}")
                    break
        except Exception as e:
            messagebox.showerror("系统异常", f"生成时遇到问题：\n{e}")
        finally:
            if wb_target: wb_target.close()
            if wb_tpl: wb_tpl.close()

    def scan_weeks_for_month(self):
        # 步骤 1 是直接读磁盘上的周表文件，内存里没生成的录入不会被算进去
        if self.unsaved:
            choice = messagebox.askyesnocancel(
                "有未保存的录入",
                "当前周的加减分还没写进周表文件。\n\n"
                "【是】先生成周表再读取（推荐）\n"
                "【否】忽略这些录入，直接读取已有文件\n"
                "【取消】返回")
            if choice is None:
                return
            if choice:
                self.save_week_data()
                if self.unsaved:      # 保存被取消或失败
                    return

        for sid in self.month_sum:
            self.month_sum[sid] = 0.0
        folder = self.get_folder_path()
        file_path = os.path.join(folder, self.get_week_file_name())

        if not os.path.exists(file_path):
            return messagebox.showerror("文件缺失", "找不到本月周表，请先生成！")

        found_weeks = 0
        wb = None
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sht in wb.sheetnames:
                if "周" in sht:
                    found_weeks += 1
                    ws = wb[sht]
                    # 签名栏边界按栏独立判断：左栏看 A/B 列（班主任签名在 A），
                    # 右栏看 E/F 列（考评小组签名合并格在 E）。全局"任一列出现
                    # 签名就停"会误伤与签名区同行的左/右栏最后一行学生（王小领 bug）。
                    def _sig_row(cols):
                        for r in range(4, ws.max_row + 1):
                            for c in cols:
                                v = str(ws.cell(row=r, column=c).value or "")
                                if ("签名" in v) or ("考评" in v):
                                    return r
                        return ws.max_row + 1
                    sig_row_left = _sig_row((1, 2))
                    sig_row_right = _sig_row((5, 6))
                    active_l_id = None
                    active_r_id = None
                    for r in range(4, ws.max_row + 1):
                        if r >= sig_row_left:
                            active_l_id = None      # 左栏已进签名区
                        else:
                            c1_val = ws.cell(row=r, column=1).value
                            if c1_val is not None and str(c1_val).strip() != "":
                                active_l_id = self.clean_val(c1_val)
                            if active_l_id and active_l_id in self.month_sum:
                                num = self._to_number(ws.cell(row=r, column=3).value)
                                if num is not None:
                                    self.month_sum[active_l_id] += num
                        if ws.max_column >= 7 and r < sig_row_right:
                            c5_val = ws.cell(row=r, column=5).value
                            if c5_val is not None and str(c5_val).strip() != "":
                                active_r_id = self.clean_val(c5_val)
                            if active_r_id and active_r_id in self.month_sum:
                                num = self._to_number(ws.cell(row=r, column=7).value)
                                if num is not None:
                                    self.month_sum[active_r_id] += num
            if found_weeks > 0:
                self.week_scanned = True
                self.lbl_scan_status.config(text=f"成功读取 {found_weeks} 个周次", foreground=OK_C)
            else:
                self.week_scanned = False
                self.lbl_scan_status.config(text="该文件里还没有任何周次", foreground=WARN_C)
                messagebox.showwarning("没有可累加的周次",
                                       f"文件里没有找到「第N周」工作表：\n{file_path}\n\n"
                                       f"请先在周纪实表页面生成周表。")
        except Exception as e:
            messagebox.showerror("读取出错", str(e))
        finally:
            if wb: wb.close()
        self.check_month_ready()

    def load_last_month_file(self):
        # openpyxl 读不了老式 .xls，只提供 .xlsx，避免选了文件才报错
        filepath = filedialog.askopenfilename(
            title="选择上个月的月度总表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
        if not filepath:
            return
        if os.path.splitext(filepath)[1].lower() not in (".xlsx", ".xlsm"):
            return messagebox.showerror(
                "格式不支持",
                "只能读取 .xlsx 文件。\n如果是老的 .xls，请先在 Excel 里另存为 .xlsx。")
        for sid in self.last_month_scores:
            self.last_month_scores[sid] = 0.0
        wb = None
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.worksheets[0]
            loaded_count = 0
            for r in range(2, ws.max_row + 1):
                c1 = ws.cell(row=r, column=1)
                if type(c1).__name__ == 'MergedCell':
                    continue
                sid = self.clean_val(c1.value)
                if sid in self.last_month_scores:
                    num = self._to_number(ws.cell(row=r, column=5).value)
                    if num is not None:
                        self.last_month_scores[sid] = num
                        loaded_count += 1
            self.last_month_loaded = True
            self.lbl_last_status.config(text=f"提取了 {loaded_count} 人的积分", foreground=OK_C)
        except Exception as e:
            self.last_month_loaded = False
            messagebox.showerror("解析失败", f"错误：{e}")
        finally:
            if wb: wb.close()
        self.check_month_ready()

    def check_month_ready(self):
        self.refresh_month_tree()
        if self.is_first_month.get():
            self.lbl_last_status.config(text="首月模式：每人底分按 80 计", foreground=OK_C)
            self.btn_load_last.config(state="disabled")
        else:
            self.btn_load_last.config(state="normal")
            if self.last_month_loaded:
                self.lbl_last_status.config(text="已导入上月底分", foreground=OK_C)
            else:
                self.lbl_last_status.config(text="未导入上月底分", foreground=ERR_C)

        # 按钮文案直接告诉用户还缺哪一步，避免「未准备就绪」这种含糊提示
        if not self.week_scanned:
            self.btn_generate_month.config(state="disabled", text="请先点步骤 1 读取周表")
        elif not (self.is_first_month.get() or self.last_month_loaded):
            self.btn_generate_month.config(state="disabled", text="请先完成步骤 2 或勾选首月模式")
        else:
            self.btn_generate_month.config(state="normal", text="生成月度总表")

    def refresh_month_tree(self):
        for item in self.tree_month.get_children():
            self.tree_month.delete(item)
        for sid in self.week_data.keys():
            name = self.week_data[sid]['name']
            base_sc = 80.0 if self.is_first_month.get() else self.last_month_scores.get(sid, 0.0)
            add_sc = self.month_sum.get(sid, 0.0)
            total = base_sc + add_sc
            self.tree_month.insert("", "end", values=(sid, name, self._fmt_score(base_sc),
                                                      self._fmt_score(add_sc),
                                                      self._fmt_score(total)))

    def _find_dsoffice(self):
        """找 DeepSeek Harness 的 LibreOffice Kit（渲染 Excel 效果图用）。

        返回 (node, cli) 或 None。用户机器装了 DeepSeek Harness 就有。
        """
        if getattr(self, "_dsoffice_cache", None):
            return self._dsoffice_cache
        import glob
        node = r"D:\DeepSeek Harness\resources\runtime\primary-runtime\dependencies\node\bin\node.exe"
        cli = r"D:\DeepSeek Harness\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"
        if os.path.exists(node) and os.path.exists(cli):
            self._dsoffice_cache = (node, cli)
            return self._dsoffice_cache
        for n in glob.glob(r"D:\DeepSeek Harness\resources\runtime\*\dependencies\node\bin\node.exe"):
            for c in glob.glob(r"D:\DeepSeek Harness\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"):
                self._dsoffice_cache = (n, c)
                return self._dsoffice_cache
        return None

    def _capture_table_png(self, tree, base_name):
        """回退方案：屏幕截取表格控件当前可见区域（渲染组件不可用时）。"""
        from PIL import ImageGrab
        tree.update_idletasks()
        try:
            tree.yview_moveto(0)
            tree.update_idletasks()
        except Exception:
            pass
        children = tree.get_children("")
        x, y = tree.winfo_rootx(), tree.winfo_rooty()
        w = tree.winfo_width()
        h = tree.winfo_height()
        if children:
            last = tree.bbox(children[-1])
            if last:
                h = min(h, last[1] + last[3] + 4)
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
        folder = os.path.join(self.get_folder_path(), "截图")
        os.makedirs(folder, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        path = os.path.join(folder, f"{base_name}_{stamp}.png")
        img.save(path)
        return path

    def _render_sheet_png(self, xlsx_path, sheet, base_name):
        """把 xlsx 里指定工作表渲染成高清 PNG（LibreOffice Kit 300 DPI）。

        多页时纵向拼接为一张长图。返回保存路径；渲染器不可用抛 RuntimeError。
        """
        from PIL import Image
        kit = self._find_dsoffice()
        if not kit:
            raise RuntimeError("未找到渲染组件（LibreOffice Kit）")
        node, cli = kit
        if not os.path.exists(xlsx_path):
            raise FileNotFoundError(f"找不到生成表文件：{xlsx_path}")
        wb = openpyxl.load_workbook(xlsx_path, read_only=True)
        try:
            if sheet not in wb.sheetnames:
                raise ValueError(f"工作表「{sheet}」不存在（现有：{'、'.join(wb.sheetnames)}）")
            ws = wb[sheet]
            max_row, max_col = ws.max_row, ws.max_column
        finally:
            wb.close()
        from openpyxl.utils import get_column_letter
        rng = f"A1:{get_column_letter(max_col)}{max_row}"
        out_root = os.path.join(self.get_folder_path(), "截图")
        os.makedirs(out_root, exist_ok=True)
        tmp = os.path.join(out_root, f".render-{datetime.datetime.now().strftime('%H%M%S%f')}")
        flags = 0x08000000  # CREATE_NO_WINDOW
        proc = subprocess.run([node, cli, "render", "--input", xlsx_path,
                               "--output-dir", tmp, "--sheet", sheet,
                               "--range", rng, "--dpi", "300"],
                              capture_output=True, timeout=240, creationflags=flags)
        pages = sorted(glob.glob(os.path.join(tmp, "page-*.png")))
        if not pages:
            err = (proc.stderr or b"").decode("utf-8", "replace")[-200:]
            raise RuntimeError(f"渲染失败：{err or '无输出'}")
        imgs = [Image.open(pg).convert("RGB") for pg in pages]
        if len(imgs) == 1:
            final_img = imgs[0]
        else:
            gap = 6
            total_w = max(im.width for im in imgs)
            total_h = sum(im.height for im in imgs) + gap * (len(imgs) - 1)
            final_img = Image.new("RGB", (total_w, total_h), "#FFFFFF")
            yy = 0
            for im in imgs:
                final_img.paste(im, (0, yy))
                yy += im.height + gap
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        final = os.path.join(out_root, f"{base_name}_{stamp}.png")
        final_img.save(final)
        for im in imgs:
            im.close()
        shutil.rmtree(tmp, ignore_errors=True)
        return final

    def screenshot_tree(self, tree, base_name, mode="preview"):
        """截图入口：渲染生成的 Excel 文件（内容完整，首选）；
        渲染组件不可用时回退为屏幕截取预览。"""
        # 判断目标文件
        is_week = base_name.startswith("周纪实表截图")
        try:
            if is_week:
                xlsx = os.path.join(self.get_folder_path(), self.get_week_file_name())
                sheet = f"第{self.current_week.get()}周"
            else:
                xlsx = os.path.join(self.get_folder_path(), self.get_month_file_name())
                sheet = f"{self.current_month.get()}月总表"
            path = self._render_sheet_png(xlsx, sheet, base_name)
            kind = "success"
            msg = f"已生成表格图片：{path}"
        except Exception as exc:
            # 回退：屏幕截取预览区
            try:
                path = self._capture_table_png(tree, base_name)
                kind = "warn"
                msg = f"渲染组件不可用，已截取屏幕预览：{path}"
            except Exception as exc2:
                return messagebox.showerror("截图失败", f"渲染：{exc}\n截屏：{exc2}")
        try:
            notify(self.root, msg, title="截图完成", kind=kind, duration=8000,
                   action=("打开文件夹", lambda p=os.path.dirname(path): self._open_folder(p)))
        except Exception:
            pass

    def _apply_ai_screenshot(self, batch):
        """AI 请求截图：渲染生成的 Excel 为高清 PNG，结果路径写入
        screenshot_result.json 供 AI 读取。"""
        target = str(batch.get("target") or "全部")
        results, errors = [], []

        def do(target_kind, xlsx, sheet, base):
            try:
                results.append(self._render_sheet_png(xlsx, sheet, base))
            except Exception as exc:
                errors.append(f"{target_kind}: {exc}")

        if target in ("周", "全部"):
            do("周", os.path.join(self.get_folder_path(), self.get_week_file_name()),
               f"第{self.current_week.get()}周",
               f"周纪实表截图_{self.class_name}{self.current_month.get()}月第{self.current_week.get()}周")
        if target in ("月", "全部") or (target == "月"):
            do("月", os.path.join(self.get_folder_path(), self.get_month_file_name()),
               f"{self.current_month.get()}月总表",
               f"月度总表截图_{self.class_name}{self.current_month.get()}月")
        ok = bool(results)
        try:
            with open(os.path.join(self._bridge_dir(), "screenshot_result.json"),
                      "w", encoding="utf-8") as f:
                json.dump({"ok": ok,
                           "files": results,
                           "errors": errors}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        try:
            if ok:
                notify(self.root, f"已按 AI 要求生成 {len(results)} 张表格图片（截图文件夹）",
                       title="AI 截图", kind="success", duration=6000)
        except Exception:
            pass
        return ok

    def open_data_folder(self):
        """侧边栏「打开文件夹」：打开本月生成表所在的大文件夹（不存在则先创建）。"""
        folder = self.get_folder_path()
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception:
            folder = self._base_dir()
        self._open_folder(folder)

    def _apply_ai_screenshot(self, batch):
        """AI 请求截图：target 周/月/全部，结果路径写入 screenshot_result.json 供 AI 读取。"""
        target = str(batch.get("target") or "全部")
        results = []
        try:
            if target in ("周", "全部"):
                results.append(self._capture_table_png(
                    self.tree_week,
                    f"周纪实表截图_{self.class_name}{self.current_month.get()}月第{self.current_week.get()}周"))
            if target in ("月", "全部"):
                results.append(self._capture_table_png(
                    self.tree_month,
                    f"月度总表截图_{self.class_name}{self.current_month.get()}月"))
        except Exception as exc:
            try:
                with open(os.path.join(self._bridge_dir(), "screenshot_result.json"),
                          "w", encoding="utf-8") as f:
                    json.dump({"ok": False, "error": str(exc)}, f, ensure_ascii=False)
            except Exception:
                pass
            return False
        try:
            with open(os.path.join(self._bridge_dir(), "screenshot_result.json"),
                      "w", encoding="utf-8") as f:
                json.dump({"ok": True, "files": results}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        try:
            notify(self.root, f"已按 AI 要求截图 {len(results)} 张（存在「截图」文件夹）",
                   title="AI 截图", kind="success", duration=6000)
        except Exception:
            pass
        return True

    def save_month_data(self):
        folder = self.get_folder_path()
        if not os.path.exists(folder): os.makedirs(folder, exist_ok=True)
        file_path = os.path.join(folder, self.get_month_file_name())
        y, m = self.current_year.get(), self.current_month.get()
        wb = None
        try:
            wb = openpyxl.load_workbook(self._template_path())
            ws = wb["月表模板"]
            ws.title = f"{m}月总表"
            for r in range(2, ws.max_row + 1):
                c1 = ws.cell(row=r, column=1)
                if type(c1).__name__ == 'MergedCell': continue
                sid = self.clean_val(c1.value)
                if sid in self.month_sum:
                    c4 = ws.cell(row=r, column=4)
                    if type(c4).__name__ != 'MergedCell':
                        v = str(c4.value)
                        if '#' in v or '*' in v: c4.value = f"{y}年{m}月"
                    c5 = ws.cell(row=r, column=5)
                    if type(c5).__name__ != 'MergedCell':
                        base_sc = 80.0 if self.is_first_month.get() else self.last_month_scores.get(sid, 0.0)
                        total = base_sc + self.month_sum.get(sid, 0.0)
                        # 0 分也要写 0，不能空着——空白和「0 分」在统计上是两回事
                        c5.value = round(total, 2) if total != int(total) else int(total)
            for sht in wb.sheetnames:
                if sht != ws.title: del wb[sht]
            while True:
                try:
                    wb.save(file_path)
                    notify(self.root, f"月度总表已生成：{os.path.basename(file_path)}",
                           title="生成完成", kind="success", duration=15000,
                           action=("打开文件", lambda p=file_path: self._open_path(p)))
                    break
                except PermissionError:
                    if not self.ask_retry_save(file_path): break
                except Exception as e:
                    messagebox.showerror("写入失败", f"发生错误：\n{e}")
                    break
        except Exception as e:
            messagebox.showerror("系统异常", f"生成时遇到问题：\n{e}")
        finally:
            if wb: wb.close()


if __name__ == "__main__":
    enable_dpi_awareness()
    try:
        # TkinterDnD 根窗口：支持把名单模板 xlsx/csv 拖进窗口导入
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()          # 没有拖拄件时功能照常，只是不支持拖入
    apply_theme(root, accent=ACCENT, accent_hover=ACCENT_HOVER)
    configure_app_styles(root)
    app = MoralScoreApp(root)
    root.mainloop()