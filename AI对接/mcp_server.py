# -*- coding: utf-8 -*-
"""德育分自动化系统 · MCP 服务（stdlib 实现，仅额外依赖 openpyxl）

作用：让任意支持 MCP 的 AI 客户端（Cherry Studio / Claude Desktop / Cursor / Trae 等）
通过对话把加减分记录写进「德育分自动化系统」。

原理：
    AI 调 add_records  →  本服务把记录写成 _ai_bridge/inbox/*.json
    软件前台每 2 秒扫描 inbox  →  匹配当前周期的批次 → 填入并高亮

传输：stdio，JSON-RPC 2.0，每行一条消息（stdout 只放协议消息，日志走 stderr）。
"""

import datetime
import json
import os
import re
import sys
import traceback

# 冻结成 exe 后 __file__ 指向解包临时目录，必须改用 sys.executable 定位。
# 布局：<安装目录>\AI对接\mcp_server(.exe|.py) → 上溯两级 = 安装目录
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(sys.executable)))
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BRIDGE_DIR = os.path.join(BASE_DIR, "_ai_bridge")
INBOX_DIR = os.path.join(BRIDGE_DIR, "inbox")
STATE_FILE = os.path.join(BRIDGE_DIR, "state.json")
TEMPLATE = os.path.join(BASE_DIR, "名单模板.xlsx")
LOG_FILE = os.path.join(BRIDGE_DIR, "mcp_server.log")

PROTOCOL_VERSION = "2025-06-18"
SUPPORTED_VERSIONS = {"2024-11-05", "2025-03-26", "2025-06-18"}
SERVER_INFO = {"name": "deyufen-mcp", "title": "德育分自动化系统", "version": "1.0.0"}

STAR_TEXT_MAP = {"三星": 3, "四星": 4, "五星": 5, "3星": 3, "4星": 4, "5星": 5}


# ────────────────────────────── 基础设施 ──────────────────────────────
def log(message):
    """诊断日志写 stderr + 文件，绝不能写到 stdout（那是协议通道）。"""
    line = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}"
    print(line, file=sys.stderr, flush=True)
    try:
        os.makedirs(BRIDGE_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def read_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json_atomic(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ────────────────────────────── 数据读取 ──────────────────────────────
def open_template():
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"找不到名单模板：{TEMPLATE}")
    import openpyxl
    return openpyxl.load_workbook(TEMPLATE, data_only=True)


def clean(value):
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def norm_name(value):
    """与主程序 clean_name 保持一致的姓名归一化（含 奕→弈 的历史兼容）。"""
    return re.sub(r"\s+", "", str(value or "")).replace("奕", "弈")


# ─────────── 模板缓存：不再每次工具调用都重新解析 Excel ───────────
# 原来 get_context / list_students / add_records 每个工具都 load_workbook 一次，
# 一次对话下来白烧 1 秒多。模板没变就直接复用。
_CACHE = {"stamp": None, "roster": None, "class_name": None}


def _template_stamp():
    """模板的 (修改时间, 大小)。变了才重读。"""
    try:
        st = os.stat(TEMPLATE)
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def load_roster():
    """读名单：学号、姓名、寝室号、寝室角色（带缓存）。"""
    stamp = _template_stamp()
    if _CACHE["stamp"] == stamp and _CACHE["roster"] is not None:
        return _CACHE["roster"]
    wb = open_template()
    try:
        roster = []
        by_name = {}
        roles = {}
        dorms = {}

        if "寝室名单" in wb.sheetnames:
            ws = wb["寝室名单"]
            for r in range(1, ws.max_row + 1):
                room = clean(ws.cell(r, 1).value)
                if not room or "房间" in room:
                    continue
                members = clean(ws.cell(r, 3).value)
                leader = clean(ws.cell(r, 7).value)
                for text, role in ((leader, "室长"), (members, "成员")):
                    if not text:
                        continue
                    for name in text.replace("，", " ").replace(",", " ").split():
                        key = norm_name(name)
                        if key:
                            dorms[key] = room
                            roles.setdefault(key, role)

        ws = wb["月表模板"]
        for r in range(2, ws.max_row + 1):      # 第 1 行是表头
            sid = clean(ws.cell(r, 1).value)
            name = clean(ws.cell(r, 2).value)
            if not sid or not name or "签名" in sid or "考评" in sid:
                continue
            key = norm_name(name)
            item = {
                "sid": sid,
                "name": name,
                "dorm": dorms.get(key, "未知"),
                "role": roles.get(key, "成员"),
            }
            roster.append(item)
            by_name[name] = item
        _CACHE["stamp"] = stamp
        _CACHE["roster"] = roster
        return roster
    finally:
        wb.close()


def load_class_name():
    stamp = _template_stamp()
    if _CACHE["stamp"] == stamp and _CACHE["class_name"] is not None:
        return _CACHE["class_name"]
    wb = open_template()
    try:
        ws = wb["月表模板"]
        raw = clean(ws.cell(2, 3).value)
        if raw and "班" not in raw:
            raw += "班"
        name = raw or "未命名班级"
        _CACHE["stamp"] = stamp
        _CACHE["class_name"] = name
        return name
    finally:
        wb.close()


def week_file_path(year, month, class_name):
    folder = os.path.join(BASE_DIR, f"{year}年{month}月")
    return os.path.join(folder, f"{class_name}{month}月德育考评周纪实表.xlsx")


def load_week_records(year, month, week, roster):
    """读某周每个学生已有的记录：{sid: [{score, reason}]}"""
    class_name = load_class_name()
    path = week_file_path(year, month, class_name)
    result = {s["sid"]: [] for s in roster}
    if not os.path.exists(path):
        return result, False

    import openpyxl
    wb = None
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = f"第{week}周"
        if sheet not in wb.sheetnames:
            return result, True
        ws = wb[sheet]

        sig_row = ws.max_row + 1
        for r in range(4, ws.max_row + 1):
            texts = [clean(ws.cell(r, c).value) for c in range(1, 9)]
            if any("考评" in t or "签名" in t for t in texts):
                sig_row = r
                break

        left = right = None
        for r in range(4, sig_row):
            cell_l = clean(ws.cell(r, 1).value)
            if cell_l:
                left = cell_l
            if left in result:
                score = ws.cell(r, 3).value
                reason = clean(ws.cell(r, 4).value)
                if score is not None and clean(score) != "":
                    try:
                        result[left].append({"score": float(score), "reason": reason})
                    except (TypeError, ValueError):
                        pass
            cell_r = clean(ws.cell(r, 5).value)
            if cell_r:
                right = cell_r
            if right in result:
                score = ws.cell(r, 7).value
                reason = clean(ws.cell(r, 8).value)
                if score is not None and clean(score) != "":
                    try:
                        result[right].append({"score": float(score), "reason": reason})
                    except (TypeError, ValueError):
                        pass
        return result, True
    finally:
        if wb:
            wb.close()


def load_rules():
    """读加分预览表（含寝室规则、早自习），返回二维文本表。"""
    wb = open_template()
    try:
        if "加分预览" not in wb.sheetnames:
            return []
        ws = wb["加分预览"]
        rows = []
        for r in range(1, ws.max_row + 1):
            values = [clean(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
            if any(values):
                rows.append(values)
        return rows
    finally:
        wb.close()


# ─────────── 加分标准：把「加分预览」解析成结构化规则，供自动取分 ───────────
# 用户经常只说「方世杰 省赛一等奖」，不给分数。这里把模板里的加分标准解析成
# 程序可查的表，add_records 遇到没给 score 的记录就自己按表取分。
_RULES_CACHE = {"stamp": None, "table": None}


def _num(value):
    """单元格值转 float，非数字返回 None。"""
    try:
        text = clean(value)
        if not text:
            return None
        return float(re.sub(r"[^\d.\-+]", "", text))
    except (TypeError, ValueError):
        return None


def load_score_rules():
    """解析 名单模板.xlsx → 加分预览，返回结构化规则（带缓存）。"""
    stamp = _template_stamp()
    if _RULES_CACHE["stamp"] == stamp and _RULES_CACHE["table"] is not None:
        return _RULES_CACHE["table"]

    rows = load_rules()
    table = {"contest": {}, "dorm": {}, "morning": {}, "tags": {}}
    if rows:
        width = max(len(r) for r in rows)
        grid = [list(r) + [""] * (width - len(r)) for r in rows]

        def cell(r, c):
            if 0 <= r < len(grid) and 0 <= c < width:
                return clean(grid[r][c])
            return ""

        skip_rows = set()

        # 1) 竞赛表：A 列是「等级」，右侧是奖项，往下是赛事级别
        head = next((r for r in range(len(grid)) if cell(r, 0) == "等级"), None)
        if head is not None:
            prizes = []
            for c in range(1, width):
                name = cell(head, c)
                if not name:
                    break
                prizes.append((c, name))
            skip_rows.add(head)
            r = head + 1
            while r < len(grid):
                level = cell(r, 0)
                if not level or level in ("寝室", "标兵、荣誉称号"):
                    break
                scores = {}
                for c, prize in prizes:
                    v = _num(cell(r, c))
                    if v is not None:
                        scores[prize] = v
                if scores:
                    table["contest"][level] = scores
                skip_rows.add(r)
                r += 1

        # 2) 寝室表：A 列「寝室」，下一行是成员/室长，上方行标「周」「月」
        dorm_head = next((r for r in range(len(grid)) if cell(r, 0) == "寝室"), None)
        if dorm_head is not None:
            period_row = [cell(dorm_head, c) for c in range(width)]
            role_row = dorm_head + 1
            role_cols = []
            for c in range(width):
                role = cell(role_row, c)
                if role in ("寝室成员", "寝室长", "成员", "室长"):
                    period = ""
                    for k in range(c, -1, -1):
                        if period_row[k] in ("周", "月"):
                            period = period_row[k]
                            break
                    role_cols.append((c, period or "周", "室长" if "长" in role else "成员"))
            star_rows = set()
            r = role_row + 1
            while r < len(grid):
                star = cell(r, 0)
                if not star or star in ("标兵、荣誉称号",):
                    break
                star_rows.add(r)
                values = {}
                for c, period, role in role_cols:
                    v = _num(cell(r, c))
                    if v is not None:
                        values.setdefault(period, {})[role] = v
                if values:
                    table["dorm"][star] = values
                r += 1
            # 只跳过寝室块本身，别把下面的「标兵」「活动类比赛」也吞了
            skip_rows = ({r for r in skip_rows if r <= role_row}
                         | star_rows | {dorm_head, role_row})

            # 3) 早自习：同一行右侧「迟到 -0.5」「缺勤 -1」
            for r in range(dorm_head, min(len(grid), dorm_head + 6)):
                for c in range(width - 1):
                    label = cell(r, c)
                    if label in ("迟到", "缺勤", "早退"):
                        v = _num(cell(r, c + 1))
                        if v is not None:
                            table["morning"][label] = v

        # 4) 其余「标签在上、数值在下」的项：标兵、荣誉称号 / 活动类比赛的一等二等
        for r in range(len(grid) - 1):
            if r in skip_rows:
                continue
            for c in range(width):
                label = cell(r, c)
                if not label:
                    continue
                v = _num(cell(r + 1, c))
                if v is not None and not re.fullmatch(r"[-+]?\d+(\.\d+)?", label):
                    table["tags"].setdefault(label, v)

    _RULES_CACHE["stamp"] = stamp
    _RULES_CACHE["table"] = table
    return table


# 同义词：用户和 AI 的说法五花八门，都收敛到模板里的写法
LEVEL_ALIASES = (
    ("国赛", ("国赛", "国家级", "全国", "国家")),
    ("省级", ("省赛", "省级", "全省")),
    ("市级", ("市赛", "市级", "全市")),
    ("院级", ("院赛", "院级", "校赛", "校级", "学院", "学校")),
)
PRIZE_ALIASES = (
    ("一等奖", ("一等奖", "一等", "金牌", "冠军", "第一名")),
    ("二等奖", ("二等奖", "二等", "银牌", "亚军", "第二名")),
    ("三等奖", ("三等奖", "三等", "铜牌", "季军", "第三名")),
    ("四等奖", ("四等奖", "四等", "第四名")),
    ("五等奖", ("五等奖", "五等", "第五名")),
    ("六等奖", ("六等奖", "六等", "第六名")),
    ("优胜", ("优胜", "优秀奖", "入围", "参与奖")),
)


def _find_prize(text, aliases):
    """按优先级找奖项，返回模板里的列名（一等奖…优胜）。"""
    for key, words in aliases:
        for word in words:
            if word in text:
                return key
    return None


def _find_contest_level(text):
    for key, words in LEVEL_ALIASES:
        for word in words:
            if word in text:
                return key
    return None


def _match_key(table_dict, wanted):
    """模板里的列名可能带别的字（如「一等」），按前缀/包含关系对上。"""
    if wanted in table_dict:
        return wanted
    for key in table_dict:
        if key.startswith(wanted) or wanted.startswith(key):
            return key
    return None


def lookup_score(text):
    """从模板加分标准里推断分值。返回 (分值, 依据) 或 (None, None)。"""
    rules = load_score_rules()
    raw = re.sub(r"\s+", "", str(text or ""))
    if not raw:
        return None, None
    compact = raw.replace("级", "").replace("赛", "").replace("奖", "")

    # 1) 竞赛：级别 + 奖项
    level = _find_contest_level(raw) or _find_contest_level(compact)
    prize = _find_prize(raw, PRIZE_ALIASES)
    if level and prize and rules["contest"].get(level):
        key = _match_key(rules["contest"][level], prize)
        if key is not None:
            return rules["contest"][level][key], f"{level}·{key}"

    # 2) 寝室星级：五星寝室 / 三星宿舍，缺省按周、按成员
    star = next((s for s in ("五星", "四星", "三星") if s in raw), None)
    if star and ("寝室" in raw or "宿舍" in raw) and rules["dorm"].get(star):
        period = "月" if "月" in raw else "周"
        role = "室长" if ("室长" in raw or "寝室长" in raw or "宿舍长" in raw) else "成员"
        value = (rules["dorm"][star].get(period) or {}).get(role)
        if value is not None:
            return value, f"{star}寝室·{period}·{role}"

    # 3) 早自习
    for label in ("迟到", "缺勤", "早退"):
        if label in raw and label in rules["morning"]:
            return rules["morning"][label], f"早自习·{label}"

    # 4) 活动类比赛：一等/二等/三等/参与
    if "活动" in raw or "比赛" in raw:
        for word in ("参与", "一等", "二等", "三等"):
            if word in raw:
                key = _match_key(rules["tags"], word)
                if key is not None and rules["tags"].get(key) is not None:
                    return rules["tags"][key], f"活动类比赛·{word}"

    # 5) 标兵 / 荣誉称号
    for word in ("标兵", "荣誉称号"):
        if word in raw and "标兵、荣誉称号" in rules["tags"]:
            return rules["tags"]["标兵、荣誉称号"], "标兵、荣誉称号"

    # 6) 兜底：模板里的标签整串出现在原因里
    for label, value in rules["tags"].items():
        if label and label in raw:
            return value, label

    return None, None


def load_month_summary(year, month, roster):
    """累加该月所有周次的分数。"""
    class_name = load_class_name()
    path = week_file_path(year, month, class_name)
    summary = {s["sid"]: 0.0 for s in roster}
    if not os.path.exists(path):
        return summary, 0

    import openpyxl
    wb = None
    weeks = 0
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        for name in wb.sheetnames:
            if "周" not in name:
                continue
            weeks += 1
            ws = wb[name]
            sig_row = ws.max_row + 1
            for r in range(4, ws.max_row + 1):
                texts = [clean(ws.cell(r, c).value) for c in range(1, 9)]
                if any("考评" in t or "签名" in t for t in texts):
                    sig_row = r
                    break
            left = right = None
            for r in range(4, sig_row):
                cell_l = clean(ws.cell(r, 1).value)
                if cell_l:
                    left = cell_l
                if left in summary:
                    try:
                        summary[left] += float(ws.cell(r, 3).value or 0)
                    except (TypeError, ValueError):
                        pass
                cell_r = clean(ws.cell(r, 5).value)
                if cell_r:
                    right = cell_r
                if right in summary:
                    try:
                        summary[right] += float(ws.cell(r, 7).value or 0)
                    except (TypeError, ValueError):
                        pass
        return summary, weeks
    finally:
        if wb:
            wb.close()


def read_app_state():
    """读软件写的心跳状态。超过 90 秒没更新视为软件未运行。"""
    state = read_json(STATE_FILE, None) or {}
    running = False
    if state.get("running"):
        try:
            ts = datetime.datetime.fromisoformat(state.get("updated", ""))
            running = (datetime.datetime.now() - ts).total_seconds() < 90
        except Exception:
            running = False
    state["running"] = running
    return state


# ────────────────────────────── 学生匹配 ──────────────────────────────
def name_similarity(a, b):
    """姓名相似度 0~1。

    面向「拍照 OCR 认错字 / 手写看错一个字」的场景：
      · 整串字符序列相似度
      · 共同字符占比
      · 同姓额外加成（中文姓名姓氏权重最高）
      · 长度差太多则扣分，避免「张三」匹配到「张三丰」这种长名字
    """
    from difflib import SequenceMatcher

    a, b = norm_name(a), norm_name(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ratio = SequenceMatcher(None, a, b).ratio()
    common = len(set(a) & set(b)) / max(len(a), len(b))
    bonus = 0.10 if a[0] == b[0] else 0.0
    if abs(len(a) - len(b)) >= 2:
        bonus -= 0.10
    return max(0.0, min(1.0, ratio * 0.72 + common * 0.28 + bonus))


# 自动近似匹配的门槛
AUTO_MATCH_RATE = 0.72      # 最高相似度到这个值才敢自动写入
AUTO_MATCH_GAP = 0.06       # 还要和第二名拉开这么多，避免认错人
SUGGEST_RATE = 0.45         # 低于这个连候选都不提示


def resolve_student(roster, key):
    """把「姓名 / 学号 / 学号后两位」解析成学生。

    返回 (student, candidates, match)：
      · 精确命中          → (s, [], None)
      · 唯一包含/学号尾号 → (s, [], {"fuzzy": True, "input": ..., "rate": ...})
      · 近似命中（唯一）  → (s, [], {"fuzzy": True, "input": ..., "rate": ...})
      · 多人 / 找不到     → (None, [候选...], None)
    """
    raw = clean(key).replace(" ", "")
    if not raw:
        return None, [], None

    # 1) 精确：姓名 或 完整学号
    exact = [s for s in roster if s["name"] == raw or s["sid"] == raw]
    if len(exact) == 1:
        return exact[0], [], None
    if len(exact) > 1:
        return None, exact, None

    # 2) 学号后两位
    if raw.isdigit() and len(raw) <= 2:
        suffix = raw.zfill(2)
        hits = [s for s in roster if s["sid"].endswith(suffix)]
        if len(hits) == 1:
            return hits[0], [], None
        if hits:
            return None, hits, None

    # 3) 姓名包含关系（只写了名的一部分）
    fuzzy = [s for s in roster if raw in s["name"]]
    if len(fuzzy) == 1:
        return fuzzy[0], [], {"fuzzy": True, "input": key, "rate": 1.0}
    if len(fuzzy) > 1:
        return None, fuzzy, None

    # 4) 近似匹配（拍照认错字、手写看不清）
    scored = sorted(((name_similarity(raw, s["name"]), s) for s in roster),
                    key=lambda pair: pair[0], reverse=True)
    top = [pair for pair in scored if pair[0] >= SUGGEST_RATE][:5]
    if not top:
        return None, [], None

    best_rate, best = top[0]
    second_rate = top[1][0] if len(top) > 1 else 0.0
    if best_rate >= AUTO_MATCH_RATE and (best_rate - second_rate) >= AUTO_MATCH_GAP:
        return best, [], {"fuzzy": True, "input": key, "rate": round(best_rate, 3)}
    return None, [s for _, s in top], None


def guess_target(node):
    """从参数字典里取目标周期，缺省用软件当前周期。"""
    state = read_app_state()
    def pick(key):
        value = node.get(key)
        if value in (None, ""):
            return state.get(key)
        try:
            return int(value)
        except (TypeError, ValueError):
            return state.get(key)
    return pick("year"), pick("month"), pick("week")


# ────────────────────────────── 工具实现 ──────────────────────────────
def text_result(text, is_error=False):
    return {"content": [{"type": "text", "text": text}], "isError": bool(is_error)}


def tool_get_context(args):
    state = read_app_state()
    roster = load_roster()
    lines = [
        f"软件正在运行：{'是' if state.get('running') else '否'}",
        f"当前考评周期：{state.get('year', '?')}年{state.get('month', '?')}月 · 第{state.get('week', '?')}周",
        f"班级：{load_class_name()}",
        f"学生人数：{len(roster)}",
        f"数据目录：{BASE_DIR}",
    ]
    if not state.get("running"):
        lines.append("")
        lines.append("提示：软件没开也没关系，记录会先排队；等你打开软件并切到对应周次，就会自动填入。")
    return text_result("\n".join(lines))


def tool_list_students(args):
    roster = load_roster()
    lines = ["学号 | 姓名 | 寝室 | 角色"]
    for s in roster:
        lines.append(f"{s['sid']} | {s['name']} | {s['dorm']} | {s['role']}")
    lines.append("")
    lines.append(f"共 {len(roster)} 人。add_records 的 student 字段可以填姓名、完整学号或学号后两位。")
    return text_result("\n".join(lines))


def tool_get_week_records(args):
    year, month, week = guess_target(args)
    if None in (year, month, week):
        return text_result("无法确定要查询哪一周，请显式传 year / month / week。", True)
    roster = load_roster()
    records, exists = load_week_records(year, month, week, roster)
    if not exists:
        return text_result(f"{year}年{month}月的周纪实表还不存在（还没生成过）。", True)

    lines = [f"{year}年{month}月 第{week}周 已录入记录："]
    total = 0
    for s in roster:
        items = records.get(s["sid"]) or []
        if not items:
            continue
        total += len(items)
        detail = "；".join(f"{r['reason']} {r['score']:g}" for r in items)
        lines.append(f"  {s['name']}({s['sid'][-2:]})：{detail}")
    if total == 0:
        lines.append("  （本周还没有任何记录）")
    lines.append("")
    lines.append(f"共 {total} 条。录入前先看这里，避免重复。")
    return text_result("\n".join(lines))


def tool_get_score_rules(args):
    rows = load_rules()
    if not rows:
        return text_result("名单模板里没有「加分预览」工作表。", True)
    lines = ["加分标准（来自 名单模板.xlsx → 加分预览）："]
    for row in rows:
        cells = [c for c in row]
        while cells and not cells[-1]:
            cells.pop()
        if cells:
            lines.append("  " + " | ".join(cells))
    lines.append("")
    lines.append("提示：add_records 的 score 可以省略 —— 只要 reason 里能看出奖项等级"
                 "（如「省赛一等奖」），服务端会自动按这张表取分，写记录时不用先调这个工具。")
    return text_result("\n".join(lines))


def tool_get_month_summary(args):
    year = args.get("year")
    month = args.get("month")
    state = read_app_state()
    year = int(year) if year not in (None, "") else state.get("year")
    month = int(month) if month not in (None, "") else state.get("month")
    if None in (year, month):
        return text_result("无法确定月份，请显式传 year / month。", True)

    roster = load_roster()
    summary, weeks = load_month_summary(year, month, roster)
    base = float(args.get("base_score") or 0)
    lines = [f"{year}年{month}月 月度汇总（读到 {weeks} 个周次，底分 {base:g}）："]
    for s in roster:
        total = base + summary.get(s["sid"], 0.0)
        lines.append(f"  {s['name']}({s['sid'][-2:]})：本月新增 {summary.get(s['sid'], 0.0):g}"
                     f"  最终 {total:g}")
    return text_result("\n".join(lines))


def tool_add_records(args):
    records = args.get("records")
    if not isinstance(records, list) or not records:
        return text_result("records 必须是非空数组。", True)

    year, month, week = guess_target(args)
    if None in (year, month, week):
        return text_result(
            "无法确定要写入哪一周。请显式传 year / month / week，"
            "或先打开软件让它把自己的当前周期告诉我。", True)

    roster = load_roster()
    resolved = []
    problems = []

    for index, raw in enumerate(records, 1):
        if not isinstance(raw, dict):
            problems.append(f"第 {index} 条不是对象")
            continue
        student_key = raw.get("student") or raw.get("name") or raw.get("sid")
        student, candidates, match = resolve_student(roster, student_key)
        if student is None:
            if len(candidates) == 1:
                only = candidates[0]
                problems.append(
                    f"第 {index} 条「{student_key}」在名单里找不到，最接近的是 "
                    f"{only['name']}({only['sid'][-2:]})，请确认姓名")
            elif candidates:
                names = "、".join(f"{c['name']}({c['sid'][-2:]})" for c in candidates[:6])
                problems.append(f"第 {index} 条「{student_key}」有多个相近的人：{names}，请写成全名")
            else:
                problems.append(f"第 {index} 条「{student_key}」在名单里找不到")
            continue
        reason = clean(raw.get("reason")) or "AI 记录"
        auto = None
        raw_score = raw.get("score")
        if raw_score is None or str(raw_score).strip() == "":
            # 没给分数：自己去模板「加分预览」里按奖项名查
            value, source = lookup_score(reason)
            if value is None:
                problems.append(
                    f"第 {index} 条（{student['name']}）没给 score，"
                    f"「{reason}」在模板加分预览里也没找到对应标准，请补上分值")
                continue
            score = value
            auto = {"from": source, "score": value}
        else:
            try:
                score = float(re.sub(r"[^\d.\-+]", "", str(raw_score)))
            except (TypeError, ValueError):
                problems.append(f"第 {index} 条（{student['name']}）score 不是数字")
                continue
            if score != score or score in (float("inf"), float("-inf")):
                problems.append(f"第 {index} 条（{student['name']}）score 不是有限数字")
                continue
        item = {
            "sid": student["sid"],
            "name": student["name"],
            "reason": reason,
            "score": score,
        }
        if auto:
            # 分数是程序按模板算的，让软件在界面上注明，方便人工核对
            item["auto_score"] = auto
        if match:
            # 让软件把这条标成「近似匹配」，界面上用不同颜色提示人工核对
            item["match"] = match
        resolved.append(item)

    if problems:
        # 有一条对不上就整批不写，避免写一半更难收拾
        lines = ["以下记录没能解析，整批都没有写入："]
        lines += [f"  - {p}" for p in problems]
        lines.append("")
        lines.append("请修正后重新调用 add_records。可以先用 list_students 核对姓名。")
        return text_result("\n".join(lines), True)

    batch_id = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    replace = bool(args.get("replace"))
    batch_key = clean(args.get("batch_key"))
    batch = {
        "id": batch_id,
        "year": int(year),
        "month": int(month),
        "week": int(week),
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "source": "mcp",
        "note": clean(args.get("note")),
        "records": resolved,
    }
    if replace:
        batch["replace"] = True
    if batch_key:
        batch["batch_key"] = batch_key
    write_json_atomic(os.path.join(INBOX_DIR, f"{batch_id}.json"), batch)

    state = read_app_state()
    same_cycle = (str(state.get("year")) == str(year)
                  and str(state.get("month")) == str(month)
                  and str(state.get("week")) == str(week))
    lines = [f"已排入导入队列：{year}年{month}月 第{week}周，共 {len(resolved)} 条"]
    fuzzy_items = []
    auto_items = []
    for item in resolved:
        line = f"  {item['name']}({item['sid'][-2:]}) | {item['reason']} | {item['score']:g}"
        auto = item.get("auto_score")
        if auto:
            line += f"   [模板取分：{auto['from']}]"
            auto_items.append(item)
        match = item.get("match")
        if match:
            line += f"   [近似匹配：原文「{match['input']}」，相似度 {match['rate']:.0%}]"
            fuzzy_items.append(item)
        lines.append(line)
    lines.append("")
    if auto_items:
        names = "、".join(f"{i['name']}（{i['auto_score']['from']} {i['score']:g}分）"
                          for i in auto_items)
        lines.append(f"说明：有 {len(auto_items)} 条没给分数，已按模板「加分预览」自动取分 —— {names}。")
        lines.append("如果取分不对，请让用户到软件「AI 连接」页改分值，或直接用明确的 score 重写。")
    if fuzzy_items:
        names = "、".join(f"{i['name']}（原文 {i['match']['input']}）" for i in fuzzy_items)
        lines.append(f"注意：有 {len(fuzzy_items)} 条是近似匹配 —— {names}。")
        lines.append("软件「AI 连接」页会把它们标成橙色，请提醒用户核对一下名字。")
    if not state.get("running"):
        lines.append("软件当前没有运行，记录已排队；打开软件并切到这一周后会自动填入。")
    elif same_cycle:
        lines.append("软件界面正好停在这一周，大约 2 秒内会自动填入并高亮。")
    else:
        lines.append(f"注意：软件界面现在停在 {state.get('year')}年{state.get('month')}月"
                     f"第{state.get('week')}周，切换到目标周次后才会填入。")
    return text_result("\n".join(lines))


def tool_set_dorm_stars(args):
    """把卫生检查结果直接发到软件「寝室统计」页的星级滑块上（不写记录）。

    用户核对后自己点「应用打分」才会真正加减分。
    """
    raw = args.get("stars")
    if not isinstance(raw, dict) or not raw:
        return text_result("stars 需要是 {寝室号: 星级} 的对象，例如 "
                           '{"柳412": "五星", "梅412": "三星"}', True)

    # 模板里真实存在的寝室号（走名单缓存，未知=没分到寝室的不算）
    rooms = {it["dorm"] for it in load_roster() if it.get("dorm") != "未知"}

    star_map = dict(STAR_TEXT_MAP)
    star_map.update({"三": 3, "四": 4, "五": 5})

    stars, problems = {}, []
    for dorm, star in raw.items():
        d = clean(dorm)
        key = clean(star)
        value = star_map.get(key) or (int(key) if key in ("3", "4", "5") else None)
        if value is None:
            problems.append(f"{d or '（空寝室号）'}：星级「{star}」不认识（可用 五星/四星/三星 或 5/4/3）")
            continue
        if not d:
            problems.append("寝室号为空")
            continue
        matched = [r for r in rooms if r == d]
        if not matched:
            problems.append(f"{d}：名单里没有这个寝室（模板里是：{'、'.join(sorted(rooms))}）")
            continue
        stars[d] = value

    # 部分识别也照常写入有效的（未识别的留给用户补报，软件里挂红色「未发」标）；
    # 一个有效的都没有才报错。
    if not stars:
        lines = ["一个寝室都没识别出来，没有写入。"]
        if problems:
            lines += [f"  - {x}" for x in problems]
        lines.append("")
        lines.append("寝室号是「园区首字+房号」，如 柳412 / 梅420。"
                     "学校 App 里的「柳园420」对应 柳420。")
        return text_result("\n".join(lines), True)

    state = read_app_state()
    year = int(args.get("year") or state.get("year") or datetime.datetime.now().year)
    month = int(args.get("month") or state.get("month") or datetime.datetime.now().month)
    week = int(args.get("week") or state.get("week") or 1)
    batch_id = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    write_json_atomic(os.path.join(INBOX_DIR, f"{batch_id}.json"), {
        "id": batch_id,
        "year": year, "month": month, "week": week,
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "source": "mcp",
        "type": "dorm_stars",
        "stars": stars,
    })

    same_cycle = (str(state.get("year")) == str(year)
                  and str(state.get("month")) == str(month)
                  and str(state.get("week")) == str(week))
    lines = [f"已把 {len(stars)} 个寝室的星级发给软件："]
    lines += [f"  {d} {'三四五'[v - 3]}星" for d, v in sorted(stars.items())]
    if problems:
        lines.append(f"未识别 {len(problems)} 个（没有写入，软件里会保持红色「未发」标）：")
        lines += [f"  - {x}" for x in problems]
    lines.append("")
    lines.append("这些星级直接设在「寝室统计」页的滑块上（寝室号显示为青色），"
                 "用户核对后点「应用打分」才会真正加减分并生成表格。")
    if not state.get("running"):
        lines.append("软件当前没有运行，已排队；打开软件并切到对应周次后会自动设上。")
    elif not same_cycle:
        lines.append(f"注意：软件界面现在停在 {state.get('year')}年{state.get('month')}月"
                     f"第{state.get('week')}周，切换到目标周次后才会设上。")
    return text_result("\n".join(lines))


def tool_undo_last(args):
    """请软件撤回最近的导入（只对软件当前周期有效）。"""
    year, month, week = guess_target(args)
    if None in (year, month, week):
        return text_result("无法确定要继续撤回哪一周，请显式传 year / month / week。", True)
    try:
        count = int(args.get("count") or 1)
    except (TypeError, ValueError):
        count = 1
    count = max(1, min(count, 5))

    state = read_app_state()
    same_cycle = (str(state.get("year")) == str(year)
                  and str(state.get("month")) == str(month)
                  and str(state.get("week")) == str(week))
    if not same_cycle:
        return text_result(
            f"撤回只能对软件当前周期生效。软件现在停在 "
            f"{state.get('year')}年{state.get('month')}月第{state.get('week')}周，"
            f"而你要撤的是 {year}年{month}月第{week}周。\n"
            f"请让用户先把软件切到那一周，再重新调用。", True)

    batch_id = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    batch = {
        "id": batch_id,
        "type": "undo",
        "count": count,
        "year": int(year),
        "month": int(month),
        "week": int(week),
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "source": "mcp",
    }
    write_json_atomic(os.path.join(INBOX_DIR, f"{batch_id}.json"), batch)
    return text_result(
        f"已请求撤回最近 {count} 步（{year}年{month}月第{week}周）。软件约 2 秒内处理。\n"
        "如果软件提示撤回栈是空的，说明中途切过周期——撤回快照只在同一周期内有效。")


TOOLS = [
    {
        "name": "add_records",
        "description": "【写记录的主入口，一次调用即可完成，不需要先调任何其他工具】"
                       "把加减分记录排入导入队列，软件约 2 秒后自动填入并高亮。"
                       "year/month/week 可以完全不传，服务端会自动用软件当前考评周期。"
                       "student 支持姓名、完整学号、学号后两位；"
                       "姓名允许模糊（拍照或手写认错个别字也没关系），服务端会自动找最像的人，"
                       "并在返回里标出哪几条是近似匹配。"
                       "score 也可以不传：只要 reason 里能看出奖项（如「省赛一等奖」「国赛二等奖」"
                       "「五星寝室」「早自习迟到」「活动类比赛二等」「标兵」），"
                       "服务端会自己按模板「加分预览」表取分，并在返回里说明取了多少。"
                       "只有在「同时有多个差不多像的人」或「完全查不到分值」时才会报错，"
                       "那时补全姓名或直接给 score 重传即可。"
                       "只要有一条彻底对不上，整批都不写入。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "records": {
                    "type": "array",
                    "description": "要写入的记录列表，一次最多 50 条",
                    "items": {
                        "type": "object",
                        "properties": {
                            "student": {"type": "string",
                                        "description": "姓名（可模糊）/ 完整学号 / 学号后两位"},
                            "reason": {"type": "string",
                                       "description": "加减分原因，会写进 Excel 的「项目明细」列。"
                                                      "奖项名写清楚（如「省赛一等奖」）就能自动取分"},
                            "score": {"type": "number",
                                      "description": "分值：正数加分，负数扣分。"
                                                     "省略时按 reason 里的奖项名去模板自动查分"},
                        },
                        "required": ["student"],
                    },
                },
                "year": {"type": "integer", "description": "目标年份，省略=软件当前周期"},
                "month": {"type": "integer", "description": "目标月份，省略=软件当前周期"},
                "week": {"type": "integer", "description": "目标周次，省略=软件当前周期"},
                "note": {"type": "string", "description": "备注，便于日后追溯"},
                "replace": {"type": "boolean",
                            "description": "true = 同一个人同一明细的旧 AI 记录先删掉再写"
                                           "（重发时用这个，避免越改越多）；默认 false 追加"},
                "batch_key": {"type": "string",
                              "description": "幂等键。重试同一批时带相同的值，"
                                             "软件不会重复写入"},
            },
            "required": ["records"],
        },
    },
    {
        "name": "set_dorm_stars",
        "description": "【寝室卫生检查结果专用】把一批寝室的星级直接设到软件「寝室统计」页的"
                       "滑块上（不写加减分记录），用户核对后自己点「应用打分」。"
                       "收到一批寝室检查照片/结果时优先用这个，不要用 add_records 逐人写寝室分。"
                       "寝室号是「园区首字+房号」：学校 App 里的「柳园420」对应 柳420。"
                       "星级可用 五星/四星/三星 或 5/4/3。没报结果的寝室不要写，"
                       "软件会标红色「未发」提醒。year/month/week 省略=软件当前周期。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "stars": {
                    "type": "object",
                    "description": "寝室号 → 星级，如 {\"柳412\": \"五星\", \"梅412\": \"三星\"}",
                    "additionalProperties": {"type": "string"},
                },
                "year": {"type": "integer", "description": "目标年份，省略=软件当前周期"},
                "month": {"type": "integer", "description": "目标月份，省略=软件当前周期"},
                "week": {"type": "integer", "description": "目标周次，省略=软件当前周期"},
            },
            "required": ["stars"],
        },
    },
    {
        "name": "get_context",
        "description": "查看软件状态：是否在运行、当前考评周期（年/月/周）、班级、人数。"
                       "只有用户问「现在第几周」「软件开没开」时才需要调；"
                       "调用 add_records 写记录时不用先调它（会自动用当前周期）。",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_week_records",
        "description": "读某一周已录入的记录，用于查重或核对「记上了吗」。"
                       "不传参数则用软件当前周期。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "year": {"type": "integer", "description": "年份，如 2026"},
                "month": {"type": "integer", "description": "月份 1-12"},
                "week": {"type": "integer", "description": "周次 1-5"},
            },
        },
    },
    {
        "name": "get_score_rules",
        "description": "读加分标准表（竞赛等级分值、寝室星级加减分、早自习扣分等）。"
                       "用户问「这个能加多少分」时才调；写记录时不用先调 —— "
                       "add_records 在 score 省略时会自动按这张表取分。",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_month_summary",
        "description": "汇总某月各周加减分并算最终分。用户问「现在多少分」「谁最低」时调。"
                       "不传年月则用软件当前周期。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "year": {"type": "integer"},
                "month": {"type": "integer"},
                "base_score": {"type": "number", "description": "每人底分，默认 0"},
            },
        },
    },
    {
        "name": "list_students",
        "description": "列出全班 47 人（学号/姓名/寝室/角色）。"
                       "只在「需要让用户从几个相近的名字里挑一个」时才调；"
                       "add_records 已支持姓名模糊匹配，一般不用先查名单。",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "undo_last",
        "description": "【写错了用它】撤回最近一次 AI 导入，可一次撤多步（最多 5）。"
                       "用户说「刚才那批错了」「删掉吧」「撤回」时调这个，"
                       "不要用写负分的方式抵消。"
                       "只能撤软件当前周期：撤回快照不跨周期，切过周就撤不回去了。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "count": {"type": "integer", "description": "撤回几步，默认 1，最多 5"},
                "year": {"type": "integer", "description": "哪一年，省略=软件当前周期"},
                "month": {"type": "integer", "description": "哪个月，省略=软件当前周期"},
                "week": {"type": "integer", "description": "第几周，省略=软件当前周期"},
            },
        },
    },
]

HANDLERS = {
    "get_context": tool_get_context,
    "list_students": tool_list_students,
    "get_week_records": tool_get_week_records,
    "get_score_rules": tool_get_score_rules,
    "get_month_summary": tool_get_month_summary,
    "add_records": tool_add_records,
    "set_dorm_stars": tool_set_dorm_stars,
    "undo_last": tool_undo_last,
}


# ────────────────────────────── 协议实现 ──────────────────────────────
def send(message):
    sys.stdout.write(json.dumps(message, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def reply_ok(msg_id, result):
    send({"jsonrpc": "2.0", "id": msg_id, "result": result})


def reply_error(msg_id, code, message, data=None):
    error = {"code": code, "message": message}
    if data is not None:
        error["data"] = data
    send({"jsonrpc": "2.0", "id": msg_id, "error": error})


def handle_initialize(msg_id, params):
    requested = (params or {}).get("protocolVersion") or PROTOCOL_VERSION
    version = requested if requested in SUPPORTED_VERSIONS else PROTOCOL_VERSION
    reply_ok(msg_id, {
        "protocolVersion": version,
        "capabilities": {"tools": {"listChanged": False}},
        "serverInfo": SERVER_INFO,
        "instructions": (
            "这是「德育分自动化系统」的写入工具。"
            "用户用自然语言描述加减分时，先用 get_context 和 list_students 核对，"
            "再调用 add_records 把记录排入导入队列；软件会自动填入并高亮。"
        ),
    })


def handle_tools_call(msg_id, params):
    name = (params or {}).get("name")
    arguments = (params or {}).get("arguments") or {}
    handler = HANDLERS.get(name)
    if handler is None:
        # 用 isError 结果而不是 JSON-RPC 错误：AI 能看到可读原因，多数客户端也更好处理
        reply_ok(msg_id, text_result(
            f"没有名为「{name}」的工具。可用工具：{'、'.join(HANDLERS)}", True))
        return
    try:
        reply_ok(msg_id, handler(arguments))
    except Exception as exc:
        log(f"工具 {name} 执行失败：{exc}\n{traceback.format_exc()}")
        reply_ok(msg_id, text_result(f"工具执行出错：{exc}", True))


def serve():
    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            message = json.loads(line)
        except json.JSONDecodeError:
            send({"jsonrpc": "2.0", "id": None,
                  "error": {"code": -32700, "message": "Parse error"}})
            continue

        method = message.get("method")
        msg_id = message.get("id")
        params = message.get("params") or {}

        if msg_id is None:
            # 通知类消息：不回应
            if method == "notifications/initialized":
                log("客户端初始化完成")
            continue

        try:
            if method == "initialize":
                handle_initialize(msg_id, params)
            elif method == "ping":
                reply_ok(msg_id, {})
            elif method == "tools/list":
                reply_ok(msg_id, {"tools": TOOLS})
            elif method == "tools/call":
                handle_tools_call(msg_id, params)
            elif method == "resources/list":
                reply_ok(msg_id, {"resources": []})
            elif method == "prompts/list":
                reply_ok(msg_id, {"prompts": []})
            elif method in ("logging/setLevel", "notifications/cancelled"):
                reply_ok(msg_id, {})
            else:
                reply_error(msg_id, -32601, f"Method not found: {method}")
        except Exception as exc:
            log(f"处理 {method} 失败：{exc}\n{traceback.format_exc()}")
            reply_error(msg_id, -32603, "Internal error", str(exc))


def selftest():
    """命令行自检：不开协议，直接跑工具函数，方便排查。"""
    print(f"程序目录: {BASE_DIR}")
    print(f"名单模板: {TEMPLATE}  存在={os.path.exists(TEMPLATE)}")
    state = read_app_state()
    print(f"软件状态: 运行中={state.get('running')}  周期={state.get('year')}-{state.get('month')}-{state.get('week')}")
    roster = load_roster()
    print(f"学生人数: {len(roster)}")
    if roster:
        print(f"  示例: {roster[0]}")
    print()
    for name in ("get_context", "add_records_none"):
        pass
    print("=== get_context ===")
    print(tool_get_context({})["content"][0]["text"])
    print("\n=== list_students（前 5 行）===")
    print("\n".join(tool_list_students({})["content"][0]["text"].splitlines()[:6]))
    print("\n=== get_score_rules（前 8 行）===")
    print("\n".join(tool_get_score_rules({})["content"][0]["text"].splitlines()[:8]))
    print("\n=== 自动取分（不给 score，只给奖项名）===")
    for text in ("省赛一等奖", "国赛二等奖", "市赛三等", "院级一等奖",
                 "五星寝室", "三星寝室", "早自习迟到", "早自习缺勤",
                 "活动类比赛二等", "标兵", "随便加的"):
        value, source = lookup_score(text)
        shown = "查不到" if value is None else f"{value:g} 分（{source}）"
        print(f"  {text:<12} → {shown}")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stdin.reconfigure(encoding="utf-8")
        except Exception:
            pass
        log(f"MCP 服务启动，程序目录 {BASE_DIR}")
        serve()
