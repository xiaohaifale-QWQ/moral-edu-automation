# -*- coding: utf-8 -*-
"""名单导入写表逻辑单测（不启动 GUI）。

用 MoralScoreApp.__new__ 绕过 __init__，只测三个写表函数：
  _write_month_roster / _write_week_roster / _write_dorm_roster
在空白模板副本上写入 → 用 openpyxl 断言结果。

跑法：python tools/test_roster_write.py
"""

import os
import shutil
import sys
import tempfile

import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO_ROOT, "软件说明"))

BLANK = os.path.join(REPO_ROOT, "_build_pkg", "payload", "名单模板.xlsx")

import importlib.util
spec = importlib.util.spec_from_file_location(
    "moralapp", os.path.join(REPO_ROOT, "软件说明", "德育分自动化系统.py"))
mod = importlib.util.module_from_spec(spec)
sys.modules["moralapp"] = mod
spec.loader.exec_module(mod)

FAILS = []


def check(label, ok, extra=""):
    print(f"  [{'OK ' if ok else 'FAIL'}] {label}" + (f"  → {extra}" if extra else ""))
    if not ok:
        FAILS.append(label)


def make_rows(n_dorms=2):
    rows = []
    for i in range(5):
        rows.append({"sid": f"20250117{i:03d}",
                     "name": f"测试{i}号",
                     "dorm": f"44{1 + i % n_dorms}",
                     "role": "室长" if i < 2 else "成员"})
    return rows


def main():
    workdir = tempfile.mkdtemp(prefix="roster-test-")
    tpl = os.path.join(workdir, "名单模板.xlsx")
    shutil.copy2(BLANK, tpl)

    app = mod.MoralScoreApp.__new__(mod.MoralScoreApp)   # 不跑 __init__
    rows = make_rows()
    wb = openpyxl.load_workbook(tpl)
    app._write_month_roster(wb["月表模板"], rows, "集成电路技术999")
    app._write_week_roster(wb["周表模板"], rows, "集成电路技术999班")
    app._write_dorm_roster(wb["寝室名单"], rows)
    wb.save(tpl)
    wb.close()

    wb = openpyxl.load_workbook(tpl)

    # 月表：5 人依次落位，班级写 C 列，第 7 行起清空
    ws = wb["月表模板"]
    got = [(ws.cell(2 + i, 1).value, ws.cell(2 + i, 2).value, ws.cell(2 + i, 3).value)
           for i in range(5)]
    check("月表 5 人学号/姓名/班级", got == [
        ("20250117000", "测试0号", "集成电路技术999"),
        ("20250117001", "测试1号", "集成电路技术999"),
        ("20250117002", "测试2号", "集成电路技术999"),
        ("20250117003", "测试3号", "集成电路技术999"),
        ("20250117004", "测试4号", "集成电路技术999")], str(got[:2]))
    check("月表第 7 行起已清空",
          all(ws.cell(r, c).value is None for r in range(7, ws.max_row + 1)
              for c in (1, 2, 3, 5)))
    check("月表 D 列占位保留", str(ws.cell(2, 4).value) == "#/#/#")

    # 周表：A1 带班级；5 人 → 左 3 右 2；第 6 行起左右都空
    ws = wb["周表模板"]
    check("周表 A1", ws["A1"].value == "集成电路技术999班#月#周记实", str(ws["A1"].value))
    l_ids = [ws.cell(4 + k, 1).value for k in range(3)]
    r_ids = [ws.cell(4 + k, 5).value for k in range(2)]
    check("周表左栏 3 人", l_ids == ["20250117000", "20250117001", "20250117002"], str(l_ids))
    check("周表右栏 2 人", r_ids == ["20250117003", "20250117004"], str(r_ids))
    check("周表右栏第 3 对为空", ws.cell(6, 5).value is None and ws.cell(6, 6).value is None)

    # 寝室：441=测试0/测试1（室长 测试0）… 注意 441 只有 i=0,2,4 → 3 人
    # i%2: i=0→441 i=1→442 i=2→441 i=3→442 i=4→441
    ws = wb["寝室名单"]
    check("寝室 441 房号", ws.cell(3, 1).value == 441, str(ws.cell(3, 1).value))
    check("寝室 441 成员", ws.cell(3, 3).value == "测试0号 测试2号 测试4号",
          str(ws.cell(3, 3).value))
    check("寝室 441 室长", ws.cell(3, 7).value == "测试0号", str(ws.cell(3, 7).value))
    check("寝室 442 房号", ws.cell(5, 1).value == 442)
    check("寝室 442 成员", ws.cell(5, 3).value == "测试1号 测试3号", str(ws.cell(5, 3).value))
    check("寝室 442 之后为空", ws.cell(7, 1).value is None and ws.cell(7, 3).value is None)
    check("合并单元格保留", len(ws.merged_cells.ranges) == 39,
          str(len(ws.merged_cells.ranges)))
    wb.close()

    # ---- 扩展用例：14 个寝室 → 必须补 2 对行 + 合并 ----
    tpl2 = os.path.join(workdir, "名单模板2.xlsx")
    shutil.copy2(BLANK, tpl2)
    rows2 = []
    for i in range(14):
        for j in range(3):
            rows2.append({"sid": f"20250117{i:02d}{j}",
                          "name": f"学{i}-{j}",
                          "dorm": f"{5 + i:03d}",
                          "role": "室长" if j == 0 else "成员"})
    wb = openpyxl.load_workbook(tpl2)
    app._write_dorm_roster(wb["寝室名单"], rows2)
    app._write_month_roster(wb["月表模板"], rows2, "扩容班")
    app._write_week_roster(wb["周表模板"], rows2, "扩容班")
    wb.save(tpl2)
    wb.close()
    wb = openpyxl.load_workbook(tpl2)
    ws = wb["寝室名单"]
    check("14 寝补行后 29 行房号正确", ws.cell(3 + 2 * 13, 1).value == 18,
          str(ws.cell(3 + 2 * 13, 1).value))
    check("14 寝合并数 = 39 + 3*2", len(ws.merged_cells.ranges) == 45,
          str(len(ws.merged_cells.ranges)))
    ws = wb["月表模板"]
    check("42 人月表第 43 行有学号", ws.cell(2 + 41, 1).value == "20250117132",
          str(ws.cell(2 + 41, 1).value))
    ws = wb["周表模板"]
    check("42 人周表第 21 对（行 24）左栏有值", ws.cell(24, 1).value == "20250117062",
          str(ws.cell(24, 1).value))
    check("42 人周表第 22 对（行 25）右栏为空", ws.cell(25, 5).value is None)
    wb.close()

    shutil.rmtree(workdir, ignore_errors=True)
    if FAILS:
        print(f"\n[失败] {len(FAILS)} 项")
        return 1
    print("\n[通过] 名单写表逻辑全部正确。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
