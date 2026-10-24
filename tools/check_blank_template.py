# -*- coding: utf-8 -*-
"""空白模板自检：所有断言通过才允许进入安装包（方案 2.2 节）。

用法：
    python tools/check_blank_template.py [源模板.xlsx] [空白模板.xlsx]

断言内容：
  1. 不含 11 位学号 / 真实姓名 / 班级名
  2. 每张表的合并单元格数量与源模板一致（结构没被破坏）
  3. 业务规则表（寝室加分规则 / 加分预览）与源模板逐格一致
  4. 周表模板2 与源模板逐格一致
  5. 占位符还在：周表模板 A1=#月#周记实，月表模板 D2=#/#/#
"""

import os
import re
import sys

import openpyxl

DEFAULT_SRC = r"E:\邮电\权益部分\德育分统计表\德育分自动化系统\名单模板.xlsx"
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DST = os.path.join(REPO_ROOT, "_build_pkg", "payload", "名单模板.xlsx")

# 真实姓名/班级名清单放在【不入库】的本地文件里（tools/敏感词-本地.txt，每行一个），
# 公开仓库不带任何真实学生信息；文件缺失时只跑结构断言并提示。
SENSITIVE_WORDS_FILE = os.path.join(REPO_ROOT, "tools", "敏感词-本地.txt")


def load_sensitive_words():
    try:
        with open(SENSITIVE_WORDS_FILE, encoding="utf-8") as f:
            words = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
        return words, True
    except OSError:
        return [], False


def all_values(ws):
    out = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                out.append(str(cell.value))
    return out


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    dst = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_DST
    if not os.path.exists(dst):
        print(f"[错误] 找不到空白模板：{dst}（先跑 make_blank_template.py）")
        return 1

    wb_src = openpyxl.load_workbook(src)
    wb = openpyxl.load_workbook(dst)
    fails = []

    def check(label, ok):
        print(f"  [{'OK ' if ok else 'FAIL'}] {label}")
        if not ok:
            fails.append(label)

    # 1) 敏感数据（真实姓名/班级清单来自本地敏感词文件）
    blob = "\n".join(v for ws in wb.worksheets for v in all_values(ws))
    check("无 11 位学号", not re.search(r"\d{11}", blob))
    words, have_words = load_sensitive_words()
    if have_words:
        for name in words:
            check(f"无敏感词「{name}」", name not in blob)
    else:
        print(f"  [提示] 未找到 {SENSITIVE_WORDS_FILE}，跳过真实姓名/班级断言")

    # 2) 结构回归：合并单元格数量必须与源模板一致
    check("sheet 列表一致", wb.sheetnames == wb_src.sheetnames)
    for ws_s in wb_src.worksheets:
        if ws_s.title in wb.sheetnames:
            n_s = len(ws_s.merged_cells.ranges)
            n_d = len(wb[ws_s.title].merged_cells.ranges)
            check(f"合并单元格一致 [{ws_s.title}] {n_s}=={n_d}", n_s == n_d)

    # 3) 业务规则必须保留（数值抽检 + 整表逐格一致）
    rules = wb["寝室加分规则"]
    check("寝室加分规则 B3=-1", rules["B3"].value == -1)
    check("寝室加分规则 E5=2.5", rules["E5"].value == 2.5)
    prev = wb["加分预览"]
    check("加分预览 B2=12", prev["B2"].value == 12)
    check("加分预览 H8=-0.5", prev["H8"].value == -0.5)
    for sheet in ("寝室加分规则", "加分预览"):
        vals_s = all_values(wb_src[sheet])
        vals_d = all_values(wb[sheet])
        check(f"[{sheet}] 内容与源模板逐格一致", vals_s == vals_d)
    # 周表模板2：除 D/H 签名值列外逐格一致，签名值列必须为空
    ws_s, ws_d = wb_src["周表模板2"], wb["周表模板2"]
    same, sig_empty = True, True
    for row_s, row_d in zip(ws_s.iter_rows(), ws_d.iter_rows()):
        for cs, cd in zip(row_s, row_d):
            if cs.column in (4, 8):                      # D / H 签名值列
                if cd.value is not None:
                    sig_empty = False
            elif str(cs.value) != str(cd.value):
                same = False
    check("[周表模板2] 标签与结构同源模板", same)
    check("[周表模板2] 签名值列（D/H）已清空", sig_empty)

    # 4) 占位符
    check("周表模板 A1 = #月#周记实", wb["周表模板"]["A1"].value == "#月#周记实")
    check("月表模板 D2 = #/#/#", str(wb["月表模板"]["D2"].value) == "#/#/#")

    # 5) 名单区确实是空的
    ws_m = wb["月表模板"]
    roster = [ws_m.cell(r, c).value for r in range(2, ws_m.max_row + 1) for c in (1, 2)]
    check("月表名单区已清空", all(v is None for v in roster))
    ws_d = wb["寝室名单"]
    dorms = [ws_d.cell(r, c).value for r in range(3, ws_d.max_row + 1) for c in (1, 3, 7)]
    check("寝室名单区已清空", all(v is None for v in dorms))

    wb_src.close()
    wb.close()

    if fails:
        print(f"\n[失败] {len(fails)} 项断言未通过，禁止打包：")
        for f in fails:
            print("   -", f)
        return 1
    print("\n[通过] 空白模板自检全部通过，可以进入安装包。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
