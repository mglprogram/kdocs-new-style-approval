# -*- coding: utf-8 -*-
"""生成「新增款号审批表」xlsx，版式与共享表一致。

spec 为 JSON 数组，每项：
  {"body":"202304", "seg":"6SetA", "sizes":["S","M","L","XL","XXL"],
   "small":"...", "大类":"...", "中类":"...", "basis":"映射说明(仅日志)"}

未提供的分类字段会用 resolve_classes.py 的三级回退自动补齐。
"""
import argparse
import datetime
import json
import sys

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("需要 openpyxl：pip install openpyxl")

sys.path.insert(0, __file__.rsplit("\\", 1)[0] if "\\" in __file__ else ".")
try:
    from resolve_classes import Resolver
except ImportError:
    Resolver = None

HEADERS = ["日期", "产品大类", "产品中类", "产品小类", "款号", "MSKU",
           "申请人", "备注", "审批", "操作人"]
WIDTHS = [11, 30, 30, 58, 19, 25, 11, 11, 9, 9]
FONT_NAME = "宋体"
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DEFAULT_SIZES = ["S", "M", "L", "XL", "XXL"]


def date_text(d=None):
    d = d or datetime.date.today()
    return f"{d.month}.{d.day}日"


def build(spec, out_path, sheet_name="新增款号审批", applicant="缪广亮",
          note="新增款号", date=None, resolver=None, base=None):
    """base 给定时在已有工作簿上加 sheet（同名先删），否则新建。"""
    if base:
        wb = load_workbook(base)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name=FONT_NAME, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS[c - 1]
    ws.row_dimensions[1].height = 22

    dt = date_text(date)
    r = 2
    first = 2
    blocks = 0
    log = []

    for gi, item in enumerate(spec):
        if gi > 0:
            r += 1  # 块间空行
        body = str(item["body"])
        seg = str(item["seg"])
        kh = f"TM-BH-{body}"
        sizes = item.get("sizes") or DEFAULT_SIZES

        big = item.get("大类")
        mid = item.get("中类")
        small = item.get("小类") or item.get("small")

        if resolver and not all([big, mid, small]):
            probe = f"{kh}-{seg}-{sizes[0]}"
            cls, src = resolver.resolve(probe)
            big = big or cls["大类"]
            mid = mid or cls["中类"]
            small = small or cls["小类"]
            log.append(f"{kh}-{seg}: 自动取分类，来源 {src}")
        if not all([big, mid, small]):
            log.append(f"[warn] {kh}-{seg} 分类不完整: 大类={big} 中类={mid} 小类={small}")

        blocks += 1
        for size in sizes:
            ws.cell(row=r, column=1, value=dt if r == first else None)
            ws.cell(row=r, column=2, value=big)
            ws.cell(row=r, column=3, value=mid)
            ws.cell(row=r, column=4, value=small)
            ws.cell(row=r, column=5, value=kh)
            ws.cell(row=r, column=6, value=f"{kh}-{seg}-{size}")
            r += 1
    last = r - 1

    for row in ws.iter_rows(min_row=1, max_row=last, max_col=10):
        for cell in row:
            cell.border = BORDER
            if cell.row == 1:
                continue
            cell.font = Font(name=FONT_NAME, size=11)
            if cell.column in (1, 5, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for rr in range(2, last + 1):
        ws.row_dimensions[rr].height = 22

    # 申请人/备注/审批/操作人 纵向合并整块
    for col, val in ((7, applicant), (8, note), (9, None), (10, None)):
        ws.merge_cells(start_row=2, start_column=col, end_row=last, end_column=col)
        cell = ws.cell(row=2, column=col, value=val)
        cell.font = Font(name=FONT_NAME, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for rr in range(2, last + 1):
        for cc in range(7, 11):
            ws.cell(row=rr, column=cc).border = BORDER

    ws.freeze_panes = "A2"
    wb.save(out_path)

    total = sum(1 for rr in range(2, last + 1) if ws.cell(row=rr, column=6).value)
    return {"blocks": blocks, "rows": total, "last_row": last, "log": log}


def main():
    ap = argparse.ArgumentParser(description="生成新增款号审批表")
    ap.add_argument("--spec", required=True, help="款号定义 JSON 文件")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    ap.add_argument("--sheet", default="新增款号审批", help="sheet 名")
    ap.add_argument("--base", help="在已有工作簿上追加 sheet")
    ap.add_argument("--applicant", default="缪广亮")
    ap.add_argument("--note", default="新增款号")
    ap.add_argument("--date", help="YYYY-MM-DD，默认今天")
    ap.add_argument("--no-lookup", action="store_true", help="不自动查分类")
    a = ap.parse_args()

    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)

    resolver = None if a.no_lookup or Resolver is None else Resolver()
    date = datetime.date.fromisoformat(a.date) if a.date else None

    info = build(spec, a.out, sheet_name=a.sheet, applicant=a.applicant,
                 note=a.note, date=date, resolver=resolver, base=a.base)
    for line in info["log"]:
        print(line)
    print(f"[ok] {a.out}  sheet={a.sheet}  款号块={info['blocks']}  明细行={info['rows']}")


if __name__ == "__main__":
    main()
