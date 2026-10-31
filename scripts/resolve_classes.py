# -*- coding: utf-8 -*-
"""三级回退查表：十部(BH) -> HW-TEMU -> SO-SHEIN。

三张表的款号前缀不同（TM-BH- / TM-HW- / SHEIN-SO-），必须先去掉前缀，
按「款号-段-尺码」归一后再匹配，否则 0 命中。

规则：每个分类字段独立回退——优先十部表，该字段缺失才依次用 HW、SO。
"""
import argparse
import json
import os
import sys

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("需要 openpyxl：pip install openpyxl")

PREFIXES = ("TM-BH-", "TM-HW-", "SHEIN-SO-")
FIELDS = ("大类", "中类", "小类")

# 默认文件名（可按需覆盖）
DEFAULT_FILES = {
    "十部-BH": "十部半托管-库存管理表",
    "HW-TEMU": "HW品牌店铺-TEMU库存管理表",
    "SO-SHEIN": "SO品牌店铺-SHEIN库存管理表",
}
SEARCH_DIRS = [
    os.path.expanduser("~/Desktop"),
    os.path.expanduser("~/Downloads"),
    os.path.expanduser("~/Documents"),
    os.path.expanduser("~/OneDrive/Desktop"),
    os.getcwd(),
]
# 允许用环境变量指定目录（多个用 ; 分隔），避免源表换位置后失效
if os.environ.get("KDOCS_APPROVAL_DIRS"):
    SEARCH_DIRS = [p for p in os.environ["KDOCS_APPROVAL_DIRS"].split(";") if p] + SEARCH_DIRS


def norm(msku):
    """去掉品牌前缀，得到跨表统一键，如 N103507-3B-S。"""
    s = str(msku).strip()
    for p in PREFIXES:
        if s.startswith(p):
            return s[len(p):]
    return s


def find_workbook(tag, explicit=None):
    """定位库存表文件。explicit 优先，否则在搜索目录里按文件名前缀找（取最新）。"""
    if explicit:
        if not os.path.exists(explicit):
            raise FileNotFoundError(explicit)
        return explicit
    stem = DEFAULT_FILES[tag]
    candidates = []
    for d in SEARCH_DIRS:
        if not os.path.isdir(d):
            continue
        for f in os.listdir(d):
            if f.startswith(stem) and f.endswith(".xlsx") and not f.startswith("~$"):
                candidates.append(os.path.join(d, f))
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def load_book(path):
    """读取 总库存表，返回 {归一MSKU: {大类,中类,小类,款号}}。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    if "总库存表" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"{path} 缺少「总库存表」sheet")
    ws = wb["总库存表"]
    d = {}
    for r in ws.iter_rows(min_row=2, max_col=7, values_only=True):
        if not r[6]:
            continue
        k = norm(r[6])
        if k in d:
            continue
        d[k] = {
            "大类": (str(r[0]).strip() if r[0] is not None else None),
            "中类": (str(r[1]).strip() if r[1] is not None else None),
            "小类": (str(r[2]).strip() if r[2] is not None else None),
            "款号": (str(r[5]).strip() if r[5] is not None else None),
        }
    wb.close()
    return d


class Resolver:
    def __init__(self, paths=None):
        paths = paths or {}
        self.books = {}
        self.missing = []
        for tag in DEFAULT_FILES:
            p = find_workbook(tag, paths.get(tag))
            if p and os.path.exists(p):
                self.books[tag] = load_book(p)
            else:
                self.missing.append(tag)

    def resolve(self, msku):
        """返回 (分类dict, 每字段来源dict)。"""
        key = norm(msku)
        out = {f: None for f in FIELDS}
        src = {f: None for f in FIELDS}
        for tag in DEFAULT_FILES:  # 严格按 BH -> HW -> SO
            rec = self.books.get(tag, {}).get(key)
            if not rec:
                continue
            for f in FIELDS:
                if out[f] is None and rec.get(f):
                    out[f] = rec[f]
                    src[f] = tag
        return out, src


def main():
    ap = argparse.ArgumentParser(description="三级回退查分类")
    ap.add_argument("--msku", action="append", default=[], help="单个 MSKU，可重复")
    ap.add_argument("--file", help="每行一个 MSKU 的文件")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    for tag in DEFAULT_FILES:
        ap.add_argument(f"--{tag}", dest=tag, help=f"{tag} 库存表路径")
    a = ap.parse_args()

    keys = list(a.msku)
    if a.file:
        with open(a.file, encoding="utf-8") as f:
            keys += [ln.strip() for ln in f if ln.strip()]
    if not keys:
        ap.error("需要 --msku 或 --file")

    paths = {tag: getattr(a, tag) for tag in DEFAULT_FILES if getattr(a, tag)}
    R = Resolver(paths)
    if R.missing:
        print(f"[warn] 未找到库存表: {', '.join(R.missing)}", file=sys.stderr)

    results = []
    for k in keys:
        cls, src = R.resolve(k)
        results.append({"msku": k, **cls, "source": src})
        if not a.json:
            print(f"{k}\t{cls['大类']} | {cls['中类']} | {cls['小类']}")
            print(f"\t来源: 大类={src['大类']} 中类={src['中类']} 小类={src['小类']}")
    if a.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
