# -*- coding: utf-8 -*-
"""把本地 xlsx 的某个 sheet 写入 Kdocs 共享表（经 kdocs-cli）。

安全设计（重要）：
- 写入前先读目标区域现有内容；若已有内容且与待写入不同，默认**拒绝写入**，
  必须加 --force 才会覆盖。审批痕迹（如「同意」）不能被静默冲掉。
- 写入后回读比对，不一致就报错，不谎报成功。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("需要 openpyxl：pip install openpyxl")


def find_cli():
    exe = shutil.which("kdocs-cli")
    if exe:
        return exe
    guess = os.path.join(os.environ.get("LOCALAPPDATA", ""), "kdocs-cli", "kdocs-cli.exe")
    if os.path.exists(guess):
        return guess
    sys.exit("找不到 kdocs-cli。安装见 references/kdocs-cli-setup.md")


def call(cli, args, payload=None):
    """调用 kdocs-cli；payload 会写成临时 JSON 用 --file 传入。"""
    tmp = None
    try:
        if payload is not None:
            fd, tmp = tempfile.mkstemp(suffix=".json")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False)
            args = args + ["--file", tmp]
        args = args + ["--compact"]
        proc = subprocess.run([cli] + args, capture_output=True, text=True, encoding="utf-8")
        out = (proc.stdout or "").strip()
        try:
            return json.loads(out)
        except json.JSONDecodeError:
            return {"code": -1, "message": out or (proc.stderr or "").strip()}
    finally:
        if tmp and os.path.exists(tmp):
            os.unlink(tmp)


def read_sheet_xlsx(path, sheet):
    """把 xlsx 的一个 sheet 读成二维字符串数组。"""
    wb = load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"{path} 里没有 sheet「{sheet}」")
    ws = wb[sheet]
    grid = []
    for r in range(1, ws.max_row + 1):
        row = []
        for c in range(1, 11):
            v = ws.cell(row=r, column=c).value
            row.append("" if v is None else str(v))
        grid.append(row)
    wb.close()
    while grid and not any(x != "" for x in grid[-1]):
        grid.pop()
    return grid


def read_cloud(cli, url, ws_id, rows, cols):
    payload = {"url": url, "worksheet_id": ws_id,
               "range": {"rowFrom": 0, "rowTo": rows - 1, "colFrom": 0, "colTo": cols - 1}}
    res = call(cli, ["sheet", "get-range-data"], payload)
    if res.get("code") != 0:
        return None, res
    cells = ((res.get("data") or {}).get("detail") or {}).get("rangeData") or []
    grid = [["" for _ in range(cols)] for _ in range(rows)]
    for c in cells:
        rf, cf = c.get("rowFrom", 0), c.get("colFrom", 0)
        if rf < rows and cf < cols:
            txt = c.get("cellText")
            if txt in (None, ""):
                txt = c.get("originalCellValue") or ""
            grid[rf][cf] = str(txt)
    return grid, res


def main():
    ap = argparse.ArgumentParser(description="把本地 sheet 写入 Kdocs 共享表")
    ap.add_argument("--url", required=True, help="共享表链接")
    ap.add_argument("--sheet", required=True, help="云端 sheet 名，如 10月")
    ap.add_argument("--xlsx", required=True, help="本地 xlsx")
    ap.add_argument("--local-sheet", help="本地 sheet 名，默认同 --sheet")
    ap.add_argument("--force", action="store_true", help="目标已有内容时强制覆盖")
    ap.add_argument("--dry-run", action="store_true", help="只比对不写入")
    a = ap.parse_args()

    cli = find_cli()
    local_sheet = a.local_sheet or a.sheet

    info = call(cli, ["sheet", "get-sheets-info", f"url={a.url}"])
    if info.get("code") != 0:
        sys.exit(f"读取工作表列表失败: {info.get('message')}")
    sheets = ((info.get("data") or {}).get("detail") or {}).get("sheetsInfo") or []
    target = next((s for s in sheets if s.get("sheetName") == a.sheet), None)
    if not target:
        names = ", ".join(s.get("sheetName", "") for s in sheets)
        sys.exit(f"云端没有 sheet「{a.sheet}」。现有: {names}")
    ws_id = target["sheetId"]

    grid = read_sheet_xlsx(a.xlsx, local_sheet)
    rows, cols = len(grid), 10
    print(f"本地: {len(grid)} 行 x {cols} 列")

    existing, res = read_cloud(cli, a.url, ws_id, max(rows, 60), cols)
    if existing is None:
        sys.exit(f"读取云端现有内容失败: {res.get('message')}")

    existing_rows = [r for r in existing[:rows]]
    occupied = any(x != "" for r in existing_rows for x in r)
    same = all(
        (existing_rows[i][j] if i < len(existing_rows) else "") == grid[i][j]
        for i in range(rows) for j in range(cols)
    ) if occupied else False

    if occupied and same:
        print("[skip] 云端内容与待写入完全一致，无需写入")
        return
    if occupied and not same:
        diff = []
        for i in range(rows):
            for j in range(cols):
                cur = existing_rows[i][j] if i < len(existing_rows) else ""
                if cur != grid[i][j]:
                    diff.append(f"  R{i+1}C{j+1}: 云端[{cur}] -> 本地[{grid[i][j]}]")
        print(f"[warn] 云端已有内容，且与本地有 {len(diff)} 处不同：")
        for d in diff[:20]:
            print(d)
        if len(diff) > 20:
            print(f"  ... 另有 {len(diff) - 20} 处")
        if not a.force:
            sys.exit("已中止：未加 --force。请先确认是否需要覆盖（审批痕迹不可冲掉）。")
        print("[force] 继续覆盖")

    if a.dry_run:
        print("[dry-run] 不写入")
        return

    range_data = []
    for i, row in enumerate(grid):
        for j, val in enumerate(row):
            if val == "":
                continue
            range_data.append({
                "rowFrom": i, "rowTo": i, "colFrom": j, "colTo": j,
                "formula": [[val]],
            })

    payload = {"url": a.url, "worksheet_id": ws_id, "rangeData": range_data}
    wres = call(cli, ["sheet", "update-range-data"], payload)
    if wres.get("code") != 0:
        sys.exit(f"写入失败: {wres.get('message')}")
    print("[ok] 写入完成，开始回读验证")

    verify, _ = read_cloud(cli, a.url, ws_id, rows, cols)
    if verify is None:
        sys.exit("回读失败，无法验证")
    bad = [(i + 1, j + 1, grid[i][j], verify[i][j])
           for i in range(rows) for j in range(cols) if verify[i][j] != grid[i][j]]
    if bad:
        print(f"[FAIL] 回读有 {len(bad)} 处不一致：")
        for r, c, exp, got in bad[:20]:
            print(f"  R{r}C{c}: 期望[{exp}] 实际[{got}]")
        sys.exit(1)
    print(f"[ok] 回读一致，{rows} 行 x {cols} 列全部匹配")


if __name__ == "__main__":
    main()
