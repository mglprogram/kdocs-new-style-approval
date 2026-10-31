---
name: kdocs-new-style-approval
description: >-
  Build the 新增款号审批表 (new style-number approval sheet) that gets pasted into the shared
  WPS/Kdocs 2026拆借新增款号审批表, and optionally write it straight to the cloud sheet with
  kdocs-cli. Resolves 产品大类/中类/小类 from the 总库存表 sheet of the 十部半托管 (TM-BH),
  HW-TEMU (TM-HW) and SO-SHEIN (SHEIN-SO) inventory workbooks with a strict
  BH -> HW -> SO fallback. Use this skill whenever the user gives a list of 新增款号 /
  TM-BH-xxx style numbers to add, says "帮我去新增" / "这些款号帮我去新增" / "新增款号审批",
  mentions 拆借新增款号审批表 or the kdocs share link, asks to fill a 月份 sheet (10月/9月/…)
  in the shared approval workbook, or asks to sync such a table to Kdocs. Also use it when
  they reference a mapping like "202304-6SetA 对应 202304-3BHP*2" or "102207-3B（102207-1B*3）".
---

# 新增款号审批表

把一批新增款号，按共享审批表的固定版式生成明细行；需要时直接用 `kdocs-cli` 写入线上。

## 何时用

- 用户给出一串 `TM-BH-xxx` 款号，说「帮我去新增」「这些款号帮我去新增」
- 要往「2026拆借新增款号审批表」的某个月份 sheet（10月 / 9月 / …）补数据
- 用户给映射说明，如 `TM-BH-202304-6SetA（对应大表 202304-3BHP*2）`、`102207-3B（102207-1B*3）`
- 要把本地表同步到 Kdocs 共享文档

## 固定版式（不要自行改动）

10 列，顺序固定：

`日期 | 产品大类 | 产品中类 | 产品小类 | 款号 | MSKU | 申请人 | 备注 | 审批 | 操作人`

| 项目 | 规则 |
|---|---|
| 日期 | 当天，写成 `M.D日`（如 `10.31日`）；**只写在整块第一行**，同一天不重复填 |
| 款号 | `TM-BH-` + 款号主体（如 `TM-BH-N103507`） |
| MSKU | `款号-段-尺码`，如 `TM-BH-N103507-3B-S`。**注意段必须保留**，漏掉会变成 `-S` |
| 尺码 | 该段在源表里实际存在的尺码，通常 `S/M/L/XL/XXL`；T1036 类只有 `S/M/L/XL` |
| 申请人 | `缪广亮` |
| 备注 | `新增款号` |
| 审批 / 操作人 | 留空（由审批流程填） |
| 分组 | 一个款号一块；**块与块之间空一行**；同款号的不同段也各占一块 |
| 合并 | 申请人/备注/审批/操作人 4 列纵向合并覆盖整块数据区（第 2 行到最后一行） |
| 字体边框 | 宋体 11；细边框 `D9D9D9`；全表垂直居中；日期/款号/MSKU 居中，文字列左对齐 |
| 列宽 | `11, 30, 30, 58, 19, 25, 11, 11, 9, 9`（B/C 窄了会截字，已踩过） |

## ⚠️ 四个必须知道的坑

**1. 三张表的款号前缀不同，不能直接比 MSKU。**
`TM-BH-` / `TM-HW-` / `SHEIN-SO-`。必须先去掉前缀，按 `款号-段-尺码` 归一后再匹配。直接字符串比较会 **0 命中**。

**2. 源表「小款号」列不带前缀，「MSKU」列才带。**
按小款号匹配时要先补 `TM-BH-`，否则查不到。

**3. 三表分类口径不一致，回退要按字段逐个兜底。**
BH 与 SO 共有的 560 个键里有 230 个分类不同（如 `102701-5SetB` 在 BH 是「男士速干Polo衫」，SO 是「男士速干」）。规则：**永远优先十部表，该字段缺失才依次回退 HW、SO**。不要整行覆盖。

**4. 用户给的「映射」要自己推导小类，且不要猜错段。**
- `202304-6SetA` 对应 `202304-3BHP*2` → 小类沿用 3BHP 的配色：`女士牛奶丝双面磨毛短款短袖-六件套-黑灰粉`
- `102207-3B`（`102207-1B*3`）→ `男士罗纹短袖-3黑`
- `102105-3B`（`102105-1B*3`）→ `男士速干连帽-细珠地布口罩款-3黑`

推导出的命名要单独列给用户确认，不要静默写入。

## 用法

### 1. 查分类（三级回退）

```bash
python scripts/resolve_classes.py --msku TM-BH-N103507-3B-S
python scripts/resolve_classes.py --file keys.txt   # 每行一个 MSKU
```

输出每个字段的取值和**来源表**，便于核对。

### 2. 生成审批表

```bash
python scripts/build_approval.py --spec styles.json --out 新增款号审批表.xlsx
```

`styles.json`：

```json
[
  {"body": "202304", "seg": "6SetA", "sizes": ["S","M","L","XL","XXL"],
   "small": "女士牛奶丝双面磨毛短款短袖-六件套-黑灰粉",
   "basis": "映射 202304-3BHP*2"},
  {"body": "204004", "seg": "1B", "sizes": ["S","M","L","XL","XXL"]}
]
```

- 不写 `大类/中类/小类` 时，脚本用 `resolve_classes.py` 的三级回退自动取。
- 写了 `small` 等字段则原样使用（用于映射推导的场景）。
- `basis` 只用于日志，不写入表格。

### 3. 写入 Kdocs 共享表

先确认 `kdocs-cli` 可用（`kdocs-cli version`），未安装见 `references/kdocs-cli-setup.md`。

```bash
python scripts/push_to_kdocs.py --url <共享表URL> --sheet 10月 --xlsx 新增款号审批表.xlsx
```

**写入前必须做**：
1. `sheet get-sheets-info` 确认目标 sheet 存在，记下 `sheetId`
2. `sheet get-range-data` 读出该区域现有内容
3. 若已有内容且与待写入不一致 → **停下来问用户**，不要覆盖。审批痕迹（如「同意」）不可覆盖。
4. 写入后 `get-range-data` 回读比对，不一致就报告，不要声称成功

## 交付

- 本地生成时：`present` 该 xlsx
- 线上写入时：给出回读验证结果 + 表格链接

## 参考

- `references/kdocs-cli-setup.md` — kdocs-cli 安装/登录/常用命令
- `references/field-rules.md` — 各款号的分类口径与历史确认结论
