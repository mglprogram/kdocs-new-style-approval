# kdocs-new-style-approval

把一批新增款号，按共享审批表的固定版式生成明细行；需要时直接用 `kdocs-cli` 写入线上金山文档。

## 内容

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 技能定义：触发条件、固定版式、四个必须知道的坑 |
| `scripts/resolve_classes.py` | 三表（BH → HW → SO）分级回退查分类 |
| `scripts/build_approval.py` | 按固定版式生成审批表 xlsx |
| `scripts/push_to_kdocs.py` | 经 kdocs-cli 写入线上共享表，带覆盖保护与回读验证 |
| `references/kdocs-cli-setup.md` | kdocs-cli 安装/登录/常用命令/错误速查 |
| `references/field-rules.md` | 分类口径与历史确认结论 |

## 安装

把整个目录放到技能根目录之一即可：

```bash
# 个人技能（所有项目可用）
~/.agents/skills/kdocs-new-style-approval/

# 项目技能
<project>/.agents/skills/kdocs-new-style-approval/
```

## 依赖

- Python 3 + `openpyxl`
- 线上写入需要 `kdocs-cli`（金山官方 CLI），见 `references/kdocs-cli-setup.md`

## 快速用法

```bash
# 1. 查分类（含来源表）
python scripts/resolve_classes.py --msku TM-BH-N103507-3B-S

# 2. 生成审批表
python scripts/build_approval.py --spec styles.json --out 新增款号审批表.xlsx

# 3. 写入线上（默认拒绝覆盖已有内容）
python scripts/push_to_kdocs.py --url <共享表链接> --sheet 10月 --xlsx 新增款号审批表.xlsx
```

## 关键约定

- 三张库存表款号前缀不同（`TM-BH-` / `TM-HW-` / `SHEIN-SO-`），必须先归一再去匹配
- 分类**按字段逐个回退**：十部表优先，缺失才用 HW、SO；不做整行覆盖
- 审批列由审批流程填写，写入时不可覆盖
- 写入后必须回读比对，不一致就报错

## License

MIT
