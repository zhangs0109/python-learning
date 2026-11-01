# 阶段四 · 物资消耗核算（超耗预警）v2
# 四关流水线：
#   ① 查重（甄别重复录入）  →  ② 同类项汇总  →  ③ 核算  →  ④ 预警并导出标红 Excel
#
# 用法：
#   python3 material_check.py                      # 默认读 物资台账.xlsx
#   python3 material_check.py 物资台账_多笔.xlsx   # 也可以命令行指定台账
#
# 运行前：终端提示符前要有 (.venv)，且当前目录是 project

import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------- 配置项（要改规则就改这里，不用翻下面的代码）----------
LEDGER_FILE = sys.argv[1] if len(sys.argv) > 1 else "物资台账.xlsx"
OUT_FILE = "物资消耗核算.xlsx"
THRESHOLD = 3              # 超耗阈值（%）：超这么多算"超耗"，超了但没到算"关注"
AUTO_REMOVE_DUP = True     # 疑似重复录入：True=剔除并提示 / False=只提示不剔除

# ---------- 读台账 ----------
df = pd.read_excel(LEDGER_FILE)
print("读入台账：", LEDGER_FILE, "，共", len(df), "行")
print()

# ---------- ① 查重：甄别"重复录入" ----------
# 整行一模一样 = 大概率是误录（同一笔录了两遍），列出来让人核对
dup = df[df.duplicated(keep=False)]
if len(dup) > 0:
    print("【疑似重复录入】以下记录整行完全相同，请回台账核对：")
    print(dup.to_string(index=False))
    if AUTO_REMOVE_DUP:
        df = df.drop_duplicates()
        print("→ 已按配置剔除，剩余", len(df), "行")
    else:
        print("→ 按配置保留原样（AUTO_REMOVE_DUP = False）")
    print()

# ---------- ② 同类项汇总：同一材料 + 同一部位的多次录入合并成一行 ----------
# 先算一个中间量：实耗金额，用来求加权平均单价
df["实耗金额"] = df["实耗量"] * df["单价"]

# dropna=False：部位为空的行也要留下，不能悄悄消失
grouped = df.groupby(["材料", "部位", "单位"], as_index=False, dropna=False).agg(
    录入笔数=("实耗量", "count"),     # 合并了几笔
    实耗量=("实耗量", "sum"),
    定额=("定额", "sum"),
    实耗金额=("实耗金额", "sum"),
)

# 加权平均单价 = 总金额 ÷ 总数量（同一材料不同批次价格不一样时用它）
grouped["单价"] = (grouped["实耗金额"] / grouped["实耗量"]).round(2)
df = grouped[["材料", "部位", "单位", "录入笔数", "实耗量", "定额", "单价"]]

raw_count = int(df["录入笔数"].sum())
print("【同类项汇总】", raw_count, "笔原始记录 → 合并成", len(df), "个核算单元")
print()

# ---------- ③ 核算 ----------
# 整列一起算，不用 for 循环（这叫向量化）
df["超耗量"] = df["实耗量"] - df["定额"]
df["超耗率%"] = (df["超耗量"] / df["定额"] * 100).round(1)
df["超耗金额"] = (df["超耗量"] * df["单价"]).round(2)
# 节约的不算"负成本"：没超耗的金额一律记 0
df.loc[df["超耗量"] < 0, "超耗金额"] = 0

# ---------- ④ 分级预警 ----------
df["预警等级"] = "正常"
df.loc[df["超耗率%"] > 0, "预警等级"] = "关注"
df.loc[df["超耗率%"] > THRESHOLD, "预警等级"] = "超耗"

# 按超耗金额从高到低排：贵的材料小幅超耗，比便宜材料大幅超耗更伤成本
df = df.sort_values("超耗金额", ascending=False)

print("===== 物资消耗核算表 =====")
print(df.to_string(index=False))
print()

print("===== 自动结论 =====")
print("核算单元：", len(df), "个（原始", raw_count, "笔）")
print("超耗项：", (df["预警等级"] == "超耗").sum(), "条")
print("关注项：", (df["预警等级"] == "关注").sum(), "条")
print("超耗金额合计：", round(df["超耗金额"].sum(), 2), "元")
print("重点盯控：", df.loc[df["超耗金额"].idxmax(), "材料"],
      "（", round(df["超耗金额"].max(), 2), "元，占比",
      round(df["超耗金额"].max() / df["超耗金额"].sum() * 100, 1), "%）")

# ---------- 导出 Excel ----------
df.to_excel(OUT_FILE, index=False)

# pandas 只负责"算"和"写"，"化妆"（标红）交给 openpyxl
wb = load_workbook(OUT_FILE)
ws = wb.active

style = {
    "超耗": (PatternFill("solid", fgColor="FFC7CE"), Font(color="9C0006", bold=True)),
    "关注": (PatternFill("solid", fgColor="FFEB9C"), Font(color="9C6500")),
}

# 从第 2 行开始（第 1 行是表头），看最后一列（预警等级）的值决定刷什么色
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    level = row[-1].value
    if level in style:
        fill, font = style[level]
        for cell in row:
            cell.fill = fill
            cell.font = font

wb.save(OUT_FILE)
print()
print("已生成带预警底色的 Excel：", OUT_FILE)
