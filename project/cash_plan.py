# -*- coding: utf-8 -*-
"""
分包资金计划表 v1
====================================================================
解决的问题：每月要向公司报"下个月要备多少钱、先付谁、压谁"。

上一课（subcontract_settle.py）算的是「我欠多少」，
这一课算的是「下个月该付多少」——两者不是一回事：

    合同说"付 80%"，所以
      付款上限 = 累计计量 × 80%          ← 合同允许我付到这儿
      可付余额 = 付款上限 − 累计实付      ← 还想付，最多还能付这么多
      可付余额 < 0 → 已经付过 80% 了，本期必须压

再加上下月预计新发生的计量，就是下月资金需求：
    下月资金需求 = 可付余额（欠着的）+ 下月预计付款（新干的）

产出两张表：
    05_资金计划      —— 每家队伍一行，带优先级与建议
    06_币种资金需求  —— 按 DZD / CNY / USD 分别汇总（跨币种不能相加！）

本课新语法：groupby().tail(n)  取每组最后 n 行（用来算"近 3 期平均"）
====================================================================
"""

import os
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))

# ==================== 配置区 ====================
LEDGER_FILE = sys.argv[1] if len(sys.argv) > 1 else "分包结算台账_样例.xlsx"
DETAIL_SHEET = "结算明细"
CONTRACT_SHEET = "合同台账"
HEADER_ROW = 1
DEPOSIT_RATE = 0.05
PAY_RATE = 0.80          # 合同约定付款比例
RECENT_N = 3             # 用最近几期平均来预测下月计量
OUT_FILE = "资金计划_产出.xlsx"

KEY = ["合同编号", "结算期次"]
STD_RATE = {"CNY": 7.2306, "DZD": 135.0, "USD": 1.0}
# ================================================

LINE = "=" * 70
SUB = "-" * 70


def title(t):
    print("\n" + LINE)
    print("  " + t)
    print(LINE)


# ==================== 一、读表并做基础计算 ====================
title("第 0 步 ｜ 读取台账、复原每期数据")

path = os.path.join(BASE, LEDGER_FILE)
if not os.path.exists(path):
    print(f"  ✘ 找不到文件：{path}")
    print("  用法：python3 cash_plan.py 你的台账.xlsx")
    sys.exit(1)

df = pd.read_excel(path, sheet_name=DETAIL_SHEET, header=HEADER_ROW)
df = df[df["合同编号"].notna()].copy()
df["序号"] = df["序号"].astype(int)

# 重复结算先剔除（资金计划不能重复计钱）
before = len(df)
df = df.drop_duplicates(subset=KEY, keep="first")
print(f"  读取 {before} 条 → 剔除重复结算后 {len(df)} 条")

df = df.sort_values(["合同编号", "结算期次"]).reset_index(drop=True)

# 本期应付（原币）
df["本期应付"] = (df["本期计量(原币)"] - df["预付款扣回"]
                  - df["质保金扣留"] - df["其他扣款"])

# 累计（每个合同单独累计）
df["累计计量"] = df.groupby("合同编号")["本期计量(原币)"].cumsum()
df["累计应付"] = df.groupby("合同编号")["本期应付"].cumsum()
df["累计实付"] = df.groupby("合同编号")["本期实付"].cumsum()

# 合同付款上限口径
df["付款上限"] = df["累计计量"] * PAY_RATE
df["可付余额"] = df["付款上限"] - df["累计实付"]

# 预测下月计量：取最近 RECENT_N 期的平均值
# ★ 本课新语法：groupby().tail(n) —— 每组取最后 n 行
recent = df.groupby("合同编号").tail(RECENT_N)
avg_measure = recent.groupby("合同编号")["本期计量(原币)"].mean()
df["近3期平均计量"] = df["合同编号"].map(avg_measure)

print(f"  已按合同复原累计计量/应付/实付，并按近 {RECENT_N} 期均值预测下月计量")

# ==================== 二、取每家队伍的"当前状态" ====================
title("第 1 步 ｜ 每家队伍的当前状态（取最后一期）")

# 每个合同只保留最后一行（最新的累计数就是当前状态）
last = df.groupby("合同编号").tail(1).copy()

# 接合同额与币种
ct = pd.read_excel(path, sheet_name=CONTRACT_SHEET, header=1)
ct["合同额USD"] = ct["合同额(原币)"] / ct["合同汇率"]
contract_amt = dict(zip(ct["合同编号"], ct["合同额(原币)"]))
contract_usd = dict(zip(ct["合同编号"], ct["合同额USD"]))
contract_cur = dict(zip(ct["合同编号"], ct["币种"]))
contract_rate = dict(zip(ct["合同编号"], ct["合同汇率"]))
last["合同额(原币)"] = last["合同编号"].map(contract_amt)
last["合同额USD"] = last["合同编号"].map(contract_usd)
# 币种与汇率以【合同台账】为准 —— 明细里的币种可能录错（本样例就有一条）
last["明细币种"] = last["币种"]
last["币种"] = last["合同编号"].map(contract_cur)
last["汇率"] = last["合同编号"].map(contract_rate)

# 提示币种被纠正的
bad_cur = last[last["明细币种"] != last["币种"]]
for _, r in bad_cur.iterrows():
    print(f"  ⚠ {r['合同编号']}：明细最后一行币种写成 {r['明细币种']}，"
          f"已按合同台账纠正为 {r['币种']}")

last["完成率"] = last["累计计量"] / last["合同额(原币)"]
last["实付占计量比"] = last["累计实付"] / last["累计计量"]

# 下月预计付款 = 近3期平均计量 × 80%
last["下月预计付款"] = last["近3期平均计量"] * PAY_RATE

# 下月资金需求 = 可付余额 + 下月预计付款
# 可付余额为负（超付）时，自动冲减下月应付；冲成负数则本期无需付钱
last["下月资金需求"] = (last["可付余额"] + last["下月预计付款"]).clip(lower=0)


def decide(r):
    """优先级判断：返回 (级别, 建议)"""
    if r["可付余额"] < -0.01:
        return "一级·压付", f"已付超合同上限 {abs(r['可付余额']):,.0f}，本期不付并在下期扣回"
    if r["可付余额"] < 1:
        return "二级·暂停", "已付满 80% 上限，本期暂停，待产值增加后再付"
    if r["完成率"] >= 0.80:
        return "三级·优先", f"完成 {r['完成率']:.0%} 临近收尾，欠款需优先安排防纠纷"
    return "四级·按约", "按合同约定正常支付"


decisions = last.apply(decide, axis=1, result_type="expand")
last["优先级"] = decisions[0]
last["建议"] = decisions[1]

cols = ["优先级", "合同编号", "分包队伍", "币种", "合同额(原币)", "累计计量",
        "付款上限", "累计实付", "可付余额", "完成率", "实付占计量比",
        "近3期平均计量", "下月预计付款", "下月资金需求", "建议"]
show = last[cols].copy()
for c in ["合同额(原币)", "累计计量", "付款上限", "累计实付", "可付余额",
          "近3期平均计量", "下月预计付款", "下月资金需求"]:
    show[c] = show[c].round(0)

print()
print(show[["优先级", "分包队伍", "币种", "累计计量", "付款上限", "累计实付",
            "可付余额", "下月资金需求"]].to_string(index=False))

print("""
  怎么读这张表：
    付款上限  = 累计计量 × 80%，合同允许我付到这条线为止
    可付余额  = 付款上限 − 累计实付
                正数 → 还能付这么多；负数 → 已经付过头了
    下月资金需求 = 可付余额（欠着的）+ 下月预计付款（新干的）
""")

# ==================== 三、按币种汇总（跨币种不能相加） ====================
title("第 2 步 ｜ 按币种汇总资金需求（人民币和美元不能直接相加）")

by_cur = last.groupby("币种", as_index=False).agg(
    队伍数=("合同编号", "count"),
    可付余额净额=("可付余额", "sum"),
    下月预计应付=("下月预计付款", "sum"),
    下月资金需求合计=("下月资金需求", "sum"),
)
by_cur["汇率"] = by_cur["币种"].map(STD_RATE)
by_cur["折美元"] = by_cur["下月资金需求合计"] / by_cur["汇率"]
by_cur = by_cur.sort_values("折美元", ascending=False)

for c in ["可付余额净额", "下月预计应付", "下月资金需求合计", "折美元"]:
    by_cur[c] = by_cur[c].round(2)

print()
print(by_cur.to_string(index=False))
print(f"\n  下月资金需求合计 ≈ {by_cur['折美元'].sum():,.2f} 美元")
print("  （「可付余额净额」为负表示超付，已在下月资金需求中自动冲减）")

# ==================== 四、优先级清单 ====================
title("第 3 步 ｜ 付款优先级（先付谁、压谁）")

order = ["一级·压付", "二级·暂停", "三级·优先", "四级·按约"]
for lvl in order:
    part = last[last["优先级"] == lvl]
    if len(part) == 0:
        continue
    print(f"\n  【{lvl}】{len(part)} 家")
    for _, r in part.iterrows():
        print(f"    ■ {r['分包队伍']}（{r['合同编号']}）")
        print(f"        累计计量 {r['累计计量']:,.0f} ｜ 付款上限 {r['付款上限']:,.0f}"
              f" ｜ 累计实付 {r['累计实付']:,.0f} ｜ 可付余额 {r['可付余额']:,.0f} {r['币种']}")
        print(f"        完成率 {r['完成率']:.0%} ｜ 实付占计量 {r['实付占计量比']:.1%}")
        print(f"        建议：{r['建议']}")

# ==================== 五、导出 ====================
title("第 4 步 ｜ 导出 Excel")

plan_cols = ["优先级", "合同编号", "分包队伍", "币种", "合同额(原币)", "累计计量",
             "付款上限", "累计实付", "可付余额", "完成率", "实付占计量比",
             "近3期平均计量", "下月预计付款", "下月资金需求", "建议"]
out_path = os.path.join(BASE, OUT_FILE)
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    last[plan_cols].to_excel(writer, sheet_name="05_资金计划", index=False)
    by_cur[["币种", "汇率", "队伍数", "可付余额净额", "下月预计应付",
            "下月资金需求合计", "折美元"]].to_excel(
        writer, sheet_name="06_币种资金需求", index=False)

wb = load_workbook(out_path)
HEAD_FILL = PatternFill("solid", fgColor="2F5597")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=10)
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
YEL_FILL = PatternFill("solid", fgColor="FFF2CC")
GRN_FILL = PatternFill("solid", fgColor="E1F5EE")
GRN_FONT = Font(color="0F6E56", bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

for ws in wb.worksheets:
    head = [c.value for c in ws[1]]
    for c in ws[1]:
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"
    for i, name in enumerate(head, start=1):
        w = 16
        if name in ("建议",):
            w = 42
        elif name in ("分包队伍", "优先级"):
            w = 18
        ws.column_dimensions[get_column_letter(i)].width = w
    for i, name in enumerate(head, start=1):
        if name is None:
            continue
        is_money = any(k in str(name) for k in ("原币", "计量", "上限", "实付", "余额", "付款", "需求", "美元"))
        is_pct = ("率" in str(name) or "比" in str(name)) and "汇率" not in str(name)
        for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            cell = row[0]
            cell.border = BORDER
            if is_money and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
            if is_pct and isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"
    # 可付余额为负标红
    if "可付余额" in head:
        col = head.index("可付余额") + 1
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            if isinstance(row[0].value, (int, float)) and row[0].value < 0:
                row[0].fill = RED_FILL
                row[0].font = RED_FONT

# 优先级列上色
ws = wb["05_资金计划"]
head = [c.value for c in ws[1]]
pcol = head.index("优先级") + 1
for row in ws.iter_rows(min_row=2, min_col=pcol, max_col=pcol):
    v = row[0].value or ""
    if v.startswith("一级") or v.startswith("二级"):
        row[0].fill = RED_FILL
        row[0].font = RED_FONT
    elif v.startswith("三级"):
        row[0].fill = YEL_FILL
        row[0].font = Font(bold=True, color="854F0B")
    else:
        row[0].fill = GRN_FILL
        row[0].font = GRN_FONT

wb["05_资金计划"].sheet_properties.tabColor = "2F5597"
wb["06_币种资金需求"].sheet_properties.tabColor = "0F6E56"
wb.save(out_path)

print(f"  已生成：{OUT_FILE}")
print("    05_资金计划     —— 每家队伍：付款上限、可付余额、下月需求、优先级、建议")
print("    06_币种资金需求 —— 按币种汇总，跨币种不能相加，最后折美元看总额")

title("完成 ✔")
print("  用法：把下月要报的资金计划直接看「06_币种资金需求」那页的数字，")
print("        汇报时说明 DZD / CNY / USD 各需多少，再附「05_资金计划」的优先级清单。")
