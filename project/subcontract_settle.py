# -*- coding: utf-8 -*-
"""
分包结算台账核算工具 v2
====================================================================
解决的问题：一张分包台账，横着看是每期的计量和付款，
竖着看必须回答三个问题——
    1. 这个队伍我一共该付他多少？（累计应付）
    2. 我一共已经付了多少？（累计实付）
    3. 现在是我欠他，还是他欠我？（应付余额）

本脚本做四件事：
    【一】数据体检 —— 重复结算、币种汇率错配、质保金漏扣、单期超付
    【二】算钱 —— 本期应付 → 累计应付 → 累计实付 → 应付余额
    【三】每个队伍一张平衡表 —— 合同额、产值完成率、付款进度
    【四】超付预警 + 与控制值对拍（证明算得对）

v2 相对 v1 的修正：
    ★ 付款进度改用正确分母：实付占计量比 = 累计实付 ÷ 累计计量，
      与合同约定的付款比例（如付80%）直接可比。
      v1 的"实付÷应付"分母不对，已纠正。

本课唯一新语法：cumsum()  分组内累计求和
====================================================================
"""

import os
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 脚本所在目录，保证在任何地方运行都能找到文件
BASE = os.path.dirname(os.path.abspath(__file__))

# ==================== 配置区（要改的都在这里） ====================
LEDGER_FILE = sys.argv[1] if len(sys.argv) > 1 else "分包结算台账_样例.xlsx"
DETAIL_SHEET = "结算明细"
CONTRACT_SHEET = "合同台账"
HEADER_ROW = 1              # 表头在 Excel 的第 2 行 → header 从 0 数，填 1
DEPOSIT_RATE = 0.05         # 合同约定的质保金比例 5%
PAY_RATE = 0.80             # 合同约定的付款比例（付计量额的 80%）
AUTO_REMOVE_DUP = True      # 体检发现重复结算后，是否按"保留第一条"处理
OUT_FILE = "分包结算_核算产出.xlsx"

# 业务主键：这两个字段组合起来，在一张台账里必须是唯一的
KEY = ["合同编号", "结算期次"]

# 各币种对美元的标准汇率（以合同/财务口径为准）
STD_RATE = {"CNY": 7.2306, "DZD": 135.0, "USD": 1.0}

# 控制值：来自台账「控制值」页（财务手工核算），用来对拍
CTRL = {
    "累计计量折美元": 3999943.30,
    "累计应付折美元": 3360473.34,
    "累计实付折美元": 3076063.12,
}
# ================================================================

LINE = "=" * 70
SUB = "-" * 70


def title(text):
    print("\n" + LINE)
    print("  " + text)
    print(LINE)


def section(text):
    print("\n" + text)
    print(SUB)


# ==================== 一、读表 ====================
title("第 0 步 ｜ 读取台账")

path = os.path.join(BASE, LEDGER_FILE)
if not os.path.exists(path):
    print(f"✘ 找不到台账文件：{path}")
    print("  请把 Excel 放在脚本同一个文件夹里，或在命令后面加文件名：")
    print("  python3 subcontract_settle.py 你的台账.xlsx")
    sys.exit(1)

raw = pd.read_excel(path, sheet_name=DETAIL_SHEET, header=HEADER_ROW)
before = len(raw)
# 台账最后一行是"合计"，它没有合同编号，去掉它（合计要自己算，不能拿来用）
raw = raw[raw["合同编号"].notna()].copy()
raw["序号"] = raw["序号"].astype(int)

print(f"  文件      ：{LEDGER_FILE}")
print(f"  读取      ：{before} 行 → 剔除合计行后剩 {len(raw)} 条结算记录")
print(f"  涉及合同  ：{raw['合同编号'].nunique()} 份 ｜ 队伍 {raw['分包队伍'].nunique()} 家 ｜ "
      f"期次 {raw['结算期次'].nunique()} 期")

# ==================== 二、数据体检 ====================
title("第 1 步 ｜ 数据体检（只报告，不擅自修改）")

issues = []      # 收集给"体检问题"页用
problem_found = 0

# --- 体检 1：重复结算（按业务主键查，不能带"序号"列）---
# 为什么不能用 df.duplicated()？因为"序号"列每行都不一样，
# 整行比对永远查不出重复。必须只挑业务字段来比。
dup = raw[raw.duplicated(subset=KEY, keep=False)]
section("① 重复结算（同一合同、同一期次出现多次）")
if len(dup) > 0:
    problem_found += 1
    print(f"  发现 {len(dup)} 条：")
    print(dup[["序号", "合同编号", "分包队伍", "结算期次", "本期计量(原币)", "本期实付"]]
          .to_string(index=False))
    print("  ↑ 请回台账核对：是同一笔录了两遍，还是同期次分两次计量（后者要改期次号）")
    dup_usd = (dup["本期计量(原币)"] / dup["汇率"]).sum() / 2   # 每条算两遍，除以2才是多记的那一份
    issues.append({
        "问题类型": "重复结算",
        "涉及记录": "、".join(f"{r['合同编号']}-{r['结算期次']}" for _, r in dup.iterrows()),
        "金额影响": f"虚增计量约 {dup_usd:,.2f} 美元",
        "处理建议": "回台账删除重复录入的一行（保留一条），重新跑脚本",
    })
else:
    print("  未发现 ✔")

# --- 体检 2：币种与汇率不匹配 ---
raw["标准汇率"] = raw["币种"].map(STD_RATE)
bad_rate = raw[(raw["标准汇率"] - raw["汇率"]).abs() > 0.0001]
section("② 币种与汇率不匹配")
if len(bad_rate) > 0:
    problem_found += 1
    print(f"  发现 {len(bad_rate)} 条：")
    print(bad_rate[["序号", "合同编号", "结算期次", "币种", "汇率", "标准汇率"]]
          .to_string(index=False))
    print("  ↑ 币种标错会让「按币种汇总」的报表全错（折美元金额可能是碰巧对的，更不能大意）")
    issues.append({
        "问题类型": "币种与汇率不匹配",
        "涉及记录": "、".join(f"{r['合同编号']}-{r['结算期次']}" for _, r in bad_rate.iterrows()),
        "金额影响": "按币种汇总的报表失真",
        "处理建议": "核对原始结算单，把币种或汇率改回正确值",
    })
else:
    print("  未发现 ✔")

# --- 体检 3：质保金比例异常（漏扣 = 白给钱）---
raw["质保金比例"] = (raw["质保金扣留"] / raw["本期计量(原币)"]).round(4)
bad_dep = raw[(raw["质保金比例"] - DEPOSIT_RATE).abs() > 0.0005]
section(f"③ 质保金比例异常（约定 {DEPOSIT_RATE:.0%}）")
if len(bad_dep) > 0:
    problem_found += 1
    print(f"  发现 {len(bad_dep)} 条：")
    print(bad_dep[["序号", "合同编号", "结算期次", "本期计量(原币)", "质保金扣留", "质保金比例"]]
          .to_string(index=False))
    for _, r in bad_dep.iterrows():
        should = r["本期计量(原币)"] * DEPOSIT_RATE
        gap = should - r["质保金扣留"]
        print(f"  第{r['序号']}条应扣 {should:,.0f}，实扣 {r['质保金扣留']:,.0f}，"
              f"少扣 {gap:,.0f} {r['币种']}")
        issues.append({
            "问题类型": "质保金漏扣",
            "涉及记录": f"{r['合同编号']}-{r['结算期次']}",
            "金额影响": f"少扣质保金 {gap:,.0f} {r['币种']}",
            "处理建议": "质保期满要退还质保金；现在少扣，将来结算时就没有钱可退",
        })
else:
    print("  未发现 ✔")

# --- 体检 4：单期超付（这一期付的钱比这一期该付的还多）---
raw["本期应付"] = (raw["本期计量(原币)"] - raw["预付款扣回"]
                   - raw["质保金扣留"] - raw["其他扣款"])
single_over = raw[raw["本期实付"] > raw["本期应付"]]
section("④ 单期超付（本期实付 > 本期应付）")
if len(single_over) > 0:
    problem_found += 1
    print(f"  发现 {len(single_over)} 条：")
    tmp = single_over[["序号", "合同编号", "结算期次", "本期应付", "本期实付"]].copy()
    tmp["超付金额"] = single_over["本期实付"] - single_over["本期应付"]
    print(tmp.to_string(index=False))
    print("  ↑ 多半是预付款未扣完、或有代垫费用没入账，需逐笔查清")
    for _, r in single_over.iterrows():
        gap = r["本期实付"] - r["本期应付"]
        issues.append({
            "问题类型": "单期超付",
            "涉及记录": f"{r['合同编号']}-{r['结算期次']}",
            "金额影响": f"本期多付 {gap:,.0f} {r['币种']}",
            "处理建议": "查预付款是否应继续扣回、代垫费用是否漏入账",
        })
else:
    print("  未发现 ✔")

print(f"\n  体检小结：发现 {problem_found} 类问题，共 {len(issues)} 条记录（详见导出表「体检问题」页）")

# ==================== 三、算钱 ====================
title("第 2 步 ｜ 算钱：本期应付 → 累计 → 余额")

# 先把折美元算出来（不同币种不能直接相加，必须先折算）
raw["本期计量USD"] = raw["本期计量(原币)"] / raw["汇率"]
raw["本期应付USD"] = raw["本期应付"] / raw["汇率"]
raw["本期实付USD"] = raw["本期实付"] / raw["汇率"]

# 对拍用：含重复行的全量合计
total_all = {
    "累计计量折美元": raw["本期计量USD"].sum(),
    "累计应付折美元": raw["本期应付USD"].sum(),
    "累计实付折美元": raw["本期实付USD"].sum(),
}

# 处理重复：按业务主键保留第一条
if AUTO_REMOVE_DUP and len(dup) > 0:
    work = raw.drop_duplicates(subset=KEY, keep="first").copy()
    print(f"  按配置剔除重复结算 {len(raw) - len(work)} 条，参与核算 {len(work)} 条")
    print("  （这是脚本替你做的决定，请在导出表里复核这几条是否真的重复）")
else:
    work = raw.copy()
    print(f"  未剔除重复，全部 {len(work)} 条参与核算")

# 按合同 + 期次排序，必须排！否则"累计"会按 Excel 里的乱序往上加
work = work.sort_values(["合同编号", "结算期次"]).reset_index(drop=True)

# ★ 本课新语法：cumsum() —— 分组内累计求和
work["累计应付USD"] = work.groupby("合同编号")["本期应付USD"].cumsum()
work["累计实付USD"] = work.groupby("合同编号")["本期实付USD"].cumsum()
work["应付余额USD"] = work["累计应付USD"] - work["累计实付USD"]

print("\n  逐期结算表（每个合同单独累计，单位：美元）：\n")
show = work[["合同编号", "分包队伍", "结算期次", "本期应付USD", "累计应付USD",
             "累计实付USD", "应付余额USD"]].copy()
for c in ["本期应付USD", "累计应付USD", "累计实付USD", "应付余额USD"]:
    show[c] = show[c].round(0).astype(int)
# 按合同分组打印，合同之间空一行，更像台账
last = None
for _, r in show.iterrows():
    if last is not None and r["合同编号"] != last:
        print()
    last = r["合同编号"]
    flag = "  ⚠" if r["应付余额USD"] < 0 else ""
    print(f"  {r['合同编号']}  {r['分包队伍']:<8} {r['结算期次']}  "
          f"应付 {r['本期应付USD']:>10,}  累计应付 {r['累计应付USD']:>11,}  "
          f"累计实付 {r['累计实付USD']:>11,}  余额 {r['应付余额USD']:>9,}{flag}")

# ==================== 四、队伍平衡表 ====================
title("第 3 步 ｜ 每家队伍一张平衡表（债权债务）")

bal = work.groupby(["合同编号", "分包队伍"], as_index=False).agg(
    结算笔数=("本期应付USD", "count"),
    累计计量USD=("本期计量USD", "sum"),
    累计应付USD=("本期应付USD", "sum"),
    累计实付USD=("本期实付USD", "sum"),
)
bal["应付余额USD"] = bal["累计应付USD"] - bal["累计实付USD"]

# 把合同额接进来（用字典映射，复用 L8 的字典知识，不用 merge）
ct = pd.read_excel(path, sheet_name=CONTRACT_SHEET, header=1)
ct["合同额USD"] = ct["合同额(原币)"] / ct["合同汇率"]
contract_usd = dict(zip(ct["合同编号"], ct["合同额USD"]))
bal["合同额USD"] = bal["合同编号"].map(contract_usd)

# 两个进度指标
bal["产值完成率"] = (bal["累计计量USD"] / bal["合同额USD"]).round(4)
# v2 修正：付款进度 = 累计实付 ÷ 累计计量，与合同约定的"付80%"直接可比
bal["实付占计量比"] = (bal["累计实付USD"] / bal["累计计量USD"]).round(4)
bal["约定付款比例"] = PAY_RATE


def judge(r):
    """队伍状态判断：先危险后提示"""
    flags = []
    if r["应付余额USD"] < -0.01:
        flags.append("超付，暂停付款")
    if r["实付占计量比"] > r["约定付款比例"] + 0.001:
        flags.append("付款快于合同约定")
    return "；".join(flags) if flags else "正常"


bal["状态判断"] = bal.apply(judge, axis=1)

money_cols = ["合同额USD", "累计计量USD", "累计应付USD", "累计实付USD", "应付余额USD"]
for c in money_cols:
    bal[c] = bal[c].round(2)

show_cols = ["合同编号", "分包队伍", "合同额USD", "累计计量USD", "产值完成率",
             "累计应付USD", "累计实付USD", "应付余额USD", "实付占计量比", "状态判断"]
print()
print(bal[show_cols].to_string(index=False))

print("""
  读表三句话：
    余额 > 0  → 我欠队伍（未付工程款，正常负债）
    余额 < 0  → 队伍欠我（超付了，风险！）
    实付占计量比 > 约定付款比例（80%） → 钱付快了，下一期要压
""")

# ==================== 五、超付预警 ====================
title("第 4 步 ｜ 预警清单")

warn = bal[bal["状态判断"] != "正常"]
if len(warn) > 0:
    print(f"  ⚠ {len(warn)} 家队伍需要关注：\n")
    for _, r in warn.iterrows():
        print(f"  ■ {r['分包队伍']}（{r['合同编号']}）")
        print(f"      合同额 {r['合同额USD']:,.0f} ｜ 完成 {r['产值完成率']:.1%} ｜ "
              f"累计应付 {r['累计应付USD']:,.0f} ｜ 累计实付 {r['累计实付USD']:,.0f}")
        print(f"      余额 {r['应付余额USD']:,.0f} ｜ 实付占计量 {r['实付占计量比']:.1%}"
              f"（约定 {PAY_RATE:.0%}）")
        print(f"      判断：{r['状态判断']}")
    print("\n  处理建议：暂停后续付款，在下期计量中扣回；")
    print("  同时查清是预付款未扣完，还是有代垫费用没有及时入账。")
else:
    print("  全部正常 ✔")

# ==================== 六、控制值对拍 ====================
title("第 5 步 ｜ 与台账控制值对拍（这一步决定你信不信这个结果）")

print(f"  {'控制项':<18}{'账上值':>16}{'脚本算出':>16}{'差异':>12}   结论")
print("  " + SUB)
all_ok = True
check_rows = []
for k, v in CTRL.items():
    calc = total_all[k]
    diff = calc - v
    ok = abs(diff) < 0.5
    if not ok:
        all_ok = False
    print(f"  {k:<18}{v:>16,.2f}{calc:>16,.2f}{diff:>12,.2f}   {'✔ 一致' if ok else '✘ 不一致'}")
    check_rows.append({"控制项": k, "账上值": v, "脚本算出": round(calc, 2),
                       "差异": round(diff, 2), "结论": "一致 ✔" if ok else "不一致 ✘"})
print("  " + SUB)

if all_ok:
    print("  三项全部对上 → 读表没读错、行没漏、口径没歪，后面的分析可以信。")
else:
    print("  有项目对不上 → 先别用分析结果！按下面顺序查：")
    print("    1) header 行数对不对（读进来先 head() 看一眼）")
    print("    2) 有没有漏读某些行（比如手工插的空行、分表）")
    print("    3) 汇率口径是不是和财务用的一致")

# 去重后的影响
gap = total_all["累计计量折美元"] - work["本期计量USD"].sum()
if abs(gap) > 0.5:
    print(f"\n  剔除重复结算后，累计计量减少 {gap:,.2f} 美元"
          f"（约合 {gap * 7.2306:,.0f} 元人民币）—— 这就是重复录入的真实影响。")

# ==================== 七、导出 ====================
title("第 6 步 ｜ 导出 Excel（5 个页签）")

# ---- 逐期结算页：按合同分块，每块末尾加一行小计 ----
export_cols = ["序号", "合同编号", "分包队伍", "工程部位", "结算期次", "币种", "汇率",
               "本期计量(原币)", "预付款扣回", "质保金扣留", "其他扣款", "本期实付",
               "本期应付", "备注", "质保金比例", "本期应付USD", "本期实付USD",
               "累计应付USD", "累计实付USD", "应付余额USD"]
blocks = []
for (no, team), g in work.groupby(["合同编号", "分包队伍"], sort=False):
    blocks.append(g[export_cols])
    sub = {c: "" for c in export_cols}
    sub["合同编号"] = f"{no} 小计"
    sub["分包队伍"] = team
    sub["本期应付USD"] = round(g["本期应付USD"].sum(), 2)
    sub["本期实付USD"] = round(g["本期实付USD"].sum(), 2)
    sub["应付余额USD"] = round(g["本期应付USD"].sum() - g["本期实付USD"].sum(), 2)
    blocks.append(pd.DataFrame([sub]))
detail_out = pd.concat(blocks, ignore_index=True)

issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(
    [{"问题类型": "未发现", "涉及记录": "-", "金额影响": "-", "处理建议": "-"}])
check_df = pd.DataFrame(check_rows)

out_path = os.path.join(BASE, OUT_FILE)
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    # 使用说明页先占位，后面用 openpyxl 写
    pd.DataFrame({"说明": ["（本页由 openpyxl 填充）"]}).to_excel(
        writer, sheet_name="00_使用说明", index=False)
    detail_out.to_excel(writer, sheet_name="01_逐期结算", index=False)
    bal[show_cols].to_excel(writer, sheet_name="02_队伍平衡表", index=False)
    issues_df.to_excel(writer, sheet_name="03_体检问题", index=False)
    check_df.to_excel(writer, sheet_name="04_对拍验证", index=False)

# ---- openpyxl 美化 ----
wb = load_workbook(out_path)

HEAD_FILL = PatternFill("solid", fgColor="2F5597")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=10)
SUBTOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
SUBTOTAL_FONT = Font(bold=True)
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
OK_FILL = PatternFill("solid", fgColor="E1F5EE")
OK_FONT = Font(color="0F6E56", bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_sheet(ws, money_keywords=("USD", "原币", "应付", "实付", "计量", "合同额", "扣")):
    head = [c.value for c in ws[1]]
    for c in ws[1]:
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    for i, name in enumerate(head, start=1):
        w = 22 if name and len(str(name)) > 8 else 15
        if name in ("备注", "处理建议", "涉及记录", "金额影响", "状态判断"):
            w = 30
        ws.column_dimensions[get_column_letter(i)].width = w
    # 金额与百分比格式 + 边框
    for i, name in enumerate(head, start=1):
        if name is None:
            continue
        is_money = any(k in str(name) for k in money_keywords)
        is_pct = ("率" in str(name) or "比例" in str(name)) and "汇率" not in str(name)
        for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            cell = row[0]
            cell.border = BORDER
            if is_money and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
            if is_pct and isinstance(cell.value, (int, float)):
                cell.number_format = "0.00%"
    # 负数（超付）标红
    if "应付余额USD" in head:
        col = head.index("应付余额USD") + 1
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            if isinstance(row[0].value, (int, float)) and row[0].value < 0:
                row[0].fill = RED_FILL
                row[0].font = RED_FONT


# 逐期结算页：小计行加底色加粗
ws = wb["01_逐期结算"]
style_sheet(ws)
head = [c.value for c in ws[1]]
no_col = head.index("合同编号") + 1
for row in ws.iter_rows(min_row=2, min_col=no_col, max_col=no_col):
    if isinstance(row[0].value, str) and row[0].value.endswith("小计"):
        r = row[0].row
        for cell in ws[r]:
            cell.fill = SUBTOTAL_FILL
            cell.font = SUBTOTAL_FONT

# 平衡表页：状态列上色
ws = wb["02_队伍平衡表"]
style_sheet(ws)
head = [c.value for c in ws[1]]
if "状态判断" in head:
    col = head.index("状态判断") + 1
    for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
        if row[0].value and row[0].value != "正常":
            row[0].fill = RED_FILL
            row[0].font = RED_FONT
        else:
            row[0].fill = OK_FILL
            row[0].font = OK_FONT

# 体检问题页、对拍验证页
ws = wb["03_体检问题"]
style_sheet(ws, money_keywords=())
ws = wb["04_对拍验证"]
style_sheet(ws, money_keywords=())
head = [c.value for c in ws[1]]
if "结论" in head:
    col = head.index("结论") + 1
    for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
        row[0].fill = OK_FILL
        row[0].font = OK_FONT
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if cell.column_letter == "D" and isinstance(cell.value, (int, float)):
            cell.number_format = "#,##0.00"

# 使用说明页：整体重写
ws = wb["00_使用说明"]
ws.delete_rows(1, ws.max_row)
ws.sheet_view.showGridLines = False

guide = [
    ("title", "分包结算核算产出 ｜ 使用说明"),
    ("blank", ""),
    ("h", "■ 这份文件是怎么来的"),
    ("t", "由脚本 subcontract_settle.py 自动生成。原始数据是你的《分包结算台账》，"),
    ("t", "本文件只做核算与标注，不改动原始台账。"),
    ("blank", ""),
    ("h", "■ 四条业务公式（看懂这四条，全表就都懂了）"),
    ("t", "  本期应付 = 本期计量 − 预付款扣回 − 质保金扣留 − 其他扣款    （这一期该给多少）"),
    ("t", "  累计应付 = 本合同历期「本期应付」相加                      （一共该给多少）"),
    ("t", "  累计实付 = 本合同历期「本期实付」相加                      （一共给了多少）"),
    ("t", "  应付余额 = 累计应付 − 累计实付                             （谁欠谁）"),
    ("blank", ""),
    ("h", "■ 符号方向（最容易看错的地方）"),
    ("t", "  应付余额 > 0  →  我欠队伍（未付工程款，正常负债）"),
    ("t", "  应付余额 < 0  →  队伍欠我（超付了，风险，已标红）"),
    ("blank", ""),
    ("h", "■ 各页签讲什么"),
    ("t", "  01_逐期结算   每期的应付、实付、累计、余额；黄色行 = 每个合同的小计"),
    ("t", "  02_队伍平衡表 每家队伍一行：合同额、完成率、付款进度、状态判断"),
    ("t", "  03_体检问题   台账里查出的问题清单（重复/币种/质保金/超付），附处理建议"),
    ("t", "  04_对拍验证   脚本结果与台账控制值的比对，三项一致才可采信"),
    ("blank", ""),
    ("h", "■ 换成你自己的真台账（三步）"),
    ("t", "  1. 把真台账放进 project 文件夹（列名与样例一致，表头在第 2 行）"),
    ("t", "  2. 把财务给的控制总额填进「控制值」页（折美元口径）"),
    ("t", "  3. 终端运行：python3 subcontract_settle.py 你的台账.xlsx"),
    ("blank", ""),
    ("h", "■ 颜色约定"),
    ("t", "  蓝色表头 = 字段名 ｜ 黄色行 = 合同小计 ｜ 红色 = 超付/异常 ｜ 绿色 = 正常/对拍一致"),
    ("blank", ""),
    ("t", "生成时间：2026-11-04 ｜ 汇率口径：CNY 7.2306 / DZD 135 / USD 1"),
]
for kind, text in guide:
    ws.append([text])

for r in range(1, ws.max_row + 1):
    c = ws.cell(row=r, column=1)
    v = c.value or ""
    if v.startswith("■") or (v.startswith("  ") and "=" in v and "→" not in v):
        pass
    if v.startswith("■"):
        c.font = Font(bold=True, size=12, color="2F5597")
    if "使用说明" in v and r == 1:
        c.font = Font(bold=True, size=16, color="1F1F1F")
ws.column_dimensions["A"].width = 110
for r in range(1, ws.max_row + 1):
    ws.cell(row=r, column=1).alignment = Alignment(vertical="center", wrap_text=False)
ws.row_dimensions[1].height = 34

# 页签颜色
wb["00_使用说明"].sheet_properties.tabColor = "854F0B"
wb["01_逐期结算"].sheet_properties.tabColor = "2F5597"
wb["02_队伍平衡表"].sheet_properties.tabColor = "2F5597"
wb["03_体检问题"].sheet_properties.tabColor = "C00000"
wb["04_对拍验证"].sheet_properties.tabColor = "0F6E56"

wb.save(out_path)

print(f"  已生成：{OUT_FILE}")
print("    00_使用说明    —— 四条公式、符号方向、换真台账三步")
print("    01_逐期结算    —— 每期明细 + 合同小计（黄）+ 超付标红")
print("    02_队伍平衡表  —— 每家队伍一行的债权债务 + 状态判断")
print("    03_体检问题    —— 问题清单与处理建议")
print("    04_对拍验证    —— 三项控制值，全绿才可采信")

title("完成 ✔")
print("  下一步：用 WPS 打开产出表，先读「00_使用说明」，再看「02_队伍平衡表」。")
print("  逻辑确认无误后，把真台账放进 project/ 直接运行即可。")
