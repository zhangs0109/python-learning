# 光伏项目基础数据表 · 分析脚本 v1
# 五步：① 读数与清洗  ② 数据体检（找输入错误）  ③ 控制总额对拍  ④ 分类汇总  ⑤ 出表 + 出图
#
# 用法（在 project 目录下）：
#   source ../.venv/bin/activate
#   python3 pv_analysis.py
#
# 产出：
#   光伏项目_分析产出.xlsx   （10 个 Sheet，含原生图表）
#   光伏项目_分析看板.html   （网页版对比分析图，浏览器直接打开）

import os
import re
import json
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "20261101  光伏项目基础数据表  V1.xlsx")
OUT_XLSX = os.path.join(BASE, "光伏项目_分析产出.xlsx")
OUT_HTML = os.path.join(BASE, "光伏项目_分析看板.html")
CNY_RATE = 7.2306          # 表内折算美元汇率（1 USD = 7.2306 CNY）

# 表内嵌在第一行里的控制总额（用来跟我们算的对拍；口径已逐个还原验证）
CTRL = {
    "合计折算美元": 50535404.3500001,
    "折人民币": 365401294.693111,
    "调遣及五险一金": 89892977.34,
    "全部原币合价(混币种)": 5682134010.2666,
    "全部原币金额含调遣(混币种)": 5772026987.6066,
}

# ---------- ① 读数与清洗 ----------
df = pd.read_excel(SRC, sheet_name="明细-按月更新", header=2)     # 第 3 行才是真表头
df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]  # 去掉空列
df = df.dropna(how="all")
df.insert(0, "源行号", df.index + 4)          # 对应 WPS/Excel 里的真实行号

total_raw_rows = len(df)
# 只保留有支出类别的正式数据行（表尾合计行、分隔行、空白行会被剔除）
df = df[df["支出类别-1级"].notna()].copy()
dropped_rows = total_raw_rows - len(df)

for c in ["数量", "原币单价", "原币合价", "折算美元汇率", "合计折算美元", "调遣合计/企业承担五险一金"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df["调遣合计/企业承担五险一金"] = df["调遣合计/企业承担五险一金"].fillna(0)

# 计价单位归一：KG / Kg / kg / 千克 → KG；立方米 / m3 → m3；平米 / ㎡ / m2 → m2
UNIT_MAP = {"KG": "KG", "Kg": "KG", "kg": "KG", "千克": "KG",
            "立方米": "m3", "m3": "m3", "平米": "m2", "㎡": "m2", "m2": "m2",
            "米": "m", "m": "m"}
df["单位(标准化)"] = df["计价单位"].map(lambda x: UNIT_MAP.get(str(x).strip(), x))

# 月份标准化：'2024年9月' → 2024-09
def norm_month(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, )):
        return v.strftime("%Y-%m")
    s = str(v).strip()
    m = re.match(r"^(\d{4})年(\d{1,2})月", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})/(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return None

df["月份"] = df["日期（年月日）"].map(norm_month)

# ---------- ② 数据体检：找输入错误 ----------
issues = []   # 每一项：类型 / 说明 / 影响行数 / 明细(源行号)

def add_issue(kind, desc, rows, detail=""):
    issues.append({"问题类型": kind, "说明": desc, "涉及行数": len(rows),
                   "涉及源行号": detail})

# 1) 完全重复录入
body_cols = [c for c in df.columns if c not in ("源行号", "单位(标准化)", "月份")]
dup_mask = df.duplicated(subset=body_cols, keep=False)
dup_rows = df.loc[dup_mask, "源行号"].tolist()
dup_extra_usd = df.loc[df.duplicated(subset=body_cols, keep="first"), "合计折算美元"].sum()
add_issue("重复录入", "整行完全相同的记录，疑似重复记账", dup_rows,
          "、".join(str(x) for x in dup_rows[:40]) + ("…" if len(dup_rows) > 40 else ""))

# 2) 日期格式不统一 / 非法
bad_date = df.loc[df["月份"].isna() & df["日期（年月日）"].notna(),
                  ["源行号", "日期（年月日）"]]
add_issue("日期异常", "日期无法识别（格式不统一或日期不存在）",
          bad_date["源行号"].tolist(),
          "；".join(f"{r}:{v}" for r, v in bad_date.head(12).values))

# 3) 日期列混入真实日期格式（与文本年月混用）
mixed = df.loc[df["日期（年月日）"].apply(lambda x: isinstance(x, pd.Timestamp)), "源行号"].tolist()
if mixed:
    add_issue("日期格式混用", "该列应为文本年月（如 2025年3月），此处填成了标准日期",
              mixed, "、".join(str(x) for x in mixed[:30]))

# 4) 计价单位写法不统一
unit_alias = df.loc[df["计价单位"].isin(["KG", "Kg", "kg", "千克", "立方米", "m3", "平米", "㎡", "m2", "米", "m"]),
                    ["源行号", "计价单位"]]
add_issue("单位不统一", "同一计量单位有多种写法（KG/Kg/kg/千克、立方米/m3、平米/㎡/m2）",
          unit_alias["源行号"].tolist(),
          "；".join(f"{r}:{v}" for r, v in unit_alias.head(12).values))

# 5) 单位列混入数值
bad_unit = df.loc[pd.to_numeric(df["计价单位"], errors="coerce").notna(), ["源行号", "计价单位"]]
if len(bad_unit):
    add_issue("单位列填错", "计价单位列填成了数字", bad_unit["源行号"].tolist(),
              "；".join(f"{r}:{v}" for r, v in bad_unit.values))

# 6) 数量为负 / 金额为零或负（冲销、调拨、退货，需人工确认）
neg = df.loc[(df["数量"] < 0) | (df["原币合价"] < 0), ["源行号", "项目名称", "数量", "原币合价", "备注"]]
add_issue("负数/冲销", "数量为负或金额为负（多为调拨、退货、冲销，请确认是否应冲减成本）",
          neg["源行号"].tolist(),
          "；".join(f"{r}:{n}" for r, n in zip(neg["源行号"].head(12), neg["项目名称"].head(12))))
zero_qty = df.loc[(df["数量"] == 0) & (df["原币单价"] > 0), ["源行号", "项目名称"]]
if len(zero_qty):
    add_issue("数量为零", "数量为 0 但有单价，合价被记为 0", zero_qty["源行号"].tolist(),
              "；".join(f"{r}:{n}" for r, n in zero_qty.head(10).values))

# 7) 数量×单价 ≠ 合价
calc_amt = (df["数量"] * df["原币单价"]).round(2)
mism = df.loc[(df["原币合价"] - calc_amt).abs() > 1, ["源行号", "项目名称", "数量", "原币单价", "原币合价", "备注"]]
add_issue("合价与量价不符", "原币合价 ≠ 数量 × 单价（合价被手工改写，公式被破坏）",
          mism["源行号"].tolist(),
          "；".join(f"{r}:{n}" for r, n in zip(mism["源行号"].head(10), mism["项目名称"].head(10))))

# 8) 汇率异常（0 或缺失）
bad_rate = df.loc[(df["折算美元汇率"].isna()) | (df["折算美元汇率"] <= 0), "源行号"].tolist()
if bad_rate:
    add_issue("汇率异常", "折算美元汇率为 0 或空，无法折算", bad_rate,
              "、".join(str(x) for x in bad_rate[:30]))

# 9) 币种与汇率不匹配
expect = {"CNY": 7.2306, "DZD": 135.0, "USD": 1.0}
bad_pair = df.loc[df.apply(lambda r: (r["原币种"] in expect) and (abs(r["折算美元汇率"] - expect[r["原币种"]]) > 0.01)
                           if pd.notna(r["折算美元汇率"]) else False, axis=1), ["源行号", "原币种", "折算美元汇率"]]
add_issue("币种汇率不匹配", "币种与汇率不对应（如 CNY 配了 135 的汇率）",
          bad_pair["源行号"].tolist(),
          "；".join(f"{r}:{c}/{v}" for r, c, v in bad_pair.head(10).values))

# 10) 关键字段缺失
for col in ["对方单位名称", "项目名称", "计价单位", "编码", "费用归集人员", "支付方式-1级"]:
    miss = df.loc[df[col].isna(), "源行号"].tolist()
    if miss:
        add_issue("字段缺失", f"「{col}」为空", miss,
                  "、".join(str(x) for x in miss[:30]) + ("…" if len(miss) > 30 else ""))

# 11) 序号列全空
if df["序号"].notna().sum() == 0:
    add_issue("字段缺失", "「序号」列整列为空，无法按序号追溯", [], "整列")

# 12) 成本归属被调整（原口径 vs 现口径）
adj = df.loc[df["成本归属（原）"] != df["成本归属"], ["源行号", "成本归属（原）", "成本归属"]]
add_issue("口径调整", "成本归属被人工调整过（原值 ≠ 现值），影响归属口径对比",
          adj["源行号"].tolist(),
          "；".join(f"{r}:{a}→{b}" for r, a, b in adj.head(10).values))

issues_df = pd.DataFrame(issues)

# ---------- ③ 控制总额对拍（证明算得准） ----------
usd_sum = df["合计折算美元"].sum()
recalc = ((df["原币合价"] + df["调遣合计/企业承担五险一金"]) / df["折算美元汇率"].replace(0, np.nan)).round(2)
formula_ok = int((df["合计折算美元"] - recalc).abs().le(1).sum())
formula_bad = int((df["合计折算美元"] - recalc).abs().gt(1).sum())

checks = [
    ["合计折算美元", round(usd_sum, 2), CTRL["合计折算美元"], round(usd_sum - CTRL["合计折算美元"], 2),
     "一致" if abs(usd_sum - CTRL["合计折算美元"]) < 1 else "有差异"],
    ["折算人民币(×7.2306)", round(usd_sum * CNY_RATE, 2), CTRL["折人民币"],
     round(usd_sum * CNY_RATE - CTRL["折人民币"], 2),
     "一致" if abs(usd_sum * CNY_RATE - CTRL["折人民币"]) < 1 else "有差异"],
    ["调遣及五险一金合计", round(df["调遣合计/企业承担五险一金"].sum(), 2), CTRL["调遣及五险一金"],
     round(df["调遣合计/企业承担五险一金"].sum() - CTRL["调遣及五险一金"], 2),
     "一致" if abs(df["调遣合计/企业承担五险一金"].sum() - CTRL["调遣及五险一金"]) < 1 else "有差异"],
    ["全部原币合价(混币种)", round(df["原币合价"].sum(), 2), CTRL["全部原币合价(混币种)"],
     round(df["原币合价"].sum() - CTRL["全部原币合价(混币种)"], 2),
     "一致" if abs(df["原币合价"].sum() - CTRL["全部原币合价(混币种)"]) < 1 else "有差异"],
    ["全部原币金额含调遣(混币种)", round(df["原币合价"].sum() + df["调遣合计/企业承担五险一金"].sum(), 2),
     CTRL["全部原币金额含调遣(混币种)"],
     round(df["原币合价"].sum() + df["调遣合计/企业承担五险一金"].sum() - CTRL["全部原币金额含调遣(混币种)"], 2),
     "一致" if abs(df["原币合价"].sum() + df["调遣合计/企业承担五险一金"].sum() - CTRL["全部原币金额含调遣(混币种)"]) < 1 else "有差异"],
    ["折算公式还原", formula_ok, formula_bad, 0,
     "全部吻合" if formula_bad == 0 else f"{formula_bad} 条不符"],
]
check_df = pd.DataFrame(checks, columns=["校验项", "本次计算值", "表内控制值/参照", "差异", "结论"])

# ---------- ④ 分类汇总 ----------
USD = "合计折算美元"

def agg_by(cols, name=None):
    # 注意：数量不跨单位求和（KG、台班、月混在一起相加没有意义），故不列入
    g = df.groupby(cols, dropna=False).agg(
        笔数=(USD, "count"),
        美元金额=(USD, "sum"),
        原币合价=("原币合价", "sum"),
    ).reset_index()
    g["美元金额"] = g["美元金额"].round(2)
    g["原币合价"] = g["原币合价"].round(2)
    g["占比%"] = (g["美元金额"] / g["美元金额"].sum() * 100).round(2)
    return g.sort_values("美元金额", ascending=False).reset_index(drop=True)

cat1 = agg_by(["支出类别-1级"])
cat2 = agg_by(["支出类别-1级", "支出类别-2级"])
cat3 = agg_by(["支出类别-2级", "支出类别-3级"])
lot = agg_by(["标段/综合"])
pay = agg_by(["支付方式-1级"])
monthly = agg_by(["月份"]).sort_values("月份").reset_index(drop=True)
party = agg_by(["对方单位名称"]).head(20)

# 成本归属：原口径 vs 调整后口径 对比
old = df.groupby("成本归属（原）")[USD].sum().round(2)
new = df.groupby("成本归属")[USD].sum().round(2)
owner = pd.DataFrame({"成本归属": sorted(set(old.index.dropna()) | set(new.index.dropna()))})
owner["原口径金额"] = owner["成本归属"].map(old).fillna(0).round(2)
owner["调整后金额"] = owner["成本归属"].map(new).fillna(0).round(2)
owner["差额"] = (owner["调整后金额"] - owner["原口径金额"]).round(2)
owner = owner.sort_values("调整后金额", ascending=False).reset_index(drop=True)

# 材料：同一材料不同单价（呼应"奶茶不同价格"问题）
mat = df[df["支出类别-2级"] == "材料"].copy()
mat_agg = mat.groupby(["项目名称", "单位(标准化)", "原币种"]).agg(
    笔数=("原币单价", "count"),
    单价种类=("原币单价", "nunique"),
    最低价=("原币单价", "min"),
    最高价=("原币单价", "max"),
    平均价=("原币单价", "mean"),
    数量合计=("数量", "sum"),
    美元金额=(USD, "sum"),
).reset_index()
mat_agg["平均价"] = mat_agg["平均价"].round(2)
mat_agg["美元金额"] = mat_agg["美元金额"].round(2)
mat_agg["价格差异%"] = ((mat_agg["最高价"] - mat_agg["最低价"]) / mat_agg["最低价"] * 100).round(1)
mat_agg["价格异常"] = mat_agg["单价种类"] > 1
mat_agg = mat_agg.sort_values("美元金额", ascending=False).reset_index(drop=True)

# 字段说明书（列逻辑 + 命名意义）
meta = [
    ["序号", "标识列", "唯一编号，用于追溯和排序", "全表为空 → 无法按序号定位，建议补录", ""],
    ["支付方式-1级", "维度", "资金支付渠道：厂存/现金/水电代付等", "区分资金来源与是否占用项目部现金", "21.4% 为空"],
    ["支出类别-1级", "维度(一级)", "成本顶层分类：直接工程费/间接费/其他直接费", "划分成本属性，是成本分析的总纲", ""],
    ["支出类别-2级", "维度(二级)", "一级下面的细分：材料/船机/分包/不可控间接费等", "定位成本中心，决定归集口径", ""],
    ["支出类别-3级", "维度(三级)", "二级再细分：工程材料、管理人员-职工、租赁船机等", "细化到人、机、料的颗粒度", ""],
    ["标段/综合", "维度", "LOT6/LOT9/LOT15/综合/不计入成本", "把成本摊到具体标段，是标段盈亏核算的关键", ""],
    ["成本归属（原）", "维度(原值)", "最初的责任归属单位", "保留原值用于追溯调整轨迹", "514 条被调整"],
    ["成本归属", "维度(现值)", "调整后的责任归属单位", "实际考核与分摊口径", ""],
    ["编码", "标识列", "材料/费用编码", "与物资系统对接的主键，便于对账", "37.2% 为空"],
    ["日期（年月日）", "维度(时间)", "所属年月（文本格式）", "按月归集成本、做趋势分析", "存在格式混用"],
    ["对方单位名称", "维度", "供应商/分包商/内部单位", "供应商集中度分析、往来核对", "10.1% 为空"],
    ["项目名称", "维度", "材料名或人员姓名（按行性质变化）", "人工类填人名、材料类填材料名", "语义双重"],
    ["规格型号/职位/工种", "维度", "材料规格或人员岗位", "人工成本按岗位分析、材料按规格比价", ""],
    ["计价单位", "维度", "计量单位（月/台班/KG/项…）", "防止量纲混算，是单价可比的前提", "写法不统一"],
    ["数量", "度量", "消耗/结算数量", "与单价相乘得到合价", "含负数冲销"],
    ["原币种", "维度", "CNY / DZD / USD", "多币种核算的基础", ""],
    ["原币单价", "度量", "原始币种单价", "价格分析、涨价追踪", ""],
    ["原币合价", "度量", "数量 × 单价", "本笔原始金额（公式列，不要手改）", "4 条被改写"],
    ["调遣合计/企业承担五险一金", "度量", "人工附加成本：调遣费、企业承担社保", "人工总成本的隐性部分，必须并入", "50.6% 为空"],
    ["折算美元汇率", "度量", "原币 → 美元 的汇率（CNY 7.2306 / DZD 135 / USD 1）", "统一口径成美元，做跨国可比", ""],
    ["合计折算美元", "度量(核心)", "(原币合价 + 调遣/五险一金) ÷ 汇率", "全表唯一可比口径，所有分析基于它", ""],
    ["备注", "说明", "人工批注", "解释异常行（调拨、退货、预计）", ""],
    ["对应代付号", "标识列", "代付单据编号", "与代付台账勾稽", ""],
    ["费用归集人员", "维度", "录入/归集责任人", "追溯责任、核对录入质量", ""],
    ["日期（年月日）.1 及后续工资列", "辅明细", "工资、奖金、海龄补贴等人工明细", "人工成本的完整构成，仅工资类行填写", "仅约 1600 行有值"],
]
meta_df = pd.DataFrame(meta, columns=["列名", "列角色", "内容解释", "设置这个名称的业务意义", "本表发现的问题"])

# ---------- ⑤ 出表 ----------
summary = pd.DataFrame({
    "指标": ["正式数据行", "被剔除的非数据行", "完全重复行", "疑似问题项", "合计折算美元",
             "折人民币(万元)", "涉及月份数", "材料种类", "对方单位数"],
    "数值": [len(df), dropped_rows, len(dup_rows), len(issues_df), round(usd_sum, 2),
             round(usd_sum * CNY_RATE / 10000, 2), df["月份"].nunique(),
             mat_agg["项目名称"].nunique(), df["对方单位名称"].nunique()],
})

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
    summary.to_excel(w, sheet_name="00_总览", index=False)
    meta_df.to_excel(w, sheet_name="01_字段说明书", index=False)
    issues_df.to_excel(w, sheet_name="02_数据体检", index=False)
    check_df.to_excel(w, sheet_name="03_总额对拍", index=False)
    cat1.to_excel(w, sheet_name="04_支出类别分析", index=False, startrow=0)
    cat2.to_excel(w, sheet_name="04_支出类别分析", index=False, startrow=len(cat1) + 3)
    cat3.to_excel(w, sheet_name="04_支出类别分析", index=False, startrow=len(cat1) + len(cat2) + 6)
    lot.to_excel(w, sheet_name="05_标段分析", index=False)
    owner.to_excel(w, sheet_name="06_成本归属对比", index=False)
    monthly.to_excel(w, sheet_name="07_月度趋势", index=False)
    party.to_excel(w, sheet_name="08_对方单位TOP20", index=False)
    mat_agg.head(60).to_excel(w, sheet_name="09_材料价格分析", index=False)
    pay.to_excel(w, sheet_name="10_支付方式分析", index=False)

# 给体检表上色 + 加图表
wb = load_workbook(OUT_XLSX)
red = PatternFill("solid", fgColor="FFC7CE")
yellow = PatternFill("solid", fgColor="FFEB9C")

ws = wb["02_数据体检"]
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    n = row[2].value
    if isinstance(n, int):
        if n >= 100:
            for c in row:
                c.fill = red
        elif n >= 10:
            for c in row:
                c.fill = yellow
ws.column_dimensions["E"].width = 80

def add_bar(sheet, cat_col, val_col, nrows, title, anchor):
    w = wb[sheet]
    ch = BarChart()
    ch.type = "col"
    ch.title = title
    ch.height, ch.width = 9, 18
    data = Reference(w, min_col=val_col, min_row=1, max_row=nrows + 1)
    cats = Reference(w, min_col=cat_col, min_row=2, max_row=nrows + 1)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend = None
    w.add_chart(ch, anchor)

add_bar("04_支出类别分析", 1, 3, len(cat1), "支出类别-1级 金额分布(USD)", "H2")
add_bar("05_标段分析", 1, 3, len(lot), "各标段成本(USD)", "H2")
w = wb["06_成本归属对比"]
ch = BarChart(); ch.type = "col"; ch.title = "成本归属：原口径 vs 调整后(USD)"; ch.height, ch.width = 9, 18
data = Reference(w, min_col=2, max_col=3, min_row=1, max_row=len(owner) + 1)
cats = Reference(w, min_col=1, min_row=2, max_row=len(owner) + 1)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
w.add_chart(ch, "G2")

w = wb["07_月度趋势"]
ch = LineChart(); ch.title = "月度成本趋势(USD)"; ch.height, ch.width = 9, 20
data = Reference(w, min_col=3, min_row=1, max_row=len(monthly) + 1)
cats = Reference(w, min_col=1, min_row=2, max_row=len(monthly) + 1)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.legend = None
w.add_chart(ch, "H2")

add_bar("08_对方单位TOP20", 1, 3, len(party), "对方单位金额 TOP20(USD)", "H2")

for name in wb.sheetnames:
    s = wb[name]
    for col in s.columns:
        width = max((len(str(c.value)) for c in col[:60] if c.value is not None), default=8)
        s.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 46)
    s.freeze_panes = "A2"
wb.save(OUT_XLSX)

# ---------- 网页看板 ----------
def to_records(d, n=None):
    d = d.head(n) if n else d
    return json.dumps(d.to_dict(orient="records"), ensure_ascii=False, default=str)

html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>光伏项目成本分析看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
body{{font-family:-apple-system,"PingFang SC",sans-serif;margin:0;background:#F7F7F5;color:#2C2C2A}}
.wrap{{max-width:1180px;margin:0 auto;padding:28px 20px 60px}}
h1{{font-size:22px;font-weight:600;margin:0 0 6px}} .sub{{color:#6b6b66;font-size:13px;margin-bottom:22px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:26px}}
.kpi{{background:#fff;border:1px solid #E5E3DC;border-radius:12px;padding:16px 18px}}
.kpi .v{{font-size:22px;font-weight:600;color:#185FA5}} .kpi .l{{font-size:12px;color:#6b6b66;margin-top:4px}}
.cards{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.card{{background:#fff;border:1px solid #E5E3DC;border-radius:12px;padding:16px 18px;margin-bottom:16px}}
.card h3{{font-size:14px;font-weight:600;margin:0 0 12px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
th,td{{padding:6px 8px;border-bottom:1px solid #EFEDE6;text-align:left}}
th{{color:#5F5E5A;font-weight:500}} td.n{{text-align:right}}
.bad{{color:#A32D2D}} .ok{{color:#0F6E56}}
</style></head><body><div class="wrap">
<h1>光伏项目成本分析看板</h1>
<div class="sub">数据源：{SRC}　正式数据 {len(df)} 行　生成时间：2026-11-01</div>
<div class="kpis">
<div class="kpi"><div class="v">{usd_sum/10000:,.1f} 万 USD</div><div class="l">成本总额（折算美元）</div></div>
<div class="kpi"><div class="v">{usd_sum*CNY_RATE/100000000:,.2f} 亿</div><div class="l">折人民币（汇率 7.2306）</div></div>
<div class="kpi"><div class="v">{len(df):,}</div><div class="l">正式数据行</div></div>
<div class="kpi"><div class="v">{len(dup_rows)}</div><div class="l">疑似重复录入行</div></div>
<div class="kpi"><div class="v">{len(issues_df)}</div><div class="l">数据问题类别</div></div>
<div class="kpi"><div class="v">{df['月份'].nunique()}</div><div class="l">覆盖月份</div></div>
</div>
<div class="cards">
<div class="card"><h3>支出类别-1级 占比</h3><canvas id="c1" height="220"></canvas></div>
<div class="card"><h3>支出类别-2级 金额（USD）</h3><canvas id="c2" height="220"></canvas></div>
<div class="card"><h3>各标段成本（USD）</h3><canvas id="c3" height="220"></canvas></div>
<div class="card"><h3>成本归属：原口径 vs 调整后</h3><canvas id="c4" height="220"></canvas></div>
</div>
<div class="card"><h3>月度成本趋势（USD）</h3><canvas id="c5" height="240"></canvas></div>
<div class="card"><h3>对方单位金额 TOP15（USD）</h3><canvas id="c6" height="340"></canvas></div>
<div class="card"><h3>数据体检：问题清单</h3>
<table><tr><th>问题类型</th><th>说明</th><th class="n">涉及行数</th></tr>
{"".join(f"<tr><td>{r['问题类型']}</td><td>{r['说明']}</td><td class='n'>{r['涉及行数']}</td></tr>" for _, r in issues_df.iterrows())}
</table></div>
<div class="card"><h3>总额对拍（本次计算 vs 表内控制值）</h3>
<table><tr><th>校验项</th><th class="n">本次计算值</th><th class="n">表内控制值</th><th class="n">差异</th><th>结论</th></tr>
{"".join(f"<tr><td>{r['校验项']}</td><td class='n'>{r['本次计算值']}</td><td class='n'>{r['表内控制值/参照']}</td><td class='n'>{r['差异']}</td><td class='{'ok' if r['结论'] in ('一致','全部吻合') else 'bad'}'>{r['结论']}</td></tr>" for _, r in check_df.iterrows())}
</table></div>
</div>
<script>
const D1={to_records(cat1)}, D2={to_records(cat2)}, D3={to_records(lot)},
      D4={to_records(owner)}, D5={to_records(monthly)}, D6={to_records(party,15)};
const opt={{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}}}};
new Chart(c1,{{type:'doughnut',data:{{labels:D1.map(x=>x['支出类别-1级']),datasets:[{{data:D1.map(x=>x['美元金额'])}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'right'}}}}}}}});
new Chart(c2,{{type:'bar',data:{{labels:D2.map(x=>x['支出类别-2级']),datasets:[{{data:D2.map(x=>x['美元金额']),backgroundColor:'#378ADD'}}]}},options:opt}});
new Chart(c3,{{type:'bar',data:{{labels:D3.map(x=>x['标段/综合']),datasets:[{{data:D3.map(x=>x['美元金额']),backgroundColor:'#1D9E75'}}]}},options:opt}});
new Chart(c4,{{type:'bar',data:{{labels:D4.map(x=>x['成本归属']),datasets:[
  {{label:'原口径',data:D4.map(x=>x['原口径金额']),backgroundColor:'#B5D4F4'}},
  {{label:'调整后',data:D4.map(x=>x['调整后金额']),backgroundColor:'#185FA5'}}]}},
  options:{{responsive:true,maintainAspectRatio:false}}}});
new Chart(c5,{{type:'line',data:{{labels:D5.map(x=>x['月份']),datasets:[{{data:D5.map(x=>x['美元金额']),borderColor:'#D85A30',backgroundColor:'rgba(216,90,48,.12)',fill:true,tension:.25,pointRadius:2}}]}},options:opt}});
new Chart(c6,{{type:'bar',data:{{labels:D6.map(x=>x['对方单位名称']),datasets:[{{data:D6.map(x=>x['美元金额']),backgroundColor:'#7F77DD'}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}}}}}});
</script></div></body></html>"""

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)

print("===== 运行完成 =====")
print("正式数据行：", len(df), " 剔除非数据行：", dropped_rows)
print("合计折算美元：", round(usd_sum, 2), " 折人民币：", round(usd_sum * CNY_RATE / 10000, 2), "万元")
print("问题类别：", len(issues_df), " 重复行：", len(dup_rows), " 重复涉及金额(USD)：", round(dup_extra_usd, 2))
print()
print(check_df.to_string(index=False))
print()
print("支出类别-1级：")
print(cat1.to_string(index=False))
print()
print("标段：")
print(lot.to_string(index=False))
print()
print("产出：", OUT_XLSX, " / ", OUT_HTML)
