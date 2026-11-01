# 临时探查脚本 4：公式还原、重复行、负值、控制总额对拍（用完可删）
import pandas as pd
import numpy as np

FILE = "/Users/zhangshuang/WorkBuddy/2026-10-28-16-21-13/python-learning/project/20261101  光伏项目基础数据表  V1.xlsx"
df = pd.read_excel(FILE, sheet_name="明细-按月更新", header=2)
df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
df = df.dropna(how="all")
df.insert(0, "源行号", df.index + 4)

amt = pd.to_numeric(df["原币合价"], errors="coerce")
rate = pd.to_numeric(df["折算美元汇率"], errors="coerce")
usd = pd.to_numeric(df["合计折算美元"], errors="coerce")
extra = pd.to_numeric(df["调遣合计/企业承担五险一金"], errors="coerce").fillna(0)

print("===== 公式还原：美元 = (合价 + 调遣/五险一金) ÷ 汇率 =====")
calc = ((amt + extra) / rate.replace(0, np.nan)).round(2)
d = (usd - calc).abs()
print("吻合：", int((d <= 1).sum()), " 不符：", int((d > 1).sum()), " 无法计算：", int(d.isna().sum()))
print(df.loc[d > 1, ["源行号", "支出类别-2级", "原币种", "原币合价", "调遣合计/企业承担五险一金", "折算美元汇率", "合计折算美元"]].head(8).to_string(index=False))
print()

print("===== 重复行（不含源行号）=====")
body = df.drop(columns=["源行号"])
dupmask = body.duplicated(keep=False)
print("完全重复行数：", int(dupmask.sum()))
cols = ["源行号", "支出类别-2级", "标段/综合", "项目名称", "数量", "原币合价", "日期（年月日）", "费用归集人员"]
print(df.loc[dupmask, cols].sort_values("项目名称").head(14).to_string(index=False))
print()

print("===== 负数量 / 负金额 上下文 =====")
qty = pd.to_numeric(df["数量"], errors="coerce")
neg = df.loc[(qty < 0) | (amt < 0), ["源行号", "支出类别-2级", "项目名称", "计价单位", "数量", "原币单价", "原币合价", "备注", "日期（年月日）"]]
print(neg.head(20).to_string(index=False))
print()

print("===== 表尾非数据行（合计行/分隔行）=====")
tail = df.tail(12)[["源行号", "项目名称", "原币合价", "合计折算美元", "备注"]]
print(tail.to_string(index=False))
print()

print("===== 控制总额对拍（表第 1 行内嵌的校验值）=====")
ctrl = {"合计折算美元": 50535404.3500001, "折人民币": 365401294.693111,
        "A": 102714685.199807, "B": 154072027799.711, "C": 5682134010.2666,
        "D": 89892977.34, "E": 5772026987.6066}
valid = df[df["支出类别-1级"].notna()]
print("我算的 合计折算美元：", round(usd.sum(), 4), " 表内：", ctrl["合计折算美元"])
print("我算的 折人民币(×7.2306)：", round(usd.sum() * 7.2306, 4), " 表内：", ctrl["折人民币"])
print("我算的 CNY 合价合计：", round(amt[df["原币种"] == "CNY"].sum(), 2))
print("我算的 DZD 合价合计：", round(amt[df["原币种"] == "DZD"].sum(), 2), " 表内 E：", ctrl["E"])
print("我算的 USD 合价合计：", round(amt[df["原币种"] == "USD"].sum(), 2))
print("我算的 调遣/五险一金合计：", round(extra.sum(), 2), " 表内 D：", ctrl["D"])
print()

print("===== 月度汇总（按文本日期）=====")
g = df.groupby("日期（年月日）").agg(笔数=("源行号", "count"), 美元金额=("合计折算美元", "sum")).round(2)
print(g.to_string())
