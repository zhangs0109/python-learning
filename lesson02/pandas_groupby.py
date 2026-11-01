# pandas 第 2 课：从"原始流水"直接汇总，并导出 Excel
# 对比 lesson02/note01.py 里手写的字典版本（约 25 行）
# pandas 版本只需要下面这几行核心代码

import pandas as pd

# 1. 读原始流水：空格分隔、没有表头，所以用 sep=" " + names 给列起名
df = pd.read_csv("日记01.txt", sep=" ", names=["物品", "金额"], encoding="utf-8")

# 2. 脏数据兜底：把金额转成数字，转不了的变成空值，再整行丢掉
df["金额"] = pd.to_numeric(df["金额"], errors="coerce")
df = df.dropna()

print("===== 原始流水（每笔一行）=====")
print(df)
print()

# 3. 核心一句：按"物品"分组，一次性算出笔数/合计/平均/最高/最低
result = df.groupby("物品")["金额"].agg(["count", "sum", "mean", "max", "min"])
result.columns = ["笔数", "合计金额", "平均金额", "最高价", "最低价"]
result = result.round(1)

# 4. 加一列判断：同一物品出现不同价格就报警（L8 遗留钩子，一句搞定）
result["价格是否异常"] = result["最高价"] != result["最低价"]

print("===== 按物品汇总 =====")
print(result)
print()

# 5. 导出成真正的 Excel 文件（这就是你问过的"能不能直接出统计表"）
result.to_excel("物品汇总.xlsx")

print("已生成 Excel：物品汇总.xlsx")
print("数据总行数：", len(df), "笔，涉及物品", df["物品"].nunique(), "种")
print("总金额：", int(df["金额"].sum()), "元")
