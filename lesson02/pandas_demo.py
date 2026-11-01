# 第一次用 pandas：读自己的账本
# 运行前确保：1) 在虚拟环境里 (.venv)  2) 在 lesson02 目录下

import pandas as pd

# 读 CSV：一行搞定（不用 split、不用循环、不用字典归堆）
df = pd.read_csv("物品汇总.csv", encoding="utf-8-sig")

print("===== 原始表格 =====")
print(df)
print()

# 按金额排序（一句话）
print("===== 按合计金额从高到低排序 =====")
print(df.sort_values("合计金额", ascending=False))
print()

# 统计结论（一句话各一个）
print("===== 自动分析 =====")
print("总支出：", df["合计金额"].sum(), "元")
print("最贵的物品：", df.loc[df["合计金额"].idxmax(), "物品"])
print("平均单笔：", round(df["合计金额"].sum() / df["笔数"].sum(), 1), "元")
