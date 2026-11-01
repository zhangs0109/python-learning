# 演示：把账本数据变成"统计表 + CSV 文件"
# 用到的全是你已经学会的东西：字典、循环、文件读写

# 第一步：读账本，按物品归堆统计
money_sum = {}
counts = {}
prices = {}

f = open("日记01.txt", "r", encoding="utf-8")
lines = f.readlines()
f.close()

for line in lines:
    parts = line.split()
    if len(parts) < 2:
        continue
    name = parts[0]
    try:
        money = int(parts[1])
    except:
        continue

    if name in money_sum:
        money_sum[name] = money_sum[name] + money
        counts[name] = counts[name] + 1
        prices[name].append(money)
    else:
        money_sum[name] = money
        counts[name] = 1
        prices[name] = [money]

# 第二步：在终端打印一张对齐的统计表
print("=" * 62)
print(f"{'物品':<10}{'笔数':>6}{'合计':>10}{'平均':>10}{'最高':>10}{'最低':>10}")
print("-" * 62)

grand_total = 0
for name in money_sum:
    avg = round(money_sum[name] / counts[name], 1)
    print(f"{name:<10}{counts[name]:>6}{money_sum[name]:>10}{avg:>10}{max(prices[name]):>10}{min(prices[name]):>10}")
    grand_total = grand_total + money_sum[name]

print("-" * 62)
print(f"{'总计':<10}{sum(counts.values()):>6}{grand_total:>10}")
print("=" * 62)

# 第三步：写出 CSV，Excel 可以直接打开
# utf-8-sig 带 BOM，Windows/Excel 打开中文不乱码
f = open("物品汇总.csv", "w", encoding="utf-8-sig")
f.write("物品,笔数,合计金额,平均金额,最高价,最低价,是否价格异常\n")
for name in money_sum:
    avg = round(money_sum[name] / counts[name], 1)
    if max(prices[name]) != min(prices[name]):
        flag = "是（出现不同价格）"
    else:
        flag = "否"
    f.write(f"{name},{counts[name]},{money_sum[name]},{avg},{max(prices[name])},{min(prices[name])},{flag}\n")
f.close()

print("\n已生成：物品汇总.csv（双击用 Excel 或 Numbers 打开）")
