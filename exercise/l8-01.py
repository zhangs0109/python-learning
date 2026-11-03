# L8-01·有字典 dict 统计：奶茶 25，电影票 30，奶茶 10，按物品汇总
records = [("奶茶", 25),("电影票", 30),("奶茶", 10)]
money_sum = {}
for name, money in records:
    if name in money_sum:
        money_sum[name] = money_sum[name] + money
    else:
        money_sum[name] = money
print(money_sum)
