money_sum = {}
money_sum["奶茶"] = 25
money_sum["奶茶"] = money_sum["奶茶"] + 18
money_sum["电影"] = 50

for name in money_sum:
    print(name, "花了", money_sum[name], "元")