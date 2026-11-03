# L8-02·同一物品出现不同价格就报警（奶茶：25和10）
records = [("奶茶", 25),("电影票", 30),("奶茶", 10)]
prices = {}
for name, money in records:
    if name in prices:
        prices[name].append(money)
    else:
        prices[name] = [money]
for name in prices:
    if max(prices[name])!= min(prices[name]):
        print(name, "价格异常：", prices[name])
