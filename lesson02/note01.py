while True:
    choice = input("请选择：1 记账  2 查账  3 退出 ")

    if choice == "1":
        name = input("买了什么？ ")
        money = input("花了多少钱？ ")
        try:
            int(money)
        except:
            print("金额要输入数字哦，这笔没记上")
            continue
        f = open("日记01.txt","a", encoding="utf-8")
        f.write(name +" "+ money + "\n")
        f.close()
        print("记好啦！", name, money, "元")

    elif choice == "2":
        f = open("日记01.txt", "r", encoding="utf-8")
        lines = f.readlines()
        f.close()

        total = 0
        print("----账本明细----")
        money_sum = {}
        counts = {}
        for line in lines:
            parts = line.split()
            try:
                money = int(parts[1])
            except:
                print("（发现坏数据，已跳过：", line.strip(), "）")
                continue
            name = parts[0]
            print(name, money, "元")
            total = total + money
            if name in money_sum:
                money_sum[name] = money_sum[name] + money
                counts[name] = counts[name] + 1
            else:
                money_sum[name] = money
                counts[name] = 1

        print("----按物品汇总-----")
        for name in money_sum:
            print(name, "共", counts[name], "笔，合计", money_sum[name], "元,平均", round(money_sum[name]/counts[name], 1), "元")

    elif choice == "3":
        print("再见！")
        break