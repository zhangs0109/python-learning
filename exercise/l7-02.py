# L7·查询时遇到坏行（金额不是数字）跳过 skip，不崩
f = open("流水.txt", "r", encoding="utf-8")
lines = f.readlines()
f.close()
total = 0
for line in lines:
    parts = line.split()
    try:
        money = int(parts[1])
    except:
        print("(跳过坏数据：", line.strip()," )")
        continue

    name =parts[0]
    money = int(parts[1])
    print(name, money, "元")
    total = total + money
print("合计", total, "元")
