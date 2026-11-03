# L7·读流水文件（每行”名称 金额“），打印明细并算总额
f = open("流水.txt", "r", encoding="utf-8")
lines = f.readlines()
f.close()
total = 0
for line in lines:
    parts = line.split()
    name = parts[0]
    money = int(parts[1])
    print(name, money, "元")
    total = total + money
print("合计", total, "元")