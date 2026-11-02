# L2-02·同样输入两笔，多算一个”平均“
a = float(input("第一笔："))
b = float(input("第二笔："))
total = (a + b)
avg = total/2
print("合计", round(total, 1), "元")
print("平均", round(avg, 1), "元")
