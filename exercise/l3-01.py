# L3·输入分数输出A（≥90）/B（≥60）/C（≥30）/D（其余）
score = int(input("分数："))
if score >= 90:
    print("A")
elif score >= 60:
    print("B")
elif score >= 30:
    print("C")
else:
    print("D")
