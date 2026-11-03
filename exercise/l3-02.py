# L3-02·改成按”超耗率“分级：>5% 红、0-5% 黄、<0% 绿
rate = float(input("超耗率%："))
if rate > 5:
    print("红（超耗）")
elif rate >= 0:
    print("黄（关注）")
else:
    print("绿（正常）")