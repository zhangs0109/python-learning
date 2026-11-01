import random
secret=random.randint(1,100)
count=0
print("我心里想好了一个1~100的数字，来猜猜看！")
while True:
    text=input("你猜猜是几？")
    count=count+1
    try:
        guess=int(text)
    except:
        print("请输入数字！")
        continue
    if guess < secret:
        print("太小了！")
    elif guess > secret:
        print("太大了！")
    else:
        print("猜对啦！答案就是",secret)
        print("你一共猜了", count,"次")
        break