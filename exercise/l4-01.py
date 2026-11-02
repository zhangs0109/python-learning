# L4-01·猜数字：随机1-100，循环数，提示大小，统计次数
import random
answer = random.randint(1,100)
count = 0
while True:
    guess = int(input("猜一个数："))
    count = count + 1
    if guess > answer:
        print("大了")
    elif guess < answer:
        print("小了")
    else:
        print("猜对了！用了", count, "次")
        break

