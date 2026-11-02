# L5·给猜数字加防崩（输入非数字不崩、重新输入）
import random
answer = random.randint(1,100)
count = 0
while True:
    text = input("猜一个数：")
    try:
        guess = int(text)
    except:
        print("请输入数字！")
        continue

    count = count + 1
    if count > 7:
        print("机会用完，答案是", answer)
        break

    if guess > answer:
        print("大了")
    elif guess < answer:
        print("小了")
    else:
        print("猜对了！用了", count, "次")
        break
