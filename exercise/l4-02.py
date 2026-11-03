# L4-02·最多猜7次，猜完没中就公布答案
import random
answer = random.randint(1,100)
count = 0
while True:
    if count > 7:
        print("机会用完，答案是", answer)
        break

    guess = int(input("猜一个数："))
    count = count + 1

    if guess > answer:
        print("大了")
    elif guess < answer:
        print("小了")
    else:
        print("猜对了！用了", count, "次")
        break

