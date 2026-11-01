while True:
    text = input("请输入一个数字：")
    try:
        number = int(text)
        print("收到，你输入的数字是", number)
        break
    except:
        print("这好像不是数字，再试一次~")