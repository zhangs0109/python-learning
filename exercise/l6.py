# L6·把输入的一句话追加append写入note.txt
text = input("记点什么：")
f = open("note.txt", "a", encoding="utf-8")
f.write(text + "\n")
f.close()
print("写好了")