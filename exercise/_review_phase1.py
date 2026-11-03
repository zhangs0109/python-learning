# 阶段一练习自动评审脚本：编译检查 + 喂输入跑行为测试
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

def run(fname, feed):
    p = subprocess.run([PY, os.path.join(BASE, fname)], input=feed,
                       capture_output=True, text=True, timeout=10)
    return p.stdout, p.stderr

# 1) 语法编译检查
print("===== ① 语法编译 =====")
for f in ["l1.py", "l2-01.py", "l2-02.py", "l3-01.py", "l3-02.py", "l4-01.py", "l4-02.py", "l5.py"]:
    r = subprocess.run([PY, "-m", "py_compile", os.path.join(BASE, f)], capture_output=True, text=True)
    print(f, "✅ 编译通过" if r.returncode == 0 else f"❌ {r.stderr[:120]}")
print()

# 2) 行为测试
print("===== ② 行为测试 =====")
out, err = run("l1.py", "张双\n光伏项目\n")
print("l1:", out.strip().replace("\n", " | "), err[:100])

out, err = run("l2-01.py", "10\n10.5\n")
print("l2-01:", out.strip().replace("\n", " | "), err[:100])

out, err = run("l2-02.py", "10\n10.5\n")
print("l2-02:", out.strip().replace("\n", " | "), err[:100])

for s in ["95", "75", "45", "20"]:
    out, err = run("l3-01.py", s + "\n")
    print(f"l3-01 输入{s}:", out.strip().replace("\n", " | "), err[:80])

for s in ["6", "0", "-2"]:
    out, err = run("l3-02.py", s + "\n")
    print(f"l3-02 输入{s}:", out.strip().replace("\n", " | "), err[:80])

feed = "\n".join(str(i) for i in range(1, 101)) + "\n"
out, err = run("l4-01.py", feed)
last = out.strip().split("\n")[-2:]
print("l4-01 末行：", " | ".join(last), err[:80])

out, err = run("l4-02.py", feed)
lines = out.strip().split("\n")
prompts = out.count("猜一个数")
print(f"l4-02：共问了 {prompts} 次，最后两行： {' | '.join(last for last in out.strip().split(chr(10))[-3:])}", err[:80])

out, err = run("l5.py", "abc\nxyz\n" + feed)
prompts = out.count("猜一个数")
has_guard = "请输入数字" in out
print(f"l5：防崩提示出现={has_guard}，共问 {prompts} 次，最后两行： {' | '.join(out.strip().split(chr(10))[-3:])}", err[:80])
