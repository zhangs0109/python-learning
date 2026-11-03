# 阶段二练习自动评审：编译 + 行为 + 数据对拍
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
os.chdir(BASE)          # 保证能读到同目录的 流水.txt / note.txt


def run(fname, feed=""):
    p = subprocess.run([PY, fname], input=feed, capture_output=True, text=True, timeout=10)
    return p.stdout, p.stderr


print("===== ① 语法编译 =====")
for f in ["l6.py", "l7-01.py", "l7-02.py", "l8-01.py", "l8-02.py"]:
    r = subprocess.run([PY, "-m", "py_compile", f], capture_output=True, text=True)
    print(f, "✅" if r.returncode == 0 else f"❌ {r.stderr[:100]}")
print()

print("===== ② 行为测试 =====")
out, err = run("l6.py", "评审测试行\n")
print("l6 追加写入：", out.strip(), "| note.txt 末行：",
      open("note.txt", encoding="utf-8").read().strip().split("\n")[-1])

out, err = run("l7-01.py")
print("l7-01（未加固版）应崩溃：", "ValueError" in err and "✅ 如预期崩溃" or "⚠️ 没崩",
      "|", err.strip().split("\n")[-1][:60] if err else "")

out, err = run("l7-02.py")
lines = [l for l in out.strip().split("\n") if l.strip()]
print("l7-02（加固版）：共输出", len(lines), "行，末行：", lines[-1] if lines else "无")

# 对拍：手工算一遍总额
expected = 25 + 30 + 2 + 6 + 120 + 10
got = [l for l in lines if "合计" in l]
print("   对拍：手工合计 =", expected, "| 程序输出 =", got[0] if got else "无",
      "✅" if got and str(expected) in got[0] else "❌")
print()

out, err = run("l8-01.py")
print("l8-01 字典汇总：", out.strip(), "✅" if "'奶茶': 35" in out.replace('"', "'") else "")
out, err = run("l8-02.py")
print("l8-02 价格异常：", out.strip(), "✅" if "奶茶" in out and "25" in out and "10" in out else "")
