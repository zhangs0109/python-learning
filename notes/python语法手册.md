# Python 常用语法手册（中文备注版）

> 📌 按主题归档的语法速查手册，示例全部来自我们自己写过的代码。
> 标记说明：✅ 已学 ｜ 🔵 学习中 ｜ 🔜 后面会学
> **最后更新**：2026-11-04

---

## 目录

1. [输出与输入](#一输出与输入-)
2. [变量与数据类型](#二变量与数据类型-)
3. [字符串操作](#三字符串操作-)
4. [条件判断](#四条件判断-)
5. [循环](#五循环-)
6. [随机数模块](#六随机数模块-)
7. [异常处理](#七异常处理-)
8. [文件读写](#八文件读写-)
9. [常见报错速查](#九常见报错速查-)
10. [编程好习惯](#十编程好习惯-)
11. [字典与按物品汇总](#十一字典与按物品汇总-)
12. [虚拟环境与 pip](#十二虚拟环境与-pip-)
13. [pandas 数据分析](#十三pandas-数据分析-)
14. [Git 命令速查](#十四git-命令速查-)

---

## 一、输出与输入 ✅

```python
# 打印（输出）到屏幕
print("你好，世界！")        # 打印文字，文字要加英文引号
print(2026)                 # 打印数字，不需要引号
print("数字是", 2026)       # 逗号隔开 → 输出时自动加空格
print("你好，" + "世界")    # 加号连接 → 紧紧相连，不加空格

# 输入（让用户打字）
name = input("你叫什么名字？ ")
# 注意：input 拿到的永远是"文字"（字符串），哪怕是数字
# 提示语结尾留个空格，显示更美观
```

## 二、变量与数据类型 ✅

```python
# 变量 = 贴了标签的盒子。等号是"装进去"，不是"等于"
name = "张双"       # 文字类型：str（字符串），要加引号
age = 26            # 整数类型：int
price = 3.5         # 小数类型：float

# 类型转换
number = int("42")      # 文字 → 整数（"42" → 42）
text = str(42)          # 数字 → 文字（42 → "42"）
```

**易错点：**
- `=` 是赋值（装盒子），`==` 才是判断相等
- `int("你好")` 会崩溃：文字转不成数字（→ 用 try/except 解决，见第七节）

## 三、字符串操作 ✅

```python
# 拼接：加号只能连接文字和文字
"你好，" + name + "！"        # ✔ 正确
"你今年" + age + "岁"          # ✘ 报错！age 是数字，要先 str(age)

# f-string（推荐写法，更清晰，后面会练）
f"你今年 {age} 岁"             # 把变量直接放进大括号
```

## 四、条件判断 ✅

```python
answer = input("1 + 1 等于几？ ")

if answer == "2":            # 注意：冒号结尾
    print("答对啦！")         # 注意：缩进 4 格，表示"属于 if"
elif answer == "11":         # "否则如果"，可以有多个
    print("方向对了一半")
else:                        # 以上都不成立时
    print("再想想～")
# 程序从上往下检查，谁的条件成立走谁的路，之后不再检查
```

**常用比较符号：**

| 符号 | 含义 |
|---|---|
| `==` | 等于（判断） |
| `!=` | 不等于 |
| `>` `<` | 大于、小于 |
| `>=` `<=` | 大于等于、小于等于 |

**易错点：**
- 缩进错误（IndentationError）：冒号下一行忘了缩进 4 格
- 用了 `=` 而不是 `==`

## 五、循环 ✅

```python
# while：条件成立就一直重复缩进里的代码
count = 1
while count <= 5:          # 循环条件
    print("第", count, "遍")
    count = count + 1      # 计数器加 1，否则永远循环停不下来
print("数完啦！")           # 不缩进 = 不属于循环，循环结束后执行

# while True + break：无限循环，满足条件时跳出
while True:
    text = input("请输入： ")
    if text == "退出":
        break              # 立刻跳出整个循环
```

**易错点：** 忘记写 `count = count + 1`，程序会永远停不下来（Ctrl+C 可以强制停止）

## 六、随机数模块 ✅

```python
import random               # 先把 random（随机）模块请进来

secret = random.randint(1, 100)   # 随机抽一个 1~100 的整数（含两端）
```

## 七、异常处理 ✅

```python
# try / except：试着做，炸了就执行补救方案，程序不崩溃
try:
    number = int(text)          # 可能出错的代码放这里
    print("收到", number)
except:
    print("这好像不是数字，再试一次～")   # 出错后执行这里
    continue                    # 常用搭配：跳回循环开头重新来
```

**两个重要教训（实战验证）：**
- try 和 except 必须成对出现，漏一个就报语法错误
- 接住错误后要用 `continue`（或重新赋值）让流程回到安全轨道，否则后面的代码可能用到没赋值的变量 → NameError

## 八、文件读写 ✅

```python
# 写文件三步走
f = open("账本.txt", "a", encoding="utf-8")   # "a"=追加不清空；"w"=清空重写（小心）
f.write("奶茶 25\n")                           # \n 表示换行
f.close()                                      # 用完关门

# 读文件
f = open("账本.txt", "r", encoding="utf-8")   # "r"=只读
lines = f.readlines()                          # 所有行读进一个列表
f.close()

for line in lines:                             # for：列表里有几个就做几次
    print("-", line.strip())                   # strip() 去掉行尾的换行符
```

**字符串拆分（记账脚本要用）：**
```python
parts = "奶茶 25\n".split()    # 按空格拆开，得到 ["奶茶", "25"]
name = parts[0]                # "奶茶"（列表第 0 个）
money = int(parts[1])          # 25
```

## 九、常见报错速查 ✅

| 报错信息 | 中文含义 | 常见原因 |
|---|---|---|
| `SyntaxError` | 语法错误 | 漏了引号/冒号/括号，或用了中文符号 |
| `IndentationError` | 缩进错误 | 冒号后忘了缩进 4 格 |
| `NameError` | 名字不存在 | 变量名拼错，或还没赋值就使用 |
| `ValueError` | 值不对 | `int("你好")` 文字转不成数字 |
| `can't open file ... No such file or directory` | 文件不存在 | 运行时文件名打错了 |
| `TypeError` | 类型不对 | 数字和文字直接用加号拼接 |
| `FileNotFoundError` | 找不到数据文件 | 数据文件不在"当前工作目录"，或文件名/路径写错 |
| `ModuleNotFoundError` | 找不到这个库 | 没装库，或忘了激活虚拟环境（少了 `(.venv)` 前缀） |
| `No configured push destination` | 没配置推送目标 | 这个仓库没有远程（多半在子目录里误跑了 `git init`） |

**看报错的秘诀：看最后几行**——错误类型 + 出错的行号。

## 十、编程好习惯 ✅

1. 改完代码 **Command + S** 保存再运行（不保存 = 跑旧代码）
2. 亲手敲代码，不要复制粘贴
3. 报错先自己读一遍，再求助
4. 代码里加 `#` 注释，给三个月后的自己看
5. 一次只改一处，改完就跑，方便定位问题
6. 用 `BASE = os.path.dirname(os.path.abspath(__file__))` 拼绝对路径，脚本放哪都能跑
7. 数字能加的前提是**单位相同**；外币要先折算
8. 工具不要替人做业务判断：发现异常先报告，改不改由人定

## 十一、字典与按物品汇总 ✅

```python
# 字典 = 带标签的柜子，按"名字"（键）存取
money_sum = {}                          # 空字典
money_sum["奶茶"] = 25                  # 新建一格
money_sum["奶茶"] = money_sum["奶茶"] + 18   # 往格子里累加 → 43
print(money_sum["奶茶"])                # 按名字取：43

if "电影" in money_sum:                 # 判断柜子里有没有这个格子
    print("有电影的记录")

for name in money_sum:                  # 遍历：挨个取出每个名字
    print(name, money_sum[name])

round(21.6666, 1)                       # 四舍五入到 1 位小数 → 21.7
```

**经典应用模式：按物品归堆求和（物资核算的雏形）**
```python
money_sum = {}      # 名字 → 总金额
counts = {}         # 名字 → 笔数
for line in lines:
    parts = line.split()
    name = parts[0]
    money = int(parts[1])
    if name in money_sum:               # 见过这个物品 → 累加
        money_sum[name] = money_sum[name] + money
        counts[name] = counts[name] + 1
    else:                               # 第一次见 → 新建
        money_sum[name] = money
        counts[name] = 1

for name in money_sum:                  # 汇总报告
    print(name, "共", counts[name], "笔，合计", money_sum[name],
          "元，平均", round(money_sum[name] / counts[name], 1), "元")
```

**易错点：**
- 键名拼错：`mongey_sum` ≠ `money_sum`，报 "is not defined"
- **过期变量**：循环结束后，循环里的变量还保留着最后一轮的值——别在错误的循环里用它做计算
- 分清"循环内"（每轮都做）和"循环外"（全部做完做一次）：缩进说了算

## 十二、虚拟环境与 pip ✅

```bash
# 在项目根目录执行
python3 -m venv .venv            # 创建（create）虚拟环境
source .venv/bin/activate        # 激活（activate）→ 提示符前出现 (.venv)
pip install pandas openpyxl      # 安装（install）第三方库
pip list                         # 看装了什么
pip freeze > requirements.txt    # 导出依赖清单
pip install -r requirements.txt  # 别人复原环境
deactivate                       # 退出（deactivate）
```

**关键认知：**
- 提示符前有 `(.venv)` = 在房间里装库；没有 = 装到系统里（危险）
- 虚拟环境与**当前目录无关**，`cd` 到哪都还在房间里
- `.venv/` 不提交 Git（体积大、与操作系统绑定），写进 `.gitignore`
- 报 `ModuleNotFoundError` 先查两件事：装了吗？激活了吗？

## 十三、pandas 数据分析 ✅

```python
import pandas as pd

# 读 Excel（WPS 的 .et/.wps 读不了，先另存为 .xlsx）
df = pd.read_excel("台账.xlsx", sheet_name="明细", header=2)  # 表头在第 3 行 → header=2
df = pd.read_csv("数据.csv", encoding="gbk")                  # WPS 导出的 CSV 常是 GBK

# 打量一下数据（每次读进来都先做这四件事）
df.head()            # 前 5 行
df.columns           # 列名
len(df)              # 行数
df.shape             # (行数, 列数)
df.dtypes            # 每列的类型

# 取数与统计
df["实耗量"].sum()      # 求和
df["实耗量"].max()      # 最大
df["实耗量"].idxmax()   # 最大值所在行号
df["实耗量"].mean()     # 平均

# 筛选
over = df[df["超耗量"] > 0]                      # 只要满足条件的行
df.loc[df["超耗量"] < 0, "超耗金额"] = 0          # 条件 + 列 → 批量赋值

# 排序
df = df.sort_values("实耗金额", ascending=False)   # 按金额降序

# 查重（不是删重！）
dup = df[df.duplicated(keep=False)]               # keep=False 连第一笔一起标出来

# 分组汇总（字典的工业版）
g = df.groupby(["材料", "部位", "单位"], as_index=False).agg(
    笔数=("实耗量", "count"),        # 具名聚合：结果列名 = (源列, 算法)
    数量=("实耗量", "sum"),
    金额=("实耗金额", "sum"),
)
g["单价"] = (g["金额"] / g["数量"]).round(2)       # 加权平均价 = 总金额 ÷ 总数量

# 清洗
df["数量"] = pd.to_numeric(df["数量"], errors="coerce")  # 转不了的变 NaN
df = df.dropna(subset=["数量"])                          # 丢掉 NaN 行
df["单位"] = df["单位"].replace({"Kg": "KG", "kg": "KG"}) # 单位归一

# 导出
df.to_excel("结果.xlsx", index=False)   # index=False：别把行号写成一列
```

**三条铁律：**
1. **`header` 数错，整张表错位**——读进来先 `head()` 看一眼
2. **加权平均 ≠ 单价平均**：成本核算永远用 `总金额 ÷ 总数量`
3. **不是所有数字都能加**：KG 和台班、人民币和美元，先归一再汇总

## 十四、Git 命令速查 ✅

```bash
# 一次性配置
git config --global user.name "张双"
git config --global user.email "zhangs0109@163.com"
git config --global core.quotepath false    # 中文文件名正常显示

# 日常三命令
git add -A                      # 全部暂存（-A 覆盖整个仓库，比 . 稳）
git commit -m "做了什么"        # 本地存档
git push                        # 同步到 GitHub

# 查看
git status                      # 颜色看状态
git diff                        # 还没 add 的改动
git diff --staged               # 已经 add 的改动
git log --oneline               # 历史（新的在上）
git rev-parse --show-toplevel   # 我到底在哪个仓库里？（排查嵌套仓库的神器）

# 撤销
git restore --staged 文件名     # 从暂存区撤回（改动保留，不丢代码）
git restore 文件名              # 丢弃工作区的改动（⚠️ 真丢，慎用）

# 远程
git remote -v                   # 看远程地址
git remote add origin <url>     # 绑定远程
git push -u origin main         # 首次推送并记住默认目标
```

**规矩：没有输出 = 一切正常**，只有报错才说话。

**两个高频坑：**
- 在子目录误跑 `git init` → 嵌套仓库 → push 报"没有推送目标"。查：`git rev-parse --show-toplevel`；治：删掉子目录的 `.git`
- `git add .` 只管当前目录以下；改名过大小写的文件要 `git add -A` 才认
