# Python 学习练习册

这里存放所有练习代码。按 lesson 编号推进，每个 lesson 目录下的 `.py` 文件都可以用终端运行：

```
cd python-learning/lesson01
python3 hello.py
```

## 配套资料（随学习进度动态更新）

- `notes/学习笔记.md` —— 每课的知识点、踩坑记录、验收记录
- `notes/python语法手册.md` —— 常用语法归档速查（中文备注，标注已学/未学）
- `notes/思维导图.html` —— 可点击的学习地图（浏览器打开，节点带超链接）

更新方式：对 AI 助手说"更新学习笔记"即可。

## 进度记录

- [x] 环境搭建：VS Code + Python 验证通过（2026-10-28）
- [x] Lesson 01：第一个程序 hello.py——已自己修改内容、修复语法错误（缺引号/多逗号）、保存并运行成功（2026-10-28）
- [x] Lesson 02：变量与 input 输入（me.py 自己新建、运行成功；2026-10-28）
- [x] Lesson 02 练习：对比 print 中逗号与加号拼接的区别
- [x] Lesson 03：if / elif / else 判断（quiz.py 三档分支全部验证通过；2026-10-28）
- [x] Lesson 04：while 循环 + random 随机数 → guess.py 猜数字游戏完成，验证通过（2026-10-28）→ **阶段一毕业 🎓**
- [x] Lesson 05：异常处理 try/except（safe.py + guess.py 防弹衣均验收通过；2026-10-29）
- [x] Lesson 06：文件读写（lesson02/note.py 电子日记本验收通过，记录跨运行累积；2026-10-29）
- [x] Lesson 07：记账脚本 note01.py 验收通过——菜单循环/记账/查账/金额防弹/跳过坏数据（2026-10-30）→ **阶段二毕业 🎓**
- [x] L8 加餐：字典 dict（按物品汇总统计验收通过——奶茶 2 笔/合计/平均全对；2026-10-30）
- [x] L9 Git 本地基础：init/add/commit/status/diff/log，两次提交验收通过（2026-11-01）
- [x] L10 GitHub 远程仓库：注册账号、VS Code OAuth 授权、push 上云成功（2026-11-01）
- [x] L10 进阶：`git diff --staged` / `git restore --staged`，修复嵌套仓库与 `add .` 路径陷阱（2026-11-04）→ **阶段三毕业 🎓**
- [x] L11 虚拟环境与 pip：venv 创建/激活/退出、pip 装库、requirements.txt 导出（2026-11-01）
- [x] L12 pandas 入门：read_excel、筛选、groupby 具名聚合、加权平均价、to_excel（2026-11-01）
- [x] 项目一：物资消耗核算与超耗预警 `project/material_check.py`——查重报告 + 同类项汇总 + 加权均价 + 预警标红（2026-11-01）
- [x] 项目二：光伏项目基础数据表分析 `project/pv_analysis.py`——8,868 行、15 类数据体检、11 Sheet + 5 图表 + 网页看板，7 个控制总额全部对拍一致（2026-11-01）
- [x] 项目三：分包结算台账 `project/subcontract_settle.py` v2——4 类体检 + 本期应付/累计/余额 + 平衡表 + 超付预警 + 控制值对拍，5 页签产出（2026-11-04）
  - 新语法：`groupby().cumsum()` 分组内累计（必须先 `sort_values`）
  - 修正：付款进度分母改为「实付 ÷ 计量」，与合同"付80%"直接可比
- [ ] 项目二后续：处理 280 条重复录入、计价单位归一化
- [ ] 项目三验收：用 WPS 核对产出 → 接入真实台账
- [x] 项目三续：资金计划表 `project/cash_plan.py`——付款上限（计量×80%）口径、超付自动冲减、近 3 期均值预测、按币种汇总（2026-11-08）
  - 新语法：`groupby().tail(n)` 取每组最后 n 行
  - 修正：币种以合同台账为准（明细末行币种录错会污染汇总）
- [ ] 下一课：把真实台账接进来，替换样例数据
- [ ] 阶段四：下一个项目（候选：分包结算台账 / 变更索赔台账 / 供应商考核汇总 / 月度环比对比）

## 环境说明

- 编辑器：VS Code（已装 Python 扩展 + Excel Viewer）
- 办公软件：**WPS Office**（本机 `/Applications/wpsoffice.app`）。注意 `.et`/`.wps` 是私有格式，pandas 读不了，需先另存为 `.xlsx`；WPS 导出的 CSV 常为 GBK 编码，读时加 `encoding="gbk"`。
- 虚拟环境：项目根目录 `.venv/`（已写入 `.gitignore`，不提交）
  ```bash
  cd python-learning
  source .venv/bin/activate    # 激活，提示符出现 (.venv)
  pip install -r requirements.txt
  deactivate                   # 退出
  ```
