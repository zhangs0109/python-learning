# 把复习手册转成 A4 打印版 HTML（供 Chrome 导出 PDF → 导入 GoodNotes 手写批注）
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))          # python-learning/notes
SRC = os.path.join(BASE, "复习手册_10课跟练.html")
OUT = os.path.join(BASE, "复习手册_打印版.html")

with open(SRC, "r", encoding="utf-8") as f:
    html = f.read()

# 1) 所有折叠答案默认展开（PDF 里点不开，必须展开）
#    注意：标签可能带其他属性（如 data-page-node-id），所以用正则，不能用字符串替换
html = re.sub(r"<details(?![^>]*\bopen\b)([^>]*)>", r"<details open\1>", html)

# 2) 追加打印专用样式
print_css = """
@page { size: A4; margin: 14mm 13mm 16mm 13mm; }
@media print {
  body { background: #fff; font-size: 10.5pt; }
  .wrap { max-width: none; padding: 0; }
  .nav { display: none; }                 /* 跳转目录在纸上没用 */
  h1 { font-size: 20pt; }
  h2 { page-break-before: always; font-size: 14pt; margin-top: 0; }
  h2#map { page-break-before: avoid; }    /* 首页不另起 */
  h3 { font-size: 11.5pt; page-break-after: avoid; }
  .grid { grid-template-columns: 1fr 1fr; gap: 8px; }
  .card, details, table, pre { page-break-inside: avoid; }
  details { border: 0.5pt solid #D8D6CE; padding: 8px 10px; }
  details summary { font-weight: 600; }
  details .ans { border-top: 0.5pt dashed #D8D6CE; padding-top: 6px; margin-top: 6px; }
  pre { font-size: 8.5pt; padding: 8px 10px; }
  table { font-size: 9.5pt; }
  th, td { padding: 4px 6px; }
  .note { page-break-inside: avoid; }
  /* 每题右侧留出批注空间 */
  .ans { background: #FCFCFA; }
}
"""

html = html.replace("</style>", print_css + "\n</style>")

# 3) 标题标注版本
html = html.replace("Python 十课复习手册 · 跟练版（双语 Bilingual）",
                    "Python 十课复习手册 · 跟练版（双语 · 打印批注版）")

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

print("已生成打印版：", OUT)
