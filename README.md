# pharm-lab-report

**药学 / 药物分析实验报告的通用生产线**：把一次实验的课件、讲义、手写记录与仪器谱图，
变成一份可复现的 Word 实验报告和一本全公式的 Excel 数据处理工作簿。

> 约定驱动，而非模板驱动：口径只定义一次，数值只计算一次，排版只由脚本产生一次，
> 最后由只读校验确认三者一致。

这个仓库既是给 AI 助手用的 **Skill**（`SKILL.md` 为入口），
也是一套可以独立使用的方法论与 Python 构建器——不用 AI 也能照着 `references/` 手工做。

---

## 它解决什么问题

做完一个药学实验，通常要交两样东西：

| 交付物 | 常见做法 | 本项目的做法 |
|---|---|---|
| 实验报告 `.docx` | 手打，数字从计算器抄 | 由 `_data.py` 推导，脚本生成，版式固定 |
| 数据处理 `.xlsx` | 另起一个文件重算 | 同一份 `_data.py` 生成，全公式、零错误格 |

**核心矛盾**：手打两份，数字必然对不上；改一次数据，两份都要重来。

所以第一条原则就是：**所有数值只写在一个地方**（`_data.py` 的输入区），
报告与工作簿都由它派生。任何人不许手改产物里的数字。

---

## 覆盖范围

| 类型 | 判据 | 骨架 | 方法 |
|---|---|---|---|
| **限量检查** | 只判"超不超限"，不测绝对含量 | 六章 | 比浊法、比色法、有关物质 |
| **含量测定** | 要给出 mg/片 或占标示量 % | 六章 | 外标一点法、对照品比较法、双波长等吸收点法、两步滴定法 |
| **方法学考察** | 要求做线性/精密度/稳定性/重复性 | 八章 | 五项方法学 + 数据一致性诊断 |
| **工艺/处方筛选** | 正交设计、多因素多水平 | 六章 + 极差分析 | 极差分析 + 方差分析 |

**明确不做**：

- 不生成"没有测过的数据"——缺就留「（待填）」，或去问做实验的人
- 不改数据凑合格——算出 87% 就写 87%，在讨论章给复核路径
- 不处理非实验报告类文档

---

## 安装

```bash
pip install python-docx openpyxl
```

作为 Skill 使用（Claude / WorkBuddy 等支持 Skills 的客户端）：
把整个目录放进用户级 skills 目录，例如 `~/.claude/skills/pharm-lab-report/`
或 `~/.workbuddy/skills/pharm-lab-report/`，入口是 `SKILL.md`。

作为普通工具库使用：把 `scripts/` 加到 `PYTHONPATH`，直接 `import docx_builder` 即可。

---

## 快速开始

```bash
# 1. 建项目骨架
mkdir -p project/{01_source,02_scripts,03_output,04_history,05_verify}

# 2. 复制模板
cp scripts/_data_template.py  project/02_scripts/_data.py
cp scripts/docx_builder.py    project/02_scripts/
cp scripts/xlsx_builder.py    project/02_scripts/
cp scripts/verify_template.py project/02_scripts/verify.py

# 3. 改 _data.py 的「输入区」，先看数字对不对
cd project/02_scripts && python _data.py
```

`_data.py` 自带一组演示数据，**直接跑就能看到完整输出**：

```
--- 线性关系 ---
    4.0 μL  A =     205600  回归     210760  偏差  -2.45%
   ...
  A = 50540·V + 8600   r = 0.99972   r > 0.999 ? 是
--- 含量测定 ---
  C样 = 0.4800 mg/mL → m = 24.00 mg（理论 25.0 mg）
  含量 = 96.00%   每片 28.80 mg   RSD = 0.29%
  判定：93.0%~107.0% → 符合规定
--- 方法学汇总 ---
  合计：5/5 项通过
```

数字对了，再写构建脚本（见 `references/04-workflow.md`），最后 `python verify.py report.docx data.xlsx`。

---

## 目录结构

```
pharm-lab-report/
├── SKILL.md                   Skill 入口：五条原则 + 类型路由 + 七步流程
├── README.md                  本文件
├── references/
│   ├── 01-extraction.md       ★ 从课件/谱图/手写记录提取内容（六问法 + 12 项自检）
│   ├── 02-skeleton.md         章节骨架（六章/八章）+ 两套版式参数 + 表格规则
│   ├── 03-data-analysis.md    ★ 公式库：限量/外标/对照品/双波长/滴定/方法学/方差
│   ├── 04-workflow.md         脚本链路、校验十节、版本流程
│   ├── 05-pitfalls.md         环境 / 数值 / 内容 / 校验 四类坑
│   └── 06-examples.md         四个完整算例（从原始读数到判定）
└── scripts/
    ├── _data_template.py      数值模型骨架（统计工具 + 计算模型 + 自检打印）
    ├── docx_builder.py        Word 构建器：封面/标题/正文/三线表/注/图
    ├── xlsx_builder.py        工作簿构建器：四表骨架 + 自动行高 + 公式防护
    └── verify_template.py     只读校验骨架（十节）
```

---

## 三个构建器提供了什么

### `docx_builder.DocxBuilder`

```python
b = DocxBuilder()            # 紧凑型；DocxBuilder("loose") 为宽松型
b.cover(标题, 副标题);  b.h1(...); b.h2(...); b.p(...); b.pc(...)
b.caption("表 1　……");  b.table(headers, rows, widths=[4,2,2,2]);  b.note("① ……")
b.picture("curve.png", width_cm=12)
b.save("报告.docx")
```

- 两套版式预设：**compact**（10.5 pt / 行距 1.28 / 边距 2.0×1.5 cm）、
  **loose**（12 pt / 行距 1.5 / 边距 2.54×3.17 cm）
- 三线表：顶线 1.5 pt、表头下 0.75 pt、底线 1.5 pt，无竖线、无内部横线，表头跨页重复
- 中文字体按 `w:eastAsia` 正确写入（宋体 / 黑体 / 楷体），西文统一 Times New Roman

### `xlsx_builder.XlsxBuilder`

```python
x = XlsxBuilder(["使用说明", "数据输入", "计算", "公式与依据"])
x.title(...);  x.section(...);  x.note(...)
x.head_row("数据输入", 4, [...]);  x.row("数据输入", 5, [...], fill_input=True)
x.widths("数据输入", {"A": 10, "B": 16})
x.save("数据处理.xlsx")
```

- **自动行高**：按「文本显示宽度 ÷ 列宽」估算行数后显式设高
  （Excel 不会自动撑开合并单元格，这是长句被截断的根因）
- **公式误写防护**：`= 相当量25mg / 50mL` 这类"看起来像公式的说明文字"会被拦下，
  避免打开时满屏 `#NAME?`；真正的公式（`=AVERAGE(数据输入!C4:C8)`）正常放行
- 浅黄底标手填区、浅蓝底标关键结果，打印设置 A4 横向按宽缩放

### `verify_template.py`

十节只读校验，核心是第 ⑥ 节 **数值复算**——把正文里的每个数值与 `_data.py` 独立算一遍比对。
结构检查只能发现"文件坏了"，复算才能发现"数字错了"。

---

## 方法论要点（详见 `references/`）

**提取**：源文件优先级 = 讲义 > 课件 > 手写记录。逐张读图回答六问
（目的 / 原理 / 仪器试剂 / 方法 / 公式与限度 / 思考题）。
凡进入公式的量（取样、称样、定容、稀释、条件、波长、平行份数）一个都不能漏。

**骨架**：验证性实验六章、方法学考察八章。**表格只留给一手数据**，
仪器试剂用文字列举——这是报告批改中最常被要求改的地方。

**算法**：每条公式都写了「口径要点 → 易错点 → 判定」。几个高频陷阱：
限量计算的 μg/g 单位必须统一；外标法的"相当量"不是片重；
两步滴定的 `W̄/W ≡ 1.0000` 可能是取样方案的必然结果；
方差分析无组内重复时必须用合并误差而不是扣减残差。

**如实报**：`n < 2` 时统计量返回 `None`，正文显示「—」，不补编 RSD；
含量异常先做一致性诊断（峰面积比 → 折算浓度 → 应有值），
诊断完仍然照原样报，只多给复核路径。

---

## 适用范围

**适用**：中文药学 / 药物分析类实验报告；片剂、注射剂等制剂的含量与限量测定；
HPLC / UV / 滴定 / 比色比浊等常规方法。

**不适用**：非中文报告；需要复杂统计建模（应用 R/Python 专门工具）；
图表的精细可视化（本项目只做必要的曲线图）。

---

## 许可

本仓库未附带许可证文件。**发布前请自行选择**（MIT / CC BY 4.0 / Apache-2.0 等）。
若希望他人自由复用其中的方法论与脚本，MIT 或 CC BY 4.0 是常见选择。

---

## 贡献

新增方法或踩到新坑，欢迎按现有格式补充：

- 新算法 → `references/03-data-analysis.md`，格式：**公式 → 口径要点 → 易错点 → 判定**
- 新坑 → `references/05-pitfalls.md` 对应分类
- 新骨架/版式 → `references/02-skeleton.md`
- **`SKILL.md` 只维护三样**：核心原则、类型路由、七步流程。细节一律进 `references/`，
  否则加载即爆上下文

提交前请确认示例脚本仍可运行：`python scripts/_data_template.py`。
