# Direct Citation Assistant 直接论文引用助手

**[English](README.en.md) | 简体中文**

![GitHub stars](https://img.shields.io/github/stars/xiaolinzi000/direct-citation-assistant?style=flat-square&label=Stars)
![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-blue?style=flat-square)
![GB/T 7714](https://img.shields.io/badge/GB%2FT%207714-2015-green?style=flat-square)
![No Zotero](https://img.shields.io/badge/No%20Zotero-%E2%9C%93-orange?style=flat-square)

> 不依赖 Zotero，纯脚本驱动，在 Word 里完成从「逐句判断引用 → 检索筛选文献 → 插入可点击超链接引用 → 自动生成参考文献目录」的完整论文引用工作流。

适用于**期刊论文、毕业论文（开题报告 / 中期报告）、课程论文**三种场景，支持**中英文文献**混合引用，参考文献格式符合 **GB/T 7714-2015**（默认），另支持 APA / Vancouver / MLA。

---

## ✨ 功能特性

| 能力 | 说明 |
|---|---|
| 🗂️ 自动工作区 | 一键创建项目文件夹（论文 / 引用目录 / 文献 三子目录），自动落在 C 盘之外 |
| 🔍 逐句判断引用 | 明确「这句要不要引」的判断清单，避免堆砌引用 |
| 🎯 选文三原则 | 强证据优先 → 越新越好 → 高引用/高评分，中英文文献都覆盖 |
| 🔗 超链接引用 | 正文编号（默认上标）**点击即跳转**到文末对应文献条目 |
| 📚 分章论文支持 | 一次处理多份 `.docx` 全文统一编号，各章引用**跨文档跳转**到主文档文献表（Word 实测可用） |
| 🔢 自动重编号 | 删除 / 新增引用后重跑脚本，编号自动连续更新，无需手工改 |
| 📑 四种格式 | GB/T 7714（默认）、APA、Vancouver、MLA；编号制 + 作者-年份制 |
| ✒️ 规范斜体 | 西文期刊名/书名按 GB/T 7714 自动斜体、中文刊名正体（`--no-italic-source` 关闭） |
| 🔎 引用真实性校验 | **防止编造引用**：`verify_refs.py` 联网查 Crossref DOI / OpenAlex 题名，判定 ✅ 已核实 / ⚠️ 有出入 / ❌ 未找到，并检查字段完整性与 note 支撑说明 |
| 🔗 引用对应性核对 | `insert_refs.py --mapping` 输出「正文引用 ↔ 文献」对照表（编号、所在句子、note 支撑说明），逐处核对对得上 |
| 🌐 网站检查 | 18 个文献检索网站内置清单，代码自动检查可用性并支持定时更新 |
| 📝 国标格式完整 | 期刊、会议、**专著 [M]、学位论文 [D]、网页 [EB/OL]** 均按 GB/T 7714 输出 |
| 🀄 中文友好 | 占位符支持中文 key（`[CITE:注意力机制]`）、大小写不敏感；中文文献/中文期刊名自动处理 |
| ✒️ 字体版式自动 | 中文宋体 + 英文/数字 Times New Roman；悬挂缩进 2 字符、两端对齐、编号 Tab 对齐 |
| 📄 PDF 归档 | 下载的论文按题名自动匹配重命名归档，重名不覆盖 |
| 💾 自动备份 | 每次修改 docx / refs.csv 前自动备份（保留最近 30 份） |

---

## 📸 效果预览

![演示效果：正文可点击上标引用 + 文末参考文献目录](assets/demo-screenshot.png)

> 上图为真实渲染截图：正文三处上标 `[1][1][2]` 可点击跳转，文末「参考文献」自动生成、悬挂缩进 2 字符、编号 Tab 对齐、条目与正文一一对应。本仓库含演示素材与自动化测试，可自行复现。

---

## 🆕 更新记录

**v1.4.0（最新）**
- 🔎 **引用真实性校验**（新增 `verify_refs.py`）：每条文献联网查 Crossref DOI / OpenAlex 题名，判定 ✅ 已核实 / ⚠️ 信息有出入 / ❌ 未找到（极可能是编造）/ 🔶 无法联网；`--offline` 仅字段检查、`--report` 导出核对报告——**保证参考文献真实存在**
- 🔗 **引用对应性核对**（`insert_refs.py --mapping`）：输出「正文引用 ↔ 文献」对照表（编号、所在句子、note 支撑说明），逐处核对引用与句子对得上；`refs.csv` 的 `note` 字段升级为必填（记录支撑句子）

**v1.3.0**
- 📚 **分章多文档**：一次处理多份 `.docx`，全文引用统一编号，各章引用**跨文档跳转**到主文档文献表（`--main` 指定主文档，Word 实测可用）
- ✒️ **规范斜体**：西文期刊名/书名按 GB/T 7714 自动斜体、中文刊名正体（`--no-italic-source` 关闭）

**v1.2.0**
- 🀄 占位符支持**中文 key** 与大小写不敏感（`[CITE:注意力机制]` / `[cite:key]`）
- 🛡️ 修复参考文献标题后正文段被误删；文档被 Word 占用时给出提示；仅支持 `.docx`
- 🧪 新增 `edge_test.py` 边界回归测试

> 完整版本线（v1.0.0 → v1.4.0）见 [CHANGELOG.md](CHANGELOG.md)。

---

## 🚀 快速开始

### 环境要求

- Windows（脚本基于 Windows 盘符检测与 Word）
- Python 3.8+
- 依赖：`lxml`（必需）、`python-docx`（demo 生成用）、`pypdf`（PDF 识别，可选）、`pypinyin`（中文排序，可选）

```bash
git clone https://github.com/xiaolinzi000/direct-citation-assistant.git
pip install lxml python-docx pypdf pypinyin
```

### 1. 创建工作区

```bash
python scripts/new_project.py --name "毕业论文_开题报告"
```

自动在 C 盘之外创建：

```
毕业论文_开题报告/
├── 论文/        # 放 Word 文稿
├── 引用目录/    # 放 refs.csv 引用信息表
└── 文献/        # 放下载的论文 PDF
```

### 2. 录入引用信息（从 Google Scholar / 知网复制题录）

```bash
python scripts/add_refs.py --refs <项目>/引用目录/refs.csv ^
    --title "Attention is all you need" ^
    --authors "Vaswani, A., Shazeer, N., Parmar, N., et al." ^
    --source "Advances in Neural Information Processing Systems" ^
    --year 2017 --volume 30 --doi 10.5555/3295222.3295349 --type journal
```

### 3. 核实引用真实性与字段完整性（防编造）

```bash
python scripts/verify_refs.py --refs <项目>/引用目录/refs.csv
```

每条文献联网查权威数据库（有 DOI 查 Crossref，无 DOI 查 OpenAlex 题名），判定 `✅ 已核实` / `⚠️ 信息有出入` / `❌ 未找到`（极可能是编造，务必删除或更正后再写入论文）/ `🔶 无法联网`；同时检查题名/作者/年份是否缺失、`note`（支撑的句子）是否填写。`--offline` 只做字段检查，`--report 报告.md` 导出核对报告。

### 4. 正文放占位符，插入引用

在 Word 正文需要引用的句子末尾放 `[?]`（按顺序自动取表内下一条）或 `[CITE:key]`（指定引用某条，大小写不敏感，支持中文 key）：

```bash
python scripts/insert_refs.py --docx <项目>/论文/文稿.docx ^
    --refs <项目>/引用目录/refs.csv --style gbt7714 --citation numbered
```

**分章论文**（毕业论文常用）：一次传入所有章节，全文统一编号，参考文献表自动生成在最后一章末尾，各章正文引用可跨文档跳转：

```bash
python scripts/insert_refs.py --docx <项目>/论文/第1章.docx <项目>/论文/第2章.docx ^
    --refs <项目>/引用目录/refs.csv
```

**对应性核对**：加 `--mapping` 输出「正文引用 ↔ 文献」对照表（每处编号对应的文献、所在句子、note 支撑说明），逐处核对引用是否真实支撑该句：

生成效果：

- 正文引用变为可点击上标 `[1]`（Ctrl+点击跳转到文末对应条目）；
- 文末自动生成「参考文献」标题与目录（没有标题会自动创建）；
- 西文期刊名/书名按 GB/T 7714 自动斜体；
- 删除中间某处引用后再跑一次，其余编号自动重排。

---

## 📋 使用流程（完整版）

```
① 定工作文件夹 → ② 定主题与期刊 → ③ 定引用粒度（句/段/整篇）
→ ④ 逐句检索 + 选文三原则 → ⑤ 收集题录 + 下载论文归档 → ⑥ 插入引用 + 生成目录
```

详细操作指引见 [SKILL.md](SKILL.md)（技能主文件，含每步命令与判定清单）。

### 选文三原则（检索时按优先级）

1. **证据强度**：选能强烈证明该句观点的文献（原始出处 / 直接结论 / 方法原文 / 权威综述）；
2. **时效性**：同强度选最新；经典文献用「原始文献 + 近年综述」组合；
3. **引用量**：同强度同年代优先高被引 / 高评分。

### 引用格式与排版（默认值，均可改）

| 项 | 默认 | 参数 |
|---|---|---|
| 格式体系 | GB/T 7714-2015 | `--style gbt7714\|apa\|vancouver\|mla` |
| 引用样式 | 编号制，上标 `[1]` | `--citation numbered\|author-year` |
| 编号括号 | `[]` | `--bracket ()` |
| 悬挂缩进 | 2 字符（五号 21pt） | `--hanging-pt 21`（小四 24 / 四号 28） |
| 对齐 | 两端对齐 | `--align both\|left` |
| 编号加粗 | 不加粗（通行惯例） | `--bold-num` |
| 中文字体 | 宋体 | `--font-cn "宋体"` |
| 西文字体 | Times New Roman | `--font-en "Times New Roman"` |

---

## 🧩 脚本清单

| 脚本 | 作用 |
|---|---|
| `new_project.py` | 创建项目工作区（C 盘外三子目录 + refs.csv 模板） |
| `add_refs.py` | 把题录写入 refs.csv（参数 / 交互模式，校验 type，自动备份） |
| `rename_papers.py` | 下载的 PDF 按题名匹配重命名归档（重名自动加序号） |
| `insert_refs.py` | **核心**：占位符 → 超链接引用；文末目录生成/更新；自动重编号；支持多文档分章统一编号与跨文档跳转；默认西文期刊名斜体；`--mapping` 输出正文引用↔文献对照表 |
| `verify_refs.py` | **引用真实性校验**：Crossref DOI / OpenAlex 题名联网核实，判定 ✅ 已核实 / ⚠️ 有出入 / ❌ 未找到 / 🔶 无法联网；`--offline` 仅字段检查；`--report` 导出核对报告 |
| `format_refs.py` | 单条题录按 gbt7714 / apa / vancouver / mla 渲染 |
| `refs_db.py` | refs.csv 读写工具库（含 city 字段的国标完整格式） |
| `check_websites.py` | 18 个文献检索网站可用性检查（确定性代码判定，`--update` 写回清单） |
| `demo/*` | 演示素材与自动化测试（`make_demo.py`、`renumber_test.py`、`author_year_test.py`、`edge_test.py`、`multi_doc_test.py`、`verify_docx.py`） |

### 文献检索网站清单与定时检查

- 完整清单：`references/websites.md`（网站 / URL / 用途 / 状态 / 备注，含中英文与各领域渠道）；
- 状态由 `check_websites.py` 自动维护（✅ 可达 / ⚠️ 反爬 / ❌ 不可达），可配置定时任务每周检查更新；
- 增删网站只需改 `check_websites.py` 内的 `SITES` 列表后重跑。

---

## 📂 目录结构

```
direct-citation-assistant/
├── SKILL.md                 # 技能主文件（完整工作流 + 判定清单 + 参数说明）
├── README.md                # 项目介绍（中文）
├── README.en.md             # 项目介绍（英文）
├── CHANGELOG.md             # 版本升级说明
├── assets/
│   └── demo-screenshot.png  # 演示效果截图
├── references/
│   └── websites.md          # 文献检索网站清单（状态由脚本维护）
├── scripts/
│   ├── new_project.py       # 创建工作区
│   ├── add_refs.py          # 题录录入
│   ├── rename_papers.py     # PDF 按题名归档
│   ├── insert_refs.py       # 核心：插入引用 + 目录 + 重编号
│   ├── format_refs.py       # 4 种格式渲染
│   ├── refs_db.py           # refs.csv 读写
│   ├── check_websites.py    # 网站可用性检查
│   └── demo/                # 演示文稿 + 自动化回归测试
└── .gitignore
```

---

## ✅ 自动化测试

```bash
cd scripts/demo
python make_demo.py          # 生成演示素材（refs.csv + 文稿.docx）
python renumber_test.py      # 场景：删除中间引用 / 新增引用 → 编号自动重排
python author_year_test.py   # 场景：作者-年份制 → 字母序 + 幂等
python edge_test.py          # 边界：中文 key / 小写占位符 / 无年份 / 正文段保护
python multi_doc_test.py     # 场景：分章多文档 → 全文统一编号 + 跨文档跳转
python verify_refs_test.py   # 离线回归：字段完整性 / --offline / 相似度函数
python verify_docx.py 文稿.docx   # 结构完整性（超链接 ↔ 书签一一对应）
```

> 真实性校验的联网判定（Crossref/OpenAlex）依赖网络，可人工抽查 demo：`python ../verify_refs.py --refs refs.csv`——真实文献（vaswani2017/devlin2019）应判 ✅，虚构演示数据（zhang2023）应判 ❌ 未找到。

---

## ⚠️ 注意事项

- **文档被 Word 占用**：运行脚本前先关闭打开的 Word 文档（脚本会检测并提示）；
- **仅支持 .docx**：.doc 请先在 Word 中另存为 .docx；
- **`[?]` 多于未引用条目**：脚本会报错，先用 `add_refs.py` 补文献；
- **同一篇引多次**：第二次放 `[CITE:同一个key]`，编号自动复用；同一处引多篇放 `[CITE:k1][CITE:k2]`；
- **专著 / 学位论文 / 会议**：记得在 refs.csv 填 `city`（出版地/保存地），`web` 类型必须填 `url`；
- **正文标注位置**：编号用上标、紧跟引文内容、放在句末标点之前（`……显著进展[1]。`），前后无空格；
- **行距**：国标不强制，随学校/期刊模板，可在 Word 中全选参考文献段统一设置。

---

## 📄 说明

本仓库为个人技能项目，当前未指定开源许可证，欢迎参考学习；如需在论文/产品中使用或二次分发，请先联系作者确认。
