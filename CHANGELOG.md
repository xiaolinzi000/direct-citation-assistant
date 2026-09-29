# 更新日志（Changelog）

本项目所有重要变更记录于此。版本线按功能里程碑划分，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [v1.9.0] - 2026-09-29

### 主题：验收门槛硬化——把"建议人工抽查"改成"未通过检查就报告失败"

用户评审结论：v1.8 已实现大部分功能与检查要求，但多为"建议人工抽查"；本次把每类要求都改成**未通过即失败**（默认退出码 1，`--warn-only` 才降级为警告）。

**1. 逐句证据核验记录（evidence 字段）**
- `refs_db.py`：`REFS_HEADER` 扩为 14 列，新增 `evidence` 列（`key,type,…,note,evidence`）；docstring 规定建议格式「论断 | 出处（段/页/图） | 核验深度（全文/摘要） | 支持程度（强/弱/待核实）」，**无法确认标「待核实」，不得自动挪引文**
- `verify_support.py --evidence-csv`：核验后把每条 key 的结论写回 refs.csv 的 evidence 列（✅强支持→强、⚠️弱支持→弱、无摘要/跨语言/表中无 key→待核实）；写入前自动备份；输出「待核实」清单提醒人工补正
- `add_refs.py` 交互模式新增 evidence 录入；`insert_refs.py --mapping` 展示 evidence 列

**2. Word 链接完整性检查（逐条验收，未通过即失败）**
- `insert_refs.py` 新增 `verify_docx_links()`：逐个确认 ①正文引用超链接（anchor=ref_key）在文末都有对应书签（无悬空链接）②文末每个书签都被正文引用（无孤立条目）③编号制下正文编号集合与文末条目编号集合一致（无孤立编号/悬空编号）④编号↔条目一致（key_to_num 核对，防张冠李戴）⑤无重复书签
- 第 9 步回读验证升级：编号不一致/书签重复/链接不完整全部进 fail 列表，**未通过默认 `sys.exit(1)`**（--warn-only 仅警告）
- 新增 `scripts/verify_links.py`：独立验收入口——**AI/他人改完文档后必跑**（`python verify_links.py --docx 文稿.docx`，多文档 `--main` 指定主文档），未通过退出码 1

**3. DOI 保留与检查（精确解析，不静默删、不猜填）**
- `format_refs.py` 新增 `DOI_RE` + `normalize_doi()`：兼容裸 DOI / `doi:` / `https://doi.org/` / `dx.doi.org` / `DOIs:` 等录入形态，自动剥离尾随标点；提取不到返回 None（视为无可信 DOI）
- gbt7714/apa/vancouver/mla/harvard 五处 DOI 拼接全部改走 `normalize_doi`（修复录入值自带 `https://doi.org/` 前缀时重复前缀的隐患）
- `insert_refs.py`：文末表生成后输出**缺 DOI / DOI 可疑清单**（明确报告、提示 fetch_doi 补查，不静默删）；新增 `--strict-doi` 未通过直接失败
- `verify_refs.py`：DOI 录入值无法精确解析时显式报"无法解析——不猜填"，校验转 OpenAlex 题名兜底

**4. 按目标期刊建立格式配置（IEEE/TIM）**
- `format_refs.py` 新增完整 **IEEE 渲染器**（IEEE Reference Guide）：
  - `IEEE_JOURNAL_ABBREV` 官方缩写表（TIM→`IEEE Trans. Instrum. Meas.`、TII、TPEL、TSP、Nature、Science、Proc. IEEE 等 17 项）
  - `IEEE_MONTHS` 月份缩写（Jan./Feb./…/Dec.）；`_ieee_month()` 解析 `2025-02`/`Feb 2025`/`2025`，无 month 时回退年份
  - `_ieee_authors()`：名缩写在前姓在后、最多列 6 人、第 7 人起 `et al.`；修复作者段在期刊/会议/学位论文/网页分支漏接的问题
  - 期刊条目 `J. K. Author, "Title," Abbrev. J., vol. x, no. x, pp. xxx-xxx, month, year, doi: …`；会议/专著/学位论文/网页分类型
  - `FORMATS` 现支持 6 种样式：gbt7714/apa/vancouver/mla/harvard/**ieee**
- `insert_refs.py --style ieee`：正文编号**不上标**（IEEE 行内方括号 [1]，`--superscript` 可覆盖）；`--style` choices 加入 ieee

**5. 最终 Word 页面检查列为必做项**
- `SKILL.md` 步骤 5c 明确：文档经引用/斜体/排版修改后**必须用 Word 打开做页面级检查**（点击跳转、条目格式、DOI 链接、缩进/分页），**不能只依赖 DOCX 内部属性或脚本报告**
- `verify_links.py` 作为自动化回读手段（AI 改稿后必跑），页面检查作为最终人工必做项

**6. 配套兼容性修复**
- `demo/refs.csv` 表头升级为 14 列（含 city/evidence）；`refs_db.load_refs` 检测旧表头并显式告警（防止旧表头 + 新行导致列错位）

### 测试与工程
- 新增 `demo/v19_test.py` 回归测试（A. normalize_doi 7 断言 / B. IEEE 渲染多类型 / C. insert --style ieee 端到端含不上标+DOI 链接 / D. verify_docx_links 悬空/孤立条目/孤立编号检出 / E. verify_links.py 独立验收 exit0/exit≠0）
- 全量回归 12/12 全绿：v19 + 既有 11 项（author_year / edge / freeze / merge / multi_doc / renumber / sentence_level / special_chars / toc / verify_refs / year_suffix）
- 修复 IEEE 渲染器作者段漏接 bug（v1.9 开发中发现并修复，新增断言覆盖）

## [v1.8.0] - 2026-09-29

### 新增（按用户实测反馈的 8 项改进）

**1. 句级落点（最高优先级）：引用从"段末堆叠"改为"逐句紧跟"**
- `insert_refs.py` 新增 `split_sentences` 句级切分器（中文/英文句末标点切分，保守规避：小数点/版本号、常见缩写点 e.g./et al./fig.、连续点省略号、切分点后闭合引号归前句）
- 引用点提取按句切分，每个编号紧跟其支撑的那句话；无法确定对应关系时宁可不放、不放段末打包
- **句内落点校正**：占位符若误放在句号后（「句子。[CITE]」），自动把句末标点剥离到编号之后——编号紧跟引文、置于句末标点之前（「句子[1]。」）
- **堆叠告警**：插入后自动扫描单句（默认 ≥3 个）与单段（默认 ≥5 个）堆叠编号并提示拆分；`--stack-sentence N` / `--stack-paragraph N` 调阈值、`--no-stack-warning` 关闭
- `--mapping` 输出升级为句级定位（段号+句号）

**2. 编号冻结 + 增量更新（key 恒定，如 collet2020 恒为 1）**
- 新增 `--freeze`：编号以 refs.csv 的 key 为准，已冻结 key 恒为原编号；新 key 复用"正文实际占用之外"的最小编号；已从正文消失的旧 key 编号自动释放
- 编号状态落盘 `<主文档同目录>/_refs_state/<主文档名>.json`，每次运行自动输出**新旧编号映射 diff**（新增/移除/变化/不变），验收有据
- `--no-reuse-gaps` 关闭空号复用（编号只增不减）；`--no-state` / `--state <路径>` 控制状态文件

**3. 文末定位修复（TOC 误插坑）**
- 定位改为取**最后一个** References/参考文献 标题；自动检测并跳过 TOC 区域（fldSimple TOC、instrText=TOC 覆盖的全部目录内容行）
- to_placeholders 模式与多文档合并同样过滤 TOC

**4. DOI 缺失与限流回退链**
- 新增 `scripts/api_client.py` 公共模块：429/500/502/503/504 指数退避重试（1.5×2ⁿ+抖动、上限 20s）；API 响应落盘缓存 `scripts/_api_cache/`（默认 7 天 TTL、sha1 键名）
- 新增 `scripts/fetch_doi.py`：回退链 **Crossref → PubMed E-utilities → OpenAlex → 出版页抓取**（`--scrape URL`），仅"题名+年份+首作者"全吻合才建议 DOI；`--apply` 写回前自动备份 refs.csv；`--report` 导出补查报告
- `verify_refs.py` 联网走 api_client，新增 `--refresh-cache` 绕缓存强制联网

**5. 内容级核验（防张冠李戴）**
- 新增 `scripts/verify_support.py`：从 docx 提取引用句（复用句级切分器）→ 拉文献摘要（Crossref JATS → OpenAlex abstract_inverted_index）→ token 覆盖率判定 ✅强支持（≥0.25）/ ⚠️弱支持 / 🔶无摘要
- 输出"弱支持引用"清单，人工复核有据可查；`--min-coverage` 调阈值、`--report` 导出 Markdown
- **跨语言检测**：中文引用句对英文摘要等自动标注「🌐 跨语言（人工核对）」，不误报弱支持

**6. 变更审计报告**
- `insert_refs.py --audit <路径>.md` 输出 5 节：编号映射 diff / 句级引用点清单 / 堆叠告警 / 文末条目 / 备份与回滚清单，配合自动备份一键回滚

**7. 样式扩展与特殊字符回归**
- `format_refs.py` 新增 **Harvard（作者-年份制）** 输出（期刊/专著/学位论文/网页分支，西文刊名斜体）
- 新增 `demo/special_chars_test.py`：Künzler、Dall'Angelo、O'Connor、Müller、für 等重音/撇号贯穿 format→insert 全链路 + XML 回读断言，特殊字符转义固化为回归测试

**8. 多章节合并编号**
- 新增 `scripts/merge_refs.py`：读入各章 → 移除各自文末表（复用 TOC 过滤）→ 合并正文 → 移除跨文档 r:id → 书签 ID 全局重排 → **全局重编号** → 文末一份表 → 输出「章节旧编号→全局新编号」映射表；备份 + 回读

### 测试与工程
- 新增 5 个回归测试（全部通过）：`sentence_level_test.py`（切分器 6 组断言+落点校正+堆叠告警）、`freeze_test.py`（冻结+幂等）、`toc_test.py`、`merge_test.py`、`special_chars_test.py`；新增 `demo/_test_util.py` 最小 docx 构造工具（手写 XML，不走 python-docx）
- 修复既有误报：`verify_docx.py` 过滤 anchor=None 的 DOI 外部链接后再比对书签
- **联网实测（2026-09-29）**：verify_refs 三条判定符合预期（devlin2019 ✅ Crossref、vaswani2017 ✅ OpenAlex 兜底、虚构 zhang2023 ❌）；verify_support 成功拉取 OpenAlex 摘要并正确标注跨语言；fetch_doi 对 BERT 命中 Crossref DOI（10.18653/v1/N19-1423），Attention 2017（10.5555 前缀 ACM 早期 DOI、未在 Crossref 注册）正确判未命中
- 全部旧测试保持全绿：verify_refs_test（14 项）/ renumber_test / author_year_test / year_suffix_test / edge_test / multi_doc_test



### 新增：一键批处理（无需记命令）
- **`转占位符草稿.bat`**：把 Word 文档拖到文件上（或双击后输入路径）→ 自动跑 `insert_refs.py --to-placeholders`，生成"纯文本草稿"交给 AI/他人修改
- **`恢复引用.bat`**：把 AI 改完的文档拖到文件上 → 自动找引用表（优先文档同目录/上级「引用目录」的 refs.csv，找不到再手动输入）→ 恢复超链接、编号、文末表，并自动带 `--mapping` 输出引用↔文献对照表
- 已实测：生成 3 处引用 → bat 转占位符（3 转换、文末 2 条移除）→ python-docx 模拟 AI 改写 → bat 恢复（3 处超链接、编号连续、文末 2 条）✓
- bat 按 GBK 编码 + `chcp 936` 生成，中文提示在 cmd 下无乱码（UTF-8 无 BOM 的 bat 会被 cmd 解析错乱，已规避）

## [v1.7.0] - 2026-09-10

### 新增：`--to-placeholders`（防"AI 改论文丢超链接"）
- **痛点**：AI 修改 .docx 时重建段落/文档，引用超链接和书签丢失、只剩纯文本 `[1]`，且重跑脚本也认不出纯文本编号（脚本只识别超链接与占位符），导致文末表被清空重建、正文编号错乱
- **方案**：修改前跑 `insert_refs.py --docx 文稿.docx --to-placeholders`——所有引用超链接转回 `[CITE:key]` 普通文本（含 key 信息，任何编辑破坏不了）、移除文末参考文献表，生成"纯文本草稿"交给 AI/他人自由改正文；对方改完后正常重跑（带 `--refs`）即恢复超链接、编号自动重排、文末表重建
- **实测闭环**：make_demo → 生成引用（3 处）→ `--to-placeholders`（3 超链接→占位符、文末 2 条移除）→ python-docx 改写正文（模拟 AI，占位符原样保留）→ 重跑 → 3 处超链接恢复、编号 1/2 连续、文末 2 条、AI 修改文字保留 ✓
- SKILL.md 新增「**5g 交给 AI/他人修改论文：防丢超链接工作流**」；README 中英更新记录与脚本表

## [v1.6.1] - 2026-09-10

### 文档
- SKILL.md 新增「**5e 增删引用后的维护（脚本方式）**」：删除某处引用 → 重跑自动重排；**老师/导师指定文献**三步法（add_refs 录入 → `[CITE:key]` 精确指定 → 重跑自动插入编号并后移）；从 refs.csv 删文献的顺序（先删正文引用再删行）
- SKILL.md 新增「**5f 手动编辑引用（仅定稿场景）与风险**」：手动添加/删除的具体做法 + 编号需手动 +1/-1；明确警告——手动纯文本编号脚本不识别、重跑会被覆盖、手写条目可能被误删；结论：论文还要改就用脚本方式

## [v1.6.0] - 2026-09-10

### 新增
- **元数据核对升级**：Crossref DOI 精确命中时，除题名外新增**年份 / 卷 / 期 / 页**与权威库比对（归一化处理 en-dash/空格/尾部标点），任何一项不符自动降级 `⚠️ 信息有出入` 并注明差异——防"真实文献但卷期页录错"（网上引用常见错误之一）
- SKILL.md 步骤 4.5 补充说明：
  - OpenAlex 题名兜底**只核对题名、不核对年份**（题名搜索可能命中不同版本/预印本，如 Attention Is All You Need 命中 2025 重印版，此前年份核对导致误报，已修复）
  - **中文文献覆盖提示**：OpenAlex 对中文题名覆盖有限，真实中文文献可能被判 ❌ ——用知网/万方/维普人工确认即可；❌ 真正要警惕的是"知网也搜不到"的条目

### 修复
- OpenAlex 兜底年份核对误报（2017 论文命中 2025 重印版被判 ⚠️）→ 年份核对仅限 Crossref DOI 精确命中路径

### 其他
- `verify_refs_test.py` 新增卷/期/页归一化断言（14 项）
- 联网实测：年份录错 2020 → ⚠️ year不符（库 2019 vs 录入 2020）✓；正常 demo 恢复 ✅2+❌1（无 ⚠️ 误报）✓

## [v1.5.0] - 2026-09-10

### 新增
- **文末条目 DOI 可点击跳转论文页面**：`document.xml.rels` 建立 External 超链接（Target=`https://doi.org/…`），条目中的 DOI 片段渲染为蓝色可点击超链接；已有同 Target 关系复用（重跑幂等，实测不新增）；Word 实测 2 条 DOI 均可点击直达论文
- **GB 格式 DOI 著录规范为 `https://doi.org/…`** 形式（此前按原文输出，如 `10.5555/…` 不可点击）；APA/MLA/Vancouver 同步支持可点击
- **著者-出版年制同年同作者 a/b 后缀**：同一（首作者, 年份）多篇文献按题名字母序自动加 2023a/2023b，正文标注 `(张伟等, 2023a)` 与文末条目年份后缀一致（GB/T 7714 规范，此前缺失）
- `verify_refs.py` **note 质量检查**：note 过短（<6 字）提示写明具体支撑句子/观点
- SKILL.md 新增「**5d 最高规格科研规范自查清单**」：真实性 / 支撑性 / 无编造 / 无多余遗漏 / 格式准确 / 正文标注 / 排序规范 / 时效性 / 可追溯 9 项，交稿前逐项核对（其中支撑性、格式、时效、可追溯为人工核对环节）

### 修复
- `author_year_test.py`：文末条目新增 DOI 超链接后，测试误把 DOI 外部链接当作正文引用收集 → 只收集带 anchor 的正文引用

### 其他
- 新增 `demo/year_suffix_test.py`（同年同作者后缀：正文 a/b、文末年份后缀、无后缀不误加、幂等）
- README 中英文：特性表（DOI 可点击、a/b 后缀）、更新记录 v1.5.0

## [v1.4.0] - 2026-09-10

### 新增
- **引用真实性校验 `verify_refs.py`**（防编造，硬性要求）：每条文献联网查权威数据库——有 DOI 查 Crossref API 核对题名/年份；无 DOI（或 DOI 非 Crossref 收录）查 OpenAlex 按题名搜索比对
  - 四级判定：`✅ 已核实`（命中且一致）/ `⚠️ 信息有出入`（命中但题名不符，录入可能有误）/ `❌ 未找到`（数据库查不到，极可能是编造，提示必须人工核实后删除或更正，不要写入论文）/ `🔶 无法联网`（API 不可达，不判假）
  - 字段完整性检查：title/authors/year 缺失、note 未填均提示
  - `--offline` 仅字段检查（不联网）；`--report <路径>.md` 导出可留档的核对报告
  - 对 Crossref/OpenAlex 限流友好：429/5xx 自动等待重试、请求间隔
- **引用对应性核对 `--mapping`**（insert_refs.py）：输出「正文引用 ↔ 文献」对照表——每处编号对应的文献、所在句子片段、note 支撑说明，逐处核对该文献是否真实支撑所在句子；`--dry-run --mapping` 只预览不写文件
- `refs.csv` 的 **`note` 字段升级为必填**（记录该文献支撑正文哪句话/观点），作为"对得上"的核对依据

### 修复
- OpenAlex 高频请求 429 限流 → 增加退避重试与请求间隔（实测从 🔶 无法联网恢复到 ✅ 正常判定）

### 其他
- SKILL.md：新增「步骤 4.5 引用真实性与对应性核实」，流程总览、录入规范、5c 验证、脚本参考同步更新
- README 中英文：特性表、快速开始（新增第 3 步校验）、脚本表、更新记录同步

## [v1.3.0] - 2026-09-10

### 新增
- **分章/多文档支持**：`--docx` 可一次传入多份文档，全文引用统一连续编号；参考文献表生成在主文档（默认最后一个，`--main <路径>` 指定其他）
- **跨文档跳转**：各章正文引用超链接指向主文档（External 关系），点击 `[n]` 打开主文档并跳转到对应文献条目（Word 实测可用）；重跑幂等（编号不变、跨文档链接不重复）
- **规范斜体**：GB/T 7714 要求的西文期刊名/书名自动斜体、中文刊名保持正体；条目文本拆分为独立 run 实现；APA / Vancouver / MLA 的斜体位置同步支持；`--no-italic-source` 可关闭
- 回读验证改为**全文跨文档**编号连续性检查（此前只检查主文档自身，多文档时误报）

### 其他
- README 中英文版：特性表、快速开始（多文档示例）、脚本表、测试清单、注意事项同步更新
- SKILL.md：斜体与分章用法说明

## [v1.2.0] - 2026-09-10

### 新增
- CITE 占位符支持**小写** `[cite:key]` 与**中文 key**（`[CITE:注意力机制]`）
- author-year 模式缺年份时输出 `(Zhang)`（此前为 `(Zhang, )`）
- 新增 `edge_test.py` 边界回归测试（中文 key / 小写占位符 / 无年份 / 正文段保护 / .doc 校验）

### 修复
- **防误删正文**：参考文献标题后的正文段落若以 `[9] 这样开头`，以前会被误判为旧条目删除；现改为「必须紧跟已有条目段」才算条目，标题后第一段正文立即停止清理
- 文档被 Word 占用时，备份/写入前检测并给出可操作提示（此前为晦涩的权限错误）
- 仅支持 `.docx`；`.doc` 给出「在 Word 中另存为」指引
- `insert_refs.py` docstring 补全新参数说明（与 argparse 一致）

### 其他
- SKILL.md：同一处引多篇用连续占位符（`[CITE:k1][CITE:k2]`）、分章/多文档限制说明、期刊名斜体说明
- 文献检索网站可用性清单刷新：18 站，13 可达 / 5 反爬 / 0 断连

## [v1.1.0] - 2026-09-10

### 新增
- `README.md` 完整项目介绍（功能特性 / 快速开始 / 使用流程 / 脚本清单 / 目录结构 / 自动化测试 / 注意事项）
- `README.en.md` 英文版 + 顶部语言切换链接（`[English]` / `[简体中文]`）
- 项目徽章：GitHub Stars（动态）、Python 3.8+、Windows、GB/T 7714-2015、No Zotero
- `assets/demo-screenshot.png` 演示效果截图（正文上标引用 + 文末参考文献表的真实渲染）

## [v1.0.0] - 2026-09-10

### 初版：不依赖 Zotero 的论文引用技能

**核心工作流**
- 定工作文件夹（用户指定或 C 盘外自动创建：论文 / 引用目录 / 文献 三子文件夹）
- 定主题与期刊 → 搜索目标引用格式 → 定引用粒度（句 / 段 / 整篇）
- 逐句判断引用必要性 + 检索筛选（三原则：证据强度 > 时效性 > 引用量；中英文文献全覆盖）
- 收集题录到 `refs.csv`（UTF-8 with BOM，Excel 可开）→ 论文 PDF 按题名自动归档重命名

**引用能力**
- 正文占位符 `[?]` / `[CITE:key]` → 可点击超链接引用（默认上标 `[1]`，Ctrl+点击跳转文末条目）
- 自动重编号：删除中间引用 / 新增引用后重跑，编号自动连续更新；同篇多处引用只占一个编号
- 文末参考文献表自动生成 / 更新；无「参考文献」标题时自动创建（宽容匹配多种标题写法）
- 四种格式：GB/T 7714-2015（默认）/ APA / Vancouver / MLA；编号制 + 作者-年份制（文末按字母序、不编号）

**排版与格式**
- 字体自动：中文（汉字）宋体 + 英文/数字 Times New Roman（rFonts 分别指定，可覆盖）
- 版式：悬挂缩进 2 字符（twips 单位实现，实测首行与正文左边缘对齐）、两端对齐、编号后 Tab 对齐、编号默认不加粗、条目间不空行
- GB/T 7714 完整：专著 [M] / 学位论文 [D] / 会议 [C] 输出「出版地: 出版者」（`city` 字段）；网页 [EB/OL] 输出 `(年)[引用日期]. URL`；作者三截断按首作者语种（中文「等」/ 西文 et al.）

**工程能力**
- `new_project.py` 一键建项目；`add_refs.py` 参数/交互双模式录入（写入前备份、type 校验）
- `rename_papers.py` 文件名 + PDF 首页文本双重匹配归档（重名自动加序号）
- 18 个文献检索网站内置清单，`check_websites.py` 代码化可用性检查（并发、超时、确定性判定），支持每周定时自动检查更新
- 每次修改 docx / refs.csv 前自动备份（保留最近 30 份）+ 写入后回读验证
- 演示素材与自动化测试：`make_demo.py` / `renumber_test.py` / `author_year_test.py` / `verify_docx.py`

**审查修复（随初版发布）**
- 🔴 F1：`--bracket ()` 编号制下第二次运行报「未找到引用点」、重编号失效 → 正则增加 `(n)` 识别
- 🔴 F2：GB 格式专著 / 学位论文 / 网页不完整（缺出版地、缺 URL 与引用日期）→ 新增 `city` 字段、`[EB/OL]` 输出获取路径
- 其余 14 项优化：无 Hyperlink 样式时补蓝色下划线 fallback、`--refs` 缺失时工作区路径提示、备份保留 30 份、MLA DOI 补 `https://doi.org/` 前缀、307/308 重定向计可达、标题宽容匹配避免重复建标题、类型校验、路径动态化、中文作者拼音排序、GB 混排截断符号语种判断、PDF 重名加序号、回读验证同步、死代码清理、author-year 自动化回归

---
*版本号说明：本仓库已随以上变更发布至 v1.9.0；后续变更将持续在此文件记录。*
