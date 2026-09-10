# Direct Citation Assistant

**[English](README.en.md) | [简体中文](README.md)**

> A Zotero-free, script-driven citation workflow for Word: from "sentence-by-sentence citation judgment" → "literature search & filtering" → "clickable hyperlink citations" → "auto-generated reference list".

Supports **journal papers, theses (proposal / mid-term reports), and course papers**, with mixed **Chinese & English references**. Reference formatting follows **GB/T 7714-2015** (default), plus APA / Vancouver / MLA.

---

## ✨ Features

| Capability | Description |
|---|---|
| 🗂️ Auto workspace | One command creates the project folder (论文 / 引用目录 / 文献, i.e. drafts / citation list / PDFs) — automatically placed outside drive C: |
| 🔍 Sentence-level judgment | A concrete checklist for deciding whether each sentence needs a citation, avoiding citation padding |
| 🎯 Three search principles | Strongest evidence first → newest first → highest citation count; covers both Chinese & English literature |
| 🔗 Hyperlink citations | In-text numbers (superscript by default) are **clickable and jump** to the matching entry in the reference list |
| 🔢 Auto renumbering | Re-run after deleting / adding a citation — numbers update automatically, no manual fixing |
| 📑 Four formats | GB/T 7714 (default), APA, Vancouver, MLA; numbered + author-year styles |
| 🌐 Site checks | Built-in list of 18 literature-search sites, availability checked by code, with scheduled updates |
| 📝 Complete GB format | Journal, conference, **book [M], thesis [D], web page [EB/OL]** all rendered per GB/T 7714 |
| ✒️ Auto typography | Chinese: SimSun (宋体) + Western/digits: Times New Roman; hanging indent 2 chars, justified, tab-aligned numbers |
| 📄 PDF archiving | Downloaded papers renamed by title and archived; name collisions get numbered suffixes, never overwritten |
| 💾 Auto backup | Every docx / refs.csv modification is backed up first (last 30 kept) |

---

## 🚀 Quick Start

### Requirements

- Windows (scripts rely on Windows drive detection & Word)
- Python 3.8+
- Dependencies: `lxml` (required), `python-docx` (demo generation), `pypdf` (PDF matching, optional), `pypinyin` (Chinese sorting, optional)

```bash
git clone https://github.com/xiaolinzi000/direct-citation-assistant.git
pip install lxml python-docx pypdf pypinyin
```

### 1. Create a workspace

```bash
python scripts/new_project.py --name "thesis_proposal"
```

Creates (outside drive C:):

```
thesis_proposal/
├── 论文/        # your Word drafts
├── 引用目录/    # refs.csv citation database
└── 文献/        # downloaded paper PDFs
```

### 2. Add a reference (copy from Google Scholar / CNKI)

```bash
python scripts/add_refs.py --refs <project>/引用目录/refs.csv ^
    --title "Attention is all you need" ^
    --authors "Vaswani, A., Shazeer, N., Parmar, N., et al." ^
    --source "Advances in Neural Information Processing Systems" ^
    --year 2017 --volume 30 --doi 10.5555/3295222.3295349 --type journal
```

### 3. Place placeholders in your text and insert citations

Put `[?]` (auto-picks the next unused entry) or `[CITE:key]` (cite a specific entry) at the end of the sentence:

```bash
python scripts/insert_refs.py --docx <project>/论文/文稿.docx ^
    --refs <project>/引用目录/refs.csv --style gbt7714 --citation numbered
```

Result:

- In-text markers become clickable superscripts `[1]` (Ctrl+click jumps to the entry in the reference list);
- A 「参考文献」 heading and list are generated (or reused) at the end of the document;
- Delete a citation anywhere and re-run — the rest renumber automatically.

---

## 📋 Full Workflow

```
① Set workspace → ② Pick topic & journal → ③ Pick granularity (sentence / paragraph / full paper)
→ ④ Search per sentence + three principles → ⑤ Collect entries + download & archive PDFs → ⑥ Insert citations + build reference list
```

Detailed step-by-step instructions live in [SKILL.md](SKILL.md) (the skill master file, with commands and checklists for every step).

### Three search principles (by priority)

1. **Evidence strength**: pick the reference that most strongly supports the sentence (original source / direct conclusion / method paper / authoritative review);
2. **Recency**: among equal strength, prefer the newest; combine "original paper + recent review" for classics;
3. **Citations**: among equal strength & era, prefer highly cited / highly rated works.

### Formatting defaults (all adjustable)

| Item | Default | Flag |
|---|---|---|
| Style system | GB/T 7714-2015 | `--style gbt7714\|apa\|vancouver\|mla` |
| Citation style | numbered, superscript `[1]` | `--citation numbered\|author-year` |
| Bracket | `[]` | `--bracket ()` |
| Hanging indent | 2 chars (21pt at 五号) | `--hanging-pt 21` (24 for 小四 / 28 for 四号) |
| Alignment | justified | `--align both\|left` |
| Number bold | off (common convention) | `--bold-num` |
| Chinese font | SimSun | `--font-cn "宋体"` |
| Western font | Times New Roman | `--font-en "Times New Roman"` |

---

## 🧩 Scripts

| Script | Purpose |
|---|---|
| `new_project.py` | Create project workspace (3 sub-folders outside drive C: + refs.csv template) |
| `add_refs.py` | Write citations into refs.csv (flags / interactive, validates type, auto-backup) |
| `rename_papers.py` | Match & rename downloaded PDFs by title, archive them (collision-safe) |
| `insert_refs.py` | **Core**: placeholders → hyperlink citations; reference list create/update; auto-renumbering |
| `format_refs.py` | Render a single entry as gbt7714 / apa / vancouver / mla |
| `refs_db.py` | refs.csv read/write library (full GB format incl. `city` field) |
| `check_websites.py` | Availability check of the 18 literature-search sites (deterministic; `--update` rewrites the list) |
| `demo/*` | Demo assets & automated tests (`make_demo.py`, `renumber_test.py`, `author_year_test.py`, `verify_docx.py`) |

### Search-site list & scheduled checks

- Full list: `references/websites.md` (site / URL / purpose / status / notes; Chinese, English & field-specific channels);
- Status maintained automatically by `check_websites.py` (✅ reachable / ⚠️ anti-bot / ❌ unreachable); can be wired to a scheduled weekly check;
- To add/remove sites, edit the `SITES` list in `check_websites.py` and re-run.

---

## 📂 Repository Layout

```
direct-citation-assistant/
├── SKILL.md                 # skill master file (workflow + checklists + flags)
├── README.md                # this doc (Chinese)
├── README.en.md             # English version
├── references/
│   └── websites.md          # literature-search site list (status maintained by script)
├── scripts/
│   ├── new_project.py       # create workspace
│   ├── add_refs.py          # add citations
│   ├── rename_papers.py     # archive PDFs by title
│   ├── insert_refs.py       # core: insert citations + reference list + renumbering
│   ├── format_refs.py       # 4-format rendering
│   ├── refs_db.py           # refs.csv I/O
│   ├── check_websites.py    # site availability checks
│   └── demo/                # demo document + automated regression tests
└── .gitignore
```

---

## ✅ Automated Tests

```bash
cd scripts/demo
python make_demo.py          # generate demo assets (refs.csv + 文稿.docx)
python renumber_test.py      # scenario: delete / add citations → auto renumbering
python author_year_test.py   # scenario: author-year → alphabetical order + idempotency
python verify_docx.py 文稿.docx   # structural integrity (hyperlink ↔ bookmark 1:1)
```

---

## ⚠️ Notes

- **Word locks the file**: close the document in Word before running the scripts;
- **`[?]` more than unused entries**: the script errors out — add more references with `add_refs.py` first;
- **Citing the same paper twice**: use `[CITE:same-key]` the second time; the number is reused;
- **Book / thesis / conference**: fill `city` (place of publication / preservation) in refs.csv; `web` entries must fill `url`;
- **In-text position**: number as superscript, directly after the cited content, **before** the sentence-final punctuation (`…significant progress[1].`), no spaces around it;
- **Line spacing**: not mandated by the national standard; follow your school / journal template, select the whole reference list and set it in Word.

---

## 📄 License

Personal project — no open-source license is specified yet. You are welcome to study it; please contact the author before using or redistributing it in papers or products.
