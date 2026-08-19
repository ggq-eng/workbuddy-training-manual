---
name: workbuddy-training-manual
description: WorkBuddy
  培训手册生成器。当用户要求生成培训手册、操作手册、新员工入职培训材料、SOP文档、产品使用指南等培训类Word文档时触发。通过填写结构化JSON模板，一键导出为专业排版、可打印成册的Word文档。适用：入职培训、产品操作手册、技能培训材料、FAQ手册。
agent_created: true
triggers:
  - 培训手册
  - 操作手册
  - 入职培训
  - 新员工培训
  - 培训文档
  - 产品手册
  - 使用指南
  - SOP
  - FAQ手册
  - 快速上手
  - training manual
  - onboarding guide
disable: true
---

# WorkBuddy Training Manual Generator

Generate professional training manuals as `.docx` Word documents from structured JSON data, ready for printing and distribution.

## Workflow

### Step 1 — Understand Requirements

Gather the following from the user (ask only if missing):

- **Manual topic** — What is being taught?
- **Target audience** — Who will read it (role, experience level)?
- **Structure** — What sections are needed? Default structure:
  1. Product / topic overview
  2. Operation procedures (5–10 scenarios, 3–5 steps each)
  3. FAQ (10–20 items)
  4. Advanced tips / best practices (3–5 items)
  5. Completion checklist
- **Word count target** — approximate total Chinese characters
- **Any special formatting or branding requirements**

### Step 2 — Build the JSON Data

Create a JSON file following the structure documented in `references/manual_data_schema.md`. A populated example is available at `assets/manual_template.json`.

Key data fields:

| Field | Purpose |
|-------|---------|
| `manual_name` | Title shown on cover page |
| `purpose` | Overview / product introduction content |
| `prerequisites` | List of preparation items before starting |
| `process_overview` | Learning path / navigation guide |
| `procedures[]` | Each is a scenario: `proc_id`, `proc_name`, `objective`, `steps[]` (with `step_num`, `action`, `screenshot_note`, `expected`, `tips`), `verification` |
| `faq[]` | Each has `question`, `cause`, `solution` |
| `error_handling` | Advanced tips content (free text) |
| `checklist[]` | Completion checklist items |
| `section_titles{}` | Optional — customize section headings (keys: `purpose`, `prerequisites`, `process_overview`, `procedures`, `faq`, `error_handling`, `checklist`) |

### Step 3 — Export to Word

Run the exporter script with the JSON file:

```bash
python scripts/docx_exporter.py --type manual --input <data.json> --output <output.docx>
```

Requirements: `python-docx` (`pip install python-docx`).

### Step 4 — Present Result

Present the generated `.docx` file to the user using the `present_files` tool. Remind the user:
- Use Word's "References → Table of Contents → Insert Table of Contents" to auto-generate the TOC
- Replace screenshot placeholders (`[截图位置：...]`) with actual screenshots
- Content can be further customized in Word

## Style Guidelines

- Use friendly, accessible language — prefer "您" and "只需几步"
- Avoid technical jargon; explain any necessary terms
- Each procedure step should include a `screenshot_note` placeholder for later screenshot insertion
- Each step should include `tips` for practical guidance
- Keep Chinese character count as close to the user's target as structure allows

## Customizing Section Titles

The `section_titles` field in the JSON data allows overriding default section headings:

```json
"section_titles": {
  "purpose": "一、产品概述",
  "prerequisites": "二、上手准备",
  "process_overview": "三、学习导航",
  "procedures": "四、核心操作场景",
  "faq": "五、常见问题",
  "error_handling": "六、进阶技巧",
  "checklist": "七、检查清单"
}
```

Omit this field to use default Chinese section headings.
