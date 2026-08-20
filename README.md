# workbuddy-training-manual

> **分类**：原创 / AI 打磨 ｜ **文件数**：9 ｜ **仓库目录**：`workbuddy-training-manual`

## 📌 简介

WorkBuddy

## 🎯 适用场景

适用于该技能的能力范围，详见下方「📖 使用说明」。

## 📂 目录结构

```text
  - .gitignore
  - LICENSE
  - README.md
  - SKILL.md
  - **assets/**
    - manual_template.json
  - **references/**
    - manual_data_schema.md
    - question_templates.md
    - training_design_framework.md
  - **scripts/**
    - docx_exporter.py
```

## 🚀 安装方法

将本文件夹整体复制到 WorkBuddy 的技能目录即可启用：

```bash
# 用户级（推荐）
cp -r . ~/.workbuddy/skills/workbuddy-training-manual

# 或项目级
cp -r . <你的项目>/.workbuddy/skills/workbuddy-training-manual
```

复制完成后，**重启或刷新 WorkBuddy**，即可在对话中用自然语言触发该技能。

## ⚙️ 配置说明

本技能开箱即用，**无需额外配置**。若涉及外部 API 调用，请在使用时按需提供您自己的密钥（不要提交到公开仓库）。

## 📖 使用说明（完整规范）

> 以下为该技能的完整说明，涵盖核心能力、工作流程与关键规则，帮助您全面了解其运作方式。

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

## 💡 命令示例

```bash
python scripts/docx_exporter.py --type manual --input <data.json> --output <output.docx>
```

## ⚠️ 注意事项

- 本技能从本地 WorkBuddy 环境导出，**所有真实密钥 / 凭据 / 个人数据均已脱敏为占位符**，重新使用前请配置您自己的 Key。
- 如为原创技能，可自由使用、修改与再分发；若对外分享请保留作者与来源信息。
- 技能提供的是自动化辅助能力，不替代专业判断；涉及交易、法律、医疗等高风险场景请谨慎并自担风险。

## 📄 许可证

MIT License —— 详见仓库内 `LICENSE` 文件。
