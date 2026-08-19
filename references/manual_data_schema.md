# Manual Data JSON Schema

Complete JSON structure for training manual generation via `docx_exporter.py --type manual`.

## Top-Level Structure

```json
{
  "manual_name": "...",
  "system_version": "...",
  "purpose": "...",
  "audience": "...",
  "section_titles": { ... },
  "prerequisites": [ ... ],
  "process_overview": "...",
  "procedures": [ ... ],
  "faq": [ ... ],
  "error_handling": "...",
  "checklist": [ ... ]
}
```

## Field Reference

### `manual_name` (string, required)
Title displayed on the cover page. Example: "XX 产品快速上手培训手册"

### `system_version` (string, optional)
System or product version info. Example: "v2.4.0"

### `purpose` (string, required)
The main overview / product introduction. Supports newlines (`\n`). Should cover:
- What the product/tool is (one-sentence definition)
- Core value proposition
- Target audience and applicable scenarios

### `audience` (string, optional)
Short description of the target audience for the cover page.

### `section_titles` (object, optional)
Override default Chinese section headings. Available keys:

| Key | Default Value |
|-----|---------------|
| `purpose` | 一、手册用途 |
| `prerequisites` | 二、使用前提 |
| `process_overview` | 三、操作流程总览 |
| `procedures` | 四、操作步骤 |
| `faq` | 五、常见问题（FAQ） |
| `error_handling` | 六、异常处理 |
| `checklist` | 七、操作检查清单 |

### `prerequisites` (array of strings, optional)
List of prerequisites before using the manual. Each string is one item.

### `process_overview` (string, optional)
Overview of the learning path / navigation guide shown before procedures.

### `procedures` (array of objects, required)
The core operational scenarios. Each object:

```json
{
  "proc_id": "场景一",
  "proc_name": "用 AI 快速写周报",
  "objective": "帮助用户...",
  "steps": [
    {
      "step_num": 1,
      "action": "详细的操作描述...",
      "screenshot_note": "此处截图：具体的截图位置说明...",
      "expected": "预期的结果状态...",
      "tips": "小技巧或注意事项..."
    }
  ],
  "verification": "验证标准..."
}
```

**Procedure fields:**
- `proc_id` — Scenario identifier (e.g., "场景一")
- `proc_name` — Human-readable scenario name
- `objective` — What the user will achieve after this scenario
- `steps[]` — Array of step objects (recommended 3–5 per scenario)
- `verification` — How to verify successful completion

**Step fields:**
- `step_num` — Integer step number (1-based)
- `action` — Detailed action description
- `screenshot_note` — Screenshot placeholder text
- `expected` — Expected outcome after this step
- `tips` — Practical tip for this step

### `faq` (array of objects, optional)
Frequently asked questions. Each object:

```json
{
  "question": "How do I...?",
  "cause": "Why users ask this...",
  "solution": "The answer..."
}
```

Recommended 10–20 items. Cover: login, permissions, billing, security, file formats, mobile support, error recovery, etc.

### `error_handling` (string, optional)
Free-text content. Can be used for advanced tips, best practices, troubleshooting, etc. Supports numbered lists and line breaks.

### `checklist` (array of strings, optional)
Completion checklist. Each string is one checklist item. Recommended 5–10 items.

## Usage Notes

- All Chinese content should use polite, accessible language
- Prefer "您" over "你" in instructional content
- `screenshot_note` fields should clearly indicate what the screenshot should show
- `tips` fields should provide actionable, non-obvious guidance
- `cause` in FAQ items helps the AI understand context but is not directly printed — only `question` and `solution` appear in the output
