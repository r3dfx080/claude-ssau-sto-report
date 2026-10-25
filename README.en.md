# ssau-sto-report

[Русская версия](README.md)

A [Claude](https://claude.com) skill for writing and checking Samara University (Самарский университет им. Королёва) student documents against **СТО 02068410-004-2018 «Общие требования к учебным текстовым документам»** (2019 edition, with Change No. 1).

Works for курсовые, ВКР / пояснительные записки, отчёты по практике и лабораторным, рефераты, РГР, and similar documents.

## What it does

- **Writes** documents (.docx, LaTeX, Markdown) that follow the standard: page setup, the order of parts, heading and list numbering, figures, tables, formulas, citations, and the writing rules.
- **Checks** an existing document and reports each problem with its clause number (п. 6.1.8 and so on).
- Includes the title page and задание templates (bachelor's, master's, specialist) and bibliography patterns per ГОСТ 7.1.

## Install

**Claude Code**: copy the folder into your skills directory:

```bash
git clone https://github.com/<you>/claude-ssau-sto-report ~/.claude/skills/ssau-sto-report
```

**Claude.ai / Claude desktop**: zip the folder and upload it under Settings → Capabilities → Skills.

Claude picks the skill up automatically when you ask about an SSAU report, ВКР, нормоконтроль, and so on.

## Files

| File | Contents |
|---|---|
| `SKILL.md` | All formatting rules, a python-docx setup snippet, and a compliance checklist |
| `references/title-pages.md` | Title page and задание на ВКР templates |
| `references/bibliography.md` | Bibliography entry patterns by source type |

## Disclaimer

This is an unofficial summary written in my own words, not the standard itself. The official text is at <https://cabinet.ssau.ru/assets/standard/sto_02068410_004_2019.pdf>. Your department's нормоконтролер has the final say, and departments may have their own forms.
