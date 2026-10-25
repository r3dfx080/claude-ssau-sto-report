---
name: ssau-sto-report
description: Formatting and structure rules of Samara University (Самарский университет, СГАУ, SSAU, "имени Королёва") standard СТО 02068410-004-2018 "Общие требования к учебным текстовым документам". Use whenever the user writes, formats, converts, or checks any SSAU academic text document — курсовая, ВКР / дипломная, пояснительная записка, отчёт по практике или лабораторной, реферат, РГР, контрольная — or asks about нормоконтроль, титульный лист, задание на ВКР, оформление рисунков/таблиц/формул/списка литературы for their university, even if they don't name the standard. Also use when generating a .docx/.tex/markdown report for an SSAU student, or reviewing one for compliance.
---

# СТО 02068410-004-2018 — SSAU academic text documents

Mandatory for all SSAU students (СПО, бакалавриат, магистратура, специалитет) for any учебный текстовый документ: ВКР explanatory notes, coursework, lab/practice reports, рефераты, РГР, etc. Edition of September 2019 (with Change No. 1).

Use it in two modes:
- **Producing** a document (docx, LaTeX, Markdown→docx): apply every rule below while generating; for title pages read `references/title-pages.md`, for the bibliography read `references/bibliography.md`.
- **Checking** a document: walk the checklist at the end and report each violation with the clause number (e.g. "п. 6.1.8: caption must be below the figure") so the student can argue with the нормоконтролер if needed.

Write explanations and reports in the language the user writes in (most SSAU users write Russian). Document text, quoted fixes, and the standard's terms (раздел, подраздел, реферат, приложение…) always stay in Russian, because that is what the student will paste back into the document.

## 1 Page and typography (п. 3)

| Parameter | Value |
|---|---|
| Paper | A4, one side only |
| Font | Times New Roman, black, 12–14 pt (≥ 1.8 mm); 14 pt is the usual choice |
| Bold | not used (п. 3.2) |
| Line spacing | 1.5 |
| Alignment | justified |
| First-line indent | 1.25 cm |
| Margins | left 30 mm, right 15 mm, top 20 mm, bottom 20 mm |
| Page numbers | Arabic, bottom centre, no period; continuous through appendices; title page counts but shows no number (п. 4.2) |

Each structural element and each section (раздел) starts on a new page (п. 3.4). Small typos may be fixed by hand in black ink; no smudges or traces of erased text.

Note: the scanned standard itself sets its own headings in bold despite п. 3.2. If bold headings matter for the user, tell them to confirm with their кафедра's нормоконтролер; default to no bold.

## 2 Document structure (п. 4.1.1)

In this order:

1. Титульный лист — university/department form; templates in `references/title-pages.md`
2. Задание (if issued) — usually a standard form
3. РЕФЕРАТ
4. СОДЕРЖАНИЕ
5. ВВЕДЕНИЕ
6. Main part — numbered sections `1`, `2`, …
7. ЗАКЛЮЧЕНИЕ
8. ОПРЕДЕЛЕНИЯ, ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ (if any) — note: after the conclusion, unlike ГОСТ 7.32
9. СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ (if any)
10. ПРИЛОЖЕНИЕ А, Б, … (if any)

Structural element headings (РЕФЕРАТ, СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, …): not numbered, centred, ALL CAPS, no trailing period, not underlined (п. 4.1.2).

## 3 Sections and numbering (п. 4.1.3–4.1.10)

- Hierarchy: раздел `1` → подраздел `1.1` → пункт `1.1.1` → подпункт `1.1.1.1`. No period after the last digit. All written from the paragraph indent (1.25 cm), not centred.
- Разделы and подразделы must have headings: sentence case (capital first letter), no trailing period, no underline, **no hyphenation**. Multiple sentences in a heading are separated by periods; still no final period.
- Пункты/подпункты usually have no headings.
- A section with a single subsection still numbers it (`1.1`).
- Lists, each item from the paragraph indent:
  - plain: each item starts with a hyphen `-`;
  - when items are referenced: `а)`, `б)`, `в)` … — skip ё, з, й, о, ч, ъ, ы, ь;
  - nested level under a letter: `1)`, `2)`; deeper: hyphen again.
  - Items end with `;`, the last with `.`; an item that introduces a sub-list ends with `:`.

## 4 Content of structural elements (п. 5)

**Реферат** (п. 5.3, example per ГОСТ 7.9, Appendix В is mandatory):
```
РЕФЕРАТ

Пояснительная записка: 85 с., 24 рисунка, 12 таблиц, 46 источников, 2 приложения.
Графическая часть: 9 листов формата А1.
КЛЮЧЕВОЕ СЛОВО, ДРУГОЕ КЛЮЧЕВОЕ СЛОВО, …
Объектом исследования является …
Цель работы – …
В процессе работы …
В результате работы …
Эффективность (значимость) работы заключается в …
```
- Counts of pages, figures, tables, sources, appendices; graphic sheets and their format (omit the line if there are none).
- 5–15 keywords/phrases from the text, nominative case, CAPS, one paragraph from the indent, comma-separated, no hyphenation, no final period.
- Text covers: object, goal, method, results and novelty, key characteristics, application area, economic efficiency/significance (for ВКР).
- Text ≤ 850 characters (recommended).

**Содержание** (п. 5.4): lists введение, every numbered section/subsection (and named пункты), заключение, definitions, sources, appendices ("Приложение А Название") with start page numbers. Entries in sentence case. Введение, заключение, sources, appendices are not numbered.

**Введение** (п. 5.5): state of the problem, input data, justification of the method, reference to the directive document if any; must show актуальность and новизна.

**Основная часть** (п. 5.6): content defined by the кафедра; text, tables, figures; split into sections.

**Заключение** (п. 5.7): brief conclusions and assessment of results; may include recommendations.

**Определения, обозначения и сокращения** (п. 5.8): a column; left — term/abbreviation/symbol in alphabetical order; right, after a dash — its expansion with units where applicable:
```
ВКР – выпускная квалификационная работа;
σв – временное сопротивление разрыву, МПа;
```

**Список использованных источников** (п. 5.9): per ГОСТ 7.1 (or ГОСТ Р 7.0.5); ordered by first mention in the text; numbered `1`, `2` … with no period, from the paragraph indent. Patterns in `references/bibliography.md`.

**Приложения** (п. 5.10):
- Go after the sources list; each starts on a new page; ordered by first reference in the text; every appendix must be referenced ("…в приложении А").
- Top centre: `ПРИЛОЖЕНИЕ А`; next line centred: its title, capital first letter, no period.
- Letters: А Б В Г Д Е Ж И К Л М Н П Р С Т У Ф Х Ц Ш Щ Э Ю Я (skip Ё З Й О Ч Ъ Ы Ь). A single appendix is still `ПРИЛОЖЕНИЕ А`. Latin letters or digits only after exhausting Cyrillic.
- Internal numbering prefixed with the letter: sections `А.1`, figures `Рисунок А.3`, tables `Таблица Б.1`, formulas `(Б.2)`.
- Share page numbering with the document; listed in Содержание; A4 (A3 allowed).

## 5 Figures (п. 6.1)

- Placed right after the first mention, on the next page, or in an appendix. Every figure is referenced: "…в соответствии с рисунком 1…" / "…рисунком 1.12…". Colour is allowed.
- Numbering: continuous `Рисунок 1`, `Рисунок 2`… or per section `Рисунок 1.12`. A single figure is `Рисунок 1`.
- Below the figure, centred: optional explanatory legend first, then the caption `Рисунок 1 – Проверка прочности шарового пальца` — en dash, capital after it, **no trailing period**. Multi-line caption: single spacing, no hyphenation.

## 6 Tables (п. 6.2)

- Placed right after the first mention or on the next page; referenced as "…в таблице 1…", "(таблица Б.1)".
- Title **above**, left-aligned, **no indent**: `Таблица 1 – Осреднённые результаты испытаний`, no trailing period; multi-line title at single spacing.
- Numbering continuous or per section (`Таблица 6.4`); appendices `Таблица Б.1`; a single table is `Таблица 1`.
- Column/row headers: capital letter, singular; subheaders that continue the header's sentence start lowercase; no periods. Header row separated by a line; no diagonal lines in header cells. Smaller font than body is allowed. Avoid a "№ п/п" column. Landscape orientation is allowed.
- Empty cell → dash `-`; empty cells are not allowed. Repeated one-word text → `»` after the first; multi-word → `То же` first, then `»`. Never ditto numbers, grades, signs, or symbols.
- Splitting: across the page the continuation is headed `Продолжение таблицы 2` (left, no title, no period); the first part has no bottom border. Header row (or head column for wide tables) is repeated, or replaced by column numbers `1 2 3…` set in the first part and repeated.
- Table note: at the end of the table, inside it, above the closing line.
- Very small data sets go into text as dot-leader columns, not a table.

## 7 Formulas (п. 6.3)

- Each on its own line, at least one blank line above and below.
- Wrapping: break after `=`, `+`, `−`, `×`, `:` and repeat the sign at the start of the next line; for multiplication at a break use `×` only. Don't use `*` or `·` as multiplication signs.
- Explanation of symbols directly below, left-aligned, starting with `где` (no colon), in the order they appear, each with its unit:
  ```
  ρ = m / V,                                         (1)
  где m – масса образца, кг;
      V – объем образца, м³.
  ```
- Consecutive formulas not separated by text are separated by a comma or semicolon.
- Only referenced formulas are numbered: `(1)` at the far right on the formula's (last) line; per section `(3.1)`; appendix `(Б.2)`. Referenced as "…по формуле (2.1)…".
- Black; don't mix handwritten and typed parts within one formula.

## 8 Citations and notes (п. 6.4)

- Source references: `[3]` — square brackets, numbered by first mention across the whole document.
- Cite a source as a whole or its sections/appendices, never its subsections, tables, or figures.
- Standards are cited by designation; the year may be dropped if the full title is in the list.
- One note: `Примечание – Текст.` (capital, from the indent). Several: `Примечания` on its own line, then `1 Текст.`, `2 Текст.` (no period after the digit). Placed right after the material it relates to. Notes don't contain requirements.

## 9 Writing rules (п. 7)

- Concise, exact, unambiguous. Requirements: "должен", "следует", "необходимо", "не допускается"…; otherwise "могут быть", "как правило", "при необходимости"; impersonal forms ("применяют", "указывают") are fine.
- No colloquialisms, jargon, synonyms for one term, foreign words when a Russian equivalent exists, coined words, or abbreviations beyond spelling rules and ГОСТ Р 7.0.12 (units in table headers and symbol explanations excepted).
- In running text (not formulas, tables, figures):
  - write "минус" instead of `−` before negative values;
  - write "диаметр" instead of `⌀`;
  - no standard index without its number (not just "ГОСТ");
  - no `>`, `<`, `=`, `≠`, `≈`, `≥`, `≤`, `№`, `%` without numbers ("процент", not "%").
- Explain a symbol before its first use ("…временное сопротивление разрыву σв…"). Non-standard symbols go into the definitions list.
- SI units per ГОСТ 8.417; other permitted units in parentheses; one unit system per document; one unit per parameter throughout.
- Numbers with units → digits; counts without units 1–9 → words: "испытания пяти труб, каждая длиной 12 м".
- Series and ranges: the unit only after the last value: "1,50; 1,75; 2,00 м", "от 1 до 5 мм", "от плюс 10 до минус 40 °С". Never split a number from its unit across lines (use a non-breaking space).
- Limits: "должно быть не более (не менее) …"; deviations: "не должно быть более (менее) …".
- Decimal comma; align decimal places within a series. Inches as slashed fractions (¾"). Otherwise if a decimal is impossible, a one-line fraction: `5/32`, `(50A – 4C)/(40B + 2E)`.

## 10 docx setup (python-docx)

```python
from docx.shared import Mm, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
s = doc.sections[0]
s.page_width, s.page_height = Mm(210), Mm(297)
s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Mm(30), Mm(15), Mm(20), Mm(20)
st = doc.styles["Normal"]
st.font.name, st.font.size, st.font.bold = "Times New Roman", Pt(14), False
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")  # Cyrillic otherwise falls back to another font
pf = st.paragraph_format
pf.line_spacing, pf.first_line_indent = 1.5, Cm(1.25)
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
pf.space_before = pf.space_after = Pt(0)
```
Page number: a centred PAGE field in the footer with "different first page" enabled so the title page shows none. Captions, centred headings and the title page need `first_line_indent = 0`.

## Compliance checklist

Report each failure with its clause:
- [ ] A4, margins 30/15/20/20 mm, TNR 12–14, 1.5 spacing, justified, 1.25 cm indent, no bold (3.1–3.2)
- [ ] Element order matches §2; each element/section on a new page (3.4, 4.1.1)
- [ ] Element headings CAPS, centred, unnumbered, no period (4.1.2)
- [ ] Section numbers without trailing period; headings from the indent, no period, no hyphenation (4.1.4–4.1.7)
- [ ] Lists use `-` / `а)` (skipping forbidden letters) / `1)` (4.1.10)
- [ ] Page numbers bottom centre, none on title page, continuous through appendices (4.2)
- [ ] Реферат: counts, 5–15 CAPS keywords, ≤ 850 characters (5.3)
- [ ] Содержание complete with page numbers (5.4)
- [ ] Sources ordered by first mention, ГОСТ 7.1 format, `[n]` citations (5.9, 6.4.1)
- [ ] Appendices: `ПРИЛОЖЕНИЕ А` with allowed letters, referenced, in contents (5.10)
- [ ] Figures: referenced, caption below centred "Рисунок N – Name", no period (6.1)
- [ ] Tables: referenced, title above left no indent "Таблица N – Name", no empty cells, "Продолжение таблицы N" (6.2)
- [ ] Formulas: own line, number `(N)` at right, "где" without colon, units given, no `*` (6.3)
- [ ] Text rules: "минус", no bare `%`/`№`/signs, units after last value, 1–9 in words (7)
