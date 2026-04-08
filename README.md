# LimeSurvey Excel Builder

Design LimeSurvey questionnaires entirely in Excel — with multi-language support, automatic HTML formatting, and one-click import.

## What is this?

An Excel template + R conversion script that lets you author LimeSurvey surveys in a spreadsheet instead of the web interface. Write your questions, answers, skip logic, and translations in Excel, run the R script, and import the resulting file directly into LimeSurvey.

**Why use this instead of the LimeSurvey UI?**

- **Format text directly in Excel** — bold, italic, underline, and color any word or phrase in Excel; the exact formatting will be preserved in the questionnaire, including partial formatting within a cell
- **Faster authoring** — type questions, answers, and logic in a familiar spreadsheet environment
- **Easy translations** — add language columns (`text_fr`, `text_es`, `text_ro`, ...) instead of duplicating rows
- **Batch editing** — find/replace, copy/paste, and reorder questions using standard Excel operations
- **Offline work** — no need for a running LimeSurvey instance during survey design
- **Reusable blocks** — copy question groups between surveys by copying rows
- **Quota management** — define response quotas per segment directly in a dedicated Excel sheet, with translated quota-full messages and AND logic across multiple criteria

## Quick Start

### Option 1: Use the Web App (no installation needed)

Go to **[limesurvey-excel-builder](https://limesurvey-excel-builder.60.md/)**, upload your `.xlsx` file, and download the converted `.txt` file ready for LimeSurvey import.

### Option 2: Run the R Script Locally

1. Open `limesurvey_survey_builder.xlsx` in Excel
2. Edit the **Survey Design** sheet — add your questions, answers, and settings
3. (Optional) Edit the **Quotas** sheet — define response quotas per segment
4. Run `xlsx_to_limesurvey_tsv.R` in RStudio
5. In LimeSurvey: **Create Survey > Import > select the generated `.txt` file**


## Features

### Excel Template

- **Color-coded rows** — conditional formatting highlights groups, questions, subquestions, answers, and settings in different colors
- **Dropdown validation** — class, mandatory, and other columns have dropdown lists
- **4 reference sheets** — Question Types, Relevance & Validation, Survey Settings, Instructions
- **Quotas sheet** — define response quotas with translated messages, AND-logic members, and active/inactive toggle
- **Per-language header colors** — each language gets a distinct color across Survey Design and Quotas sheets
- **41 advanced attribute columns** — array filtering, sliders, date ranges, validation, and more
- **Example survey** — 178 rows covering 30+ question types with skip logic, validation, calculated fields, and tailored text

### Multi-Language Support

Instead of duplicating every row for each language, add columns:

| text_en | help_en | text_fr | help_fr | text_es | help_es |
|---------|---------|---------|---------|---------|---------|
| What is your age? | | Quel est votre âge ? | | ¿Cuál es su edad? | |

The R script automatically:
- Infers the base language from the **first** `text_xx` column (e.g. `text_en` → `en`)
- Treats remaining `text_xx` columns as additional languages, in left-to-right column order
- Generates the multi-language TSV rows LimeSurvey expects
- Falls back to the base language (or any available language) for untranslated cells
- Auto-sets both the `language` and `additional_languages` survey settings from the detected columns — no manual S rows needed

### Automatic HTML Conversion

Write plain text in Excel. The R script handles all HTML:

- **Line breaks** (Alt+Enter in Excel):
  - Q/G/SL text → separate `<p>` paragraphs
  - SQ/A text → `<br>` tags (compact labels)
- **Cell formatting** → HTML tags:
  - Bold → `<strong>`
  - Italic → `<em>`
  - Underline → `<u>`
  - Font color → `<span style='color:#HEX'>`
- **Partial formatting** — if only one word in a cell is bold or colored, only that word gets HTML tags
- Cells already containing HTML are passed through unchanged
- **S rows are always plain text** — HTML conversion is never applied to survey settings rows, regardless of cell formatting. If formatting is detected on an S row (e.g. LibreOffice auto-formatting an email address as a hyperlink), the script emits a warning and ignores it

### Survey Structure

Each row in the Excel sheet represents one survey element:

| Class | Meaning | Example |
|-------|---------|---------|
| **S** | Survey setting | `format = G` (group-by-group) |
| **SL** | Survey language text | Survey title, welcome text, end text |
| **G** | Question group | Section header |
| **Q** | Question | A question with type, code, relevance |
| **SQ** | Subquestion | Row/column in array or multiple choice |
| **A** | Answer option | Choice in a list or scale |

Row order matters: S → SL → G → Q → SQ → A → Q → SQ → A → ... → G → ...

### Quota Support

Define response quotas in the dedicated **Quotas** sheet — one row per quota, no TSV syntax needed:

| quota_name | quota_limit | active | quota_action | autoload_url | message_en | message_ro | question_code_1 | answer_code_1 | question_code_2 | answer_code_2 |
|---|---|---|---|---|---|---|---|---|---|---|
| Males North | 15 | N | 2 | 0 | Quota full. Thank you. | Cota este plină. | gender | M | region | N |
| Females North | 15 | N | 2 | 0 | Quota full. Thank you. | Cota este plină. | gender | F | region | N |

**Columns:**

- **quota_name** — unique name for this quota
- **quota_limit** — maximum number of completed responses before the quota triggers
- **active** — `Y` = enforced, `N` = disabled. Set `N` while testing, switch to `Y` when ready
- **quota_action** — `1` = terminate survey silently, `2` = terminate and show message
- **autoload_url** — `0` = no redirect, `1` = auto-redirect when quota is full
- **message_xx** — translated message shown when the quota is full (one column per survey language, falls back to base language if empty)
- **question_code_N / answer_code_N** — each pair links an answer to the quota. Add `_3`, `_4`, etc. for more conditions

**How members work:**

All members within a quota are **ANDed** together: the quota only counts a response when ALL members match.

Example: `question_code_1=gender, answer_code_1=M, question_code_2=region, answer_code_2=N` → counts only male respondents from the North region.

**Important notes:**

- Only **list-type questions** (L, !) can be quota members — not numeric (N) or text (S)
- To quota on numeric ranges (e.g. age > 40), create a hidden list question with `always_hide=1` and a default expression like `{if(age >= 40, 'O40', 'U40')}`
- The R script generates the correct QTA, QTALS, and QTAM rows in the TSV with proper placement and ID linking
- The inverse converter extracts quotas from existing TSV exports back into the flat Quotas sheet format

### Per-Language Header Colors

Each language gets a distinct header color in both the Survey Design and Quotas sheets, making it easy to visually identify language columns:

| Language | Color |
|----------|-------|
| en | Dark blue-gray |
| ro | Purple |
| fr | Blue |
| es | Orange |
| de | Green |
| ar | Red |
| pt | Teal |

Additional languages are auto-assigned from a fallback palette. The same color is used for `text_xx`/`help_xx` in Survey Design and `message_xx` in Quotas.

## Requirements

- **Excel** (or LibreOffice Calc) for editing the template
- **R** (>= 4.0) with RStudio
- R packages (auto-installed on first run): `readxl`, `tidyxl`, `xml2`
- **LimeSurvey** (>= 3.x) for import

## Configuration

In the R script, set the input filename:

```r
input_file <- "limesurvey_survey_builder.xlsx"
```

The output file uses the same base name with a `.txt` extension.

## Question Codes

LimeSurvey question codes (the `name` column) must:
- Start with a letter
- Contain only alphanumeric characters (a-z, A-Z, 0-9)
- **No underscores**, hyphens, spaces, or special characters

Good: `hhsize`, `cigsperday`, `dietfreq`  
Bad: `hh_size`, `cigs_per_day`, `diet-freq`

The R script warns about codes containing underscores.

## Adding Translations

1. Insert new columns after the existing language columns (e.g., `text_de`, `help_de`)
2. Translate the text for G, Q, SQ, A, and SL rows
3. Leave `text_xx` empty for S rows (settings are language-independent)
4. Run the R script — it auto-detects the new language columns and updates the language settings automatically

Untranslated cells automatically fall back to the base language, so you can translate incrementally.

## Example Survey

The included example demonstrates:

- 6 question groups with 41 questions
- Question types: single choice, multiple choice, arrays, numeric, text, date, file upload, ranking, dual-scale, equation, boilerplate
- Skip logic (relevance conditions)
- Input validation (regex, numeric ranges)
- Calculated fields (BMI from height/weight)
- Array filtering (show only selected conditions)
- Tailored closing message with expression logic
- Multi-language content (English, French, Romanian, Spanish)
- 6 example quotas (Males/Females × North/Center/South regions with translated messages)

## Validation

The R script validates your survey before export:

- Checks for valid class values
- Detects duplicate question codes per language
- Warns about underscores in question codes
- Warns about rich text formatting on S rows (e.g. LibreOffice hyperlink auto-formatting)
- Reports row counts by class type (including QTA, QTALS, QTAM when quotas are present)
- Lists advanced attributes in use
- Reports quota count and member placement

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Colors don't show in Excel | Conditional formatting uses `bgColor` — may need a restart of Excel |
| Import fails silently | Check that all question codes are alphanumeric (no underscores) |
| Translations don't appear | The script auto-generates all rows per language; ensure columns are named `text_xx` / `help_xx` |
| Special characters garbled | Output uses UTF-8 with BOM — import should auto-detect encoding |
| LimeSurvey renames codes | Codes with underscores get stripped; use alphanumeric only |
| Import error on email/URL settings | LibreOffice auto-formats email addresses and URLs as colored hyperlinks. The script detects and ignores this formatting for S rows, but check the conversion log for warnings |
| Quotas don't appear after import | Verify class is `QTA` (not `QT`), messages go in `relevance` column of QTALS, and QTAM rows are placed after their questions. The script handles this automatically from the Quotas sheet |
| Quota not counting responses | Only list-type questions (L, !) can be quota members. Numeric (N) and text (S) questions cannot be used directly — create a hidden list question with a default expression as a workaround |
| Quota message shows in wrong language | Ensure the `message_xx` columns in the Quotas sheet match the survey languages. Empty cells fall back to the base language |

## Inverse Converter: LimeSurvey to Excel

If you need to convert an existing LimeSurvey survey back to the Excel Builder format for editing, an additional script and web tool are available.

This is useful when you want to edit a survey that was originally created in the LimeSurvey web interface.

### How to use

**Option A: Web app** -- go to [limesurvey-excel-builder](https://limesurvey-excel-builder.60.md/), scroll to the second section, upload your `.txt` export, and download the `.xlsx` file.

**Option B: R script** -- set `input_file` in `limesurvey_tsv_to_xlsx.R` and run it in RStudio. Requires the `openxlsx2` package.

> **Note:** The forward and inverse R scripts use packages that may conflict with each other (`xml2` and `openxlsx2`). If running both in the same RStudio session, restart R between them (Ctrl+Shift+F10).

### What it does

- Collapses multi-language rows into side-by-side `text_xx` / `help_xx` columns
- Converts HTML formatting back to Excel rich text (bold, italic, underline, color)
- Extracts quotas (QTA/QTALS/QTAM) into the flat Quotas sheet format with per-language messages and AND-logic members
- Drops server-specific settings for cross-server portability
- Produces a formatted workbook with conditional formatting, per-language header colors, data validations, and reference sheets (including Relevance & Validation)
- The resulting `.xlsx` can be edited and converted back to `.txt` with `xlsx_to_limesurvey_tsv.R`


## Files

| File | Description |
|------|-------------|
| `limesurvey_survey_builder.xlsx` | Excel template with example survey, quotas, reference sheets, and instructions |
| `xlsx_to_limesurvey_tsv.R` | R script that converts the Excel file to LimeSurvey TSV import format (with quota support) |
| `limesurvey_tsv_to_xlsx.R` | R script that converts a LimeSurvey TSV export back to the Excel Builder format (with quota extraction) |

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests for:

- Additional question type examples
- New language translations for the example survey
- Bug fixes or improvements to the R script
- Documentation improvements
