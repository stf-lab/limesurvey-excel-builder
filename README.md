# LimeSurvey Excel Builder

Design LimeSurvey questionnaires entirely in Excel — with multi-language support, automatic HTML formatting, and one-click import.

## What is this?

An Excel template + R conversion script that lets you author LimeSurvey surveys in a spreadsheet instead of the web interface. Write your questions, answers, skip logic, and translations in Excel, run the R script, and import the resulting file directly into LimeSurvey.

**Why use this instead of the LimeSurvey UI?**

- **Format text directly in Excel** — bold, italic, underline, and color any word or phrase in Excel; the exact formatting will be preserved in the questionnaire, including partial formatting within a cell
- **Faster authoring** — type questions, answers, and logic in a familiar spreadsheet environment
- **Easy translations** — add language columns (`text_fr`, `text_es`, `text_ro`, ...) instead of duplicating rows
- **Version control** — track changes with Git, collaborate with colleagues on survey design
- **Batch editing** — find/replace, copy/paste, and reorder questions using standard Excel operations
- **Offline work** — no need for a running LimeSurvey instance during survey design
- **Reusable blocks** — copy question groups between surveys by copying rows

## Quick Start

### Option 1: Use the Web App (no installation needed)

Go to **[shiny.os.md/limesurvey-excel-builder](https://shiny.os.md/limesurvey-excel-builder/)**, upload your `.xlsx` file, and download the converted `.txt` file ready for LimeSurvey import.

### Option 2: Run the R Script Locally

1. Open `limesurvey_survey_builder.xlsx` in Excel
2. Edit the **Survey Design** sheet — add your questions, answers, and settings
3. Run `xlsx_to_limesurvey_tsv.R` in RStudio. The script will generate a *.txt file
4. In LimeSurvey: **Create Survey → Import → select the generated `*.txt` file**

## Files

| File | Description |
|------|-------------|
| `limesurvey_survey_builder.xlsx` | Excel template with example survey, reference sheets, and instructions |
| `xlsx_to_limesurvey_tsv.R` | R script that converts the Excel file to LimeSurvey TSV import format |

## Features

### Excel Template

- **Color-coded rows** — conditional formatting highlights groups, questions, subquestions, answers, and settings in different colors
- **Dropdown validation** — class, mandatory, and other columns have dropdown lists
- **4 reference sheets** — Question Types, Relevance & Validation, Survey Settings, Instructions
- **41 advanced attribute columns** — array filtering, sliders, date ranges, validation, and more
- **Example survey** — 178 rows covering 30+ question types with skip logic, validation, calculated fields, and tailored text

### Multi-Language Support

Instead of duplicating every row for each language, add columns:

| text_en | help_en | text_fr | help_fr | text_es | help_es |
|---------|---------|---------|---------|---------|---------|
| What is your age? | | Quel est votre âge ? | | ¿Cuál es su edad? | |

The R script automatically:
- Detects all `text_xx` / `help_xx` columns
- Generates the multi-language TSV rows LimeSurvey expects
- Falls back to the base language (or any available language) for untranslated cells
- Auto-sets the `additional_languages` survey setting

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

## Requirements

- **Excel** (or LibreOffice Calc) for editing the template
- **R** (≥ 4.0) with RStudio
- R packages (auto-installed on first run): `readxl`, `tidyxl`, `xml2`
- **LimeSurvey** (≥ 3.x) for import

## Configuration

In the R script, set the input filename (line 23):

```r
input_file <- "limesurvey_survey_builder.xlsx"
```

The output file is generated automatically with a `.txt` extension.

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
4. Run the R script — it auto-detects the new language columns

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

## Validation

The R script validates your survey before export:

- Checks for valid class values
- Detects duplicate question codes per language
- Warns about underscores in question codes
- Reports row counts by class type
- Lists advanced attributes in use

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Colors don't show in Excel | Conditional formatting uses `bgColor` — may need a restart of Excel |
| Import fails silently | Check that all question codes are alphanumeric (no underscores) |
| Translations don't appear | The script auto-generates all rows per language; ensure columns are named `text_xx` / `help_xx` |
| Special characters garbled | Output uses UTF-8 with BOM — import should auto-detect encoding |
| LimeSurvey renames codes | Codes with underscores get stripped; use alphanumeric only |

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests for:

- Additional question type examples
- New language translations for the example survey
- Bug fixes or improvements to the R script
- Documentation improvements

## Acknowledgments

- [LimeSurvey](https://www.limesurvey.org/) — the open-source survey platform this tool is built for
- Built with assistance from [Claude](https://claude.ai) by Anthropic
