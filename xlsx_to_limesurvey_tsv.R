# Title: Excel file to LimeSurvey TSV Builder
# Author: Stefan Savin
# Date: 2026-02-22
## ==============================================================================
## xlsx_to_limesurvey_tsv.R
## Convert LimeSurvey Survey Builder Excel file to LimeSurvey TSV import format
##
## MULTI-LANGUAGE SUPPORT:
##   Instead of duplicating rows per language, use columns:
##     text_en, help_en, text_fr, help_fr, text_de, help_de, ...
##   The script auto-detects these columns and generates the correct
##   multi-language TSV rows for LimeSurvey import.
##   Untranslated rows automatically fall back to the base language text.
##
## OTHER FEATURES:
##   - Reads partial rich text formatting within cells (mixed bold/color)
##   - Reads whole-cell formatting (entire cell bold/italic/colored)
##   - Converts in-cell line breaks (Alt+Enter) to HTML paragraphs or <br>
##   - Validates survey structure
##   - Outputs UTF-8 TSV with BOM
##
## Dependencies: readxl, tidyxl, xml2 (auto-installed if missing)
##
## NOTE: If you also use limesurvey_tsv_to_xlsx.R (which loads openxlsx2),
##       restart R before running this script. The xml2 and openxlsx2
##       packages both define xml_ns() and conflict with each other.
##       In RStudio: Session > Restart R (Ctrl+Shift+F10)
## ==============================================================================

# --- Set working directory to script location ---
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# --- Configuration ---
input_file  <- "limesurvey_survey_builder.xlsx"

# Output: same base name with .txt extension; adds _1, _2, ... if file exists
safe_filename <- function(base, ext) {
  candidate <- paste0(base, ".", ext)
  if (!file.exists(candidate)) return(candidate)
  i <- 1
  repeat {
    candidate <- paste0(base, "_", i, ".", ext)
    if (!file.exists(candidate)) return(candidate)
    i <- i + 1
  }
}
output_file <- safe_filename(sub("\\.[^.]+$", "", input_file), "txt")

# --- Dependencies ---
for (pkg in c("readxl", "tidyxl", "xml2")) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}
library(readxl)
library(tidyxl)
# Note: xml2 is no longer loaded directly; tidyxl imports it as needed

# ==============================================================================
# STEP 1: Read plain text data
# ==============================================================================
cat("Reading data:", input_file, "\n")
df <- read_excel(input_file, sheet = "Survey Design", col_types = "text", .name_repair = "minimal")
non_blank <- rowSums(!is.na(df) & df != "") > 0
# Map: df row index -> original Excel row (1-based, row 1 = header)
excel_row_map <- which(non_blank) + 1L  # +1 because Excel row 1 is header
df <- df[non_blank, ]
if (nrow(df) == 0) stop("No data found in 'Survey Design' sheet.")

col_names_vec <- names(df)
cat(sprintf("  Found %d data rows, %d columns\n", nrow(df), ncol(df)))

# ==============================================================================
# STEP 2: Detect language columns
# ==============================================================================
text_cols <- grep("^text_[a-zA-Z]{2}(-[a-zA-Z]{2,})?$", col_names_vec, value = TRUE)
help_cols <- grep("^help_[a-zA-Z]{2}(-[a-zA-Z]{2,})?$", col_names_vec, value = TRUE)

text_langs <- sub("^text_", "", text_cols)
help_langs <- sub("^help_", "", help_cols)
all_langs <- unique(c(text_langs, help_langs))

# Determine base language from the FIRST text_xx column (columns-driven).
# Additional languages are inferred from the remaining text_xx columns, in order.
# The 'language' and 'additional_languages' S rows are auto-populated from this
# detection, so they do not need to be maintained manually in the Excel sheet.
if (length(text_cols) > 0) {
  base_lang <- sub("^text_", "", text_cols[1])
  cat(sprintf("  Detected base language from first text column: %s\n", base_lang))
} else {
  base_lang <- "en"
  cat("  WARNING: No text_xx columns found. Defaulting base language to 'en'.\n")
}

# Other languages: remaining text_xx columns in column order, then any help-only langs
other_langs <- sub("^text_", "", text_cols[-1])
help_only <- setdiff(help_langs, c(base_lang, other_langs))
other_langs <- c(other_langs, help_only)
lang_order <- c(base_lang, other_langs)

# Auto-populate 'language' and 'additional_languages' S rows from detected columns.
# This overrides whatever is written there so the Excel sheet stays self-consistent.
base_lang_col <- paste0("text_", base_lang)
for (i in seq_len(nrow(df))) {
  if (is.na(df$class[i]) || df$class[i] != "S") next
  if (is.na(df$name[i])) next

  if (df$name[i] == "language") {
    old_val <- if (base_lang_col %in% names(df)) df[[base_lang_col]][i] else NA
    if (is.na(old_val) || old_val != base_lang) {
      if (base_lang_col %in% names(df)) {
        df[[base_lang_col]][i] <- base_lang
        cat(sprintf("  Auto-updated language: '%s' -> '%s'\n",
                    ifelse(is.na(old_val), "", old_val), base_lang))
      }
    }
  }

  if (df$name[i] == "additional_languages") {
    new_val <- paste(other_langs, collapse = " ")
    old_val <- if (base_lang_col %in% names(df)) df[[base_lang_col]][i] else NA
    if (is.na(old_val) || old_val != new_val) {
      if (base_lang_col %in% names(df)) {
        df[[base_lang_col]][i] <- new_val
        cat(sprintf("  Auto-updated additional_languages: '%s' -> '%s'\n",
                    ifelse(is.na(old_val), "", old_val), new_val))
      }
    }
  }
}

# If 'language' S row is missing entirely, inject it after the last S row
has_lang_row <- any(df$class == "S" & df$name == "language", na.rm = TRUE)
has_addl_row <- any(df$class == "S" & df$name == "additional_languages", na.rm = TRUE)

if (!has_lang_row || !has_addl_row) {
  # Find position of last S row
  s_positions <- which(df$class == "S")
  insert_after <- if (length(s_positions) > 0) max(s_positions) else 0

  new_rows <- list()
  if (!has_lang_row) {
    lang_row <- setNames(rep(NA_character_, ncol(df)), names(df))
    lang_row["class"] <- "S"
    lang_row["name"] <- "language"
    if (base_lang_col %in% names(df)) lang_row[base_lang_col] <- base_lang
    new_rows[[length(new_rows) + 1]] <- lang_row
    cat(sprintf("  Auto-created S row: language = '%s'\n", base_lang))
  }
  if (!has_addl_row && length(other_langs) > 0) {
    addl_row <- setNames(rep(NA_character_, ncol(df)), names(df))
    addl_row["class"] <- "S"
    addl_row["name"] <- "additional_languages"
    if (base_lang_col %in% names(df)) addl_row[base_lang_col] <- paste(other_langs, collapse = " ")
    new_rows[[length(new_rows) + 1]] <- addl_row
    cat(sprintf("  Auto-created S row: additional_languages = '%s'\n",
                paste(other_langs, collapse = " ")))
  }

  if (length(new_rows) > 0) {
    new_df <- as.data.frame(do.call(rbind, new_rows), stringsAsFactors = FALSE)
    names(new_df) <- names(df)
    na_map <- rep(NA_integer_, nrow(new_df))
    if (insert_after == 0) {
      df <- rbind(new_df, df)
      excel_row_map <- c(na_map, excel_row_map)
    } else if (insert_after == nrow(df)) {
      df <- rbind(df[1:insert_after, ], new_df)
      excel_row_map <- c(excel_row_map[1:insert_after], na_map)
    } else {
      df <- rbind(df[1:insert_after, ], new_df, df[(insert_after + 1):nrow(df), ])
      excel_row_map <- c(excel_row_map[1:insert_after], na_map,
                         excel_row_map[(insert_after + 1):length(excel_row_map)])
    }
  }
}

cat(sprintf("  Base language: %s\n", base_lang))
cat(sprintf("  All languages: %s\n", paste(lang_order, collapse = ", ")))
cat(sprintf("  Text columns: %s\n", paste(text_cols, collapse = ", ")))
cat(sprintf("  Help columns: %s\n", paste(help_cols, collapse = ", ")))

# ==============================================================================
# STEP 3: Parse rich text runs from xlsx XML
# ==============================================================================
cat("\nParsing rich text from xlsx...\n")

tmp_dir <- file.path(tempdir(), "xlsx_extract")
unlink(tmp_dir, recursive = TRUE)
dir.create(tmp_dir, showWarnings = FALSE)
unzip(input_file, exdir = tmp_dir, overwrite = TRUE)

# Helper: apply Excel tint/shade to a hex color (OOXML spec, HSL color space)
# Negative tint = darken (luminance toward 0), positive = lighten (toward 1)
apply_tint <- function(hex6, tint) {
  if (is.na(tint) || tint == 0) return(hex6)
  r <- strtoi(substr(hex6, 1, 2), 16)
  g <- strtoi(substr(hex6, 3, 4), 16)
  b <- strtoi(substr(hex6, 5, 6), 16)
  rn <- r / 255; gn <- g / 255; bn <- b / 255
  mx <- max(rn, gn, bn); mn <- min(rn, gn, bn)
  l <- (mx + mn) / 2
  if (mx == mn) { h <- 0; s <- 0 }
  else {
    d <- mx - mn
    s <- if (l > 0.5) d / (2 - mx - mn) else d / (mx + mn)
    h <- if (mx == rn) ((gn - bn) / d + (if (gn < bn) 6 else 0))
         else if (mx == gn) ((bn - rn) / d + 2)
         else ((rn - gn) / d + 4)
    h <- h / 6
  }
  if (tint < 0) l <- l * (1 + tint)
  else l <- l * (1 - tint) + tint
  l <- max(0, min(1, l))
  hue2rgb <- function(p, q, t) {
    if (t < 0) t <- t + 1; if (t > 1) t <- t - 1
    if (t < 1/6) return(p + (q - p) * 6 * t)
    if (t < 1/2) return(q)
    if (t < 2/3) return(p + (q - p) * (2/3 - t) * 6)
    return(p)
  }
  if (s == 0) { ro <- go <- bo <- l }
  else {
    q <- if (l < 0.5) l * (1 + s) else l + s - l * s
    p <- 2 * l - q
    ro <- hue2rgb(p, q, h + 1/3)
    go <- hue2rgb(p, q, h)
    bo <- hue2rgb(p, q, h - 1/3)
  }
  toupper(sprintf("%02X%02X%02X", round(ro * 255), round(go * 255), round(bo * 255)))
}

ss_path <- file.path(tmp_dir, "xl", "sharedStrings.xml")

# Parse theme colors for resolving theme="N" color references
theme_colors <- character(0)
theme_path <- file.path(tmp_dir, "xl", "theme", "theme1.xml")
if (file.exists(theme_path)) {
  theme_raw <- readLines(theme_path, encoding = "UTF-8", warn = FALSE)
  theme_str <- paste(theme_raw, collapse = "\n")
  # Extract clrScheme colors in order: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
  scheme_m <- regmatches(theme_str, regexpr('(?s)<a:clrScheme[^>]*>(.*?)</a:clrScheme>', theme_str, perl = TRUE))
  if (length(scheme_m) > 0) {
    named_blocks <- regmatches(scheme_m, gregexpr('(?s)<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>(.*?)</a:\\1>', scheme_m, perl = TRUE))[[1]]
    color_order <- c("lt1","dk1","lt2","dk2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink")
    palette <- setNames(rep("000000", 12), color_order)
    for (nb in named_blocks) {
      nm <- regmatches(nb, regexpr("^<a:(\\w+)>", nb))
      nm <- gsub("<a:|>", "", nm)
      rgb_m <- regmatches(nb, regexpr('srgbClr val="([A-Fa-f0-9]{6})"', nb))
      sys_m <- regmatches(nb, regexpr('lastClr="([A-Fa-f0-9]{6})"', nb))
      if (length(rgb_m) > 0) palette[nm] <- toupper(gsub('srgbClr val="|"', '', rgb_m))
      else if (length(sys_m) > 0) palette[nm] <- toupper(gsub('lastClr="|"', '', sys_m))
    }
    # OOXML theme index mapping: 0=dk1, 1=lt1, 2=dk2, 3=lt2, 4-9=accent1-6, 10=hlink, 11=folHlink
    theme_colors <- palette[color_order]
    cat(sprintf("  Theme colors loaded: %d\n", length(theme_colors)))
  }
}

resolve_theme_color <- function(rpr_str) {
  # Try rgb first
  cm <- regmatches(rpr_str, regexpr('rgb="([A-Fa-f0-9]{8})"', rpr_str))
  if (length(cm) > 0 && nchar(cm) > 0) {
    rgb_val <- gsub('rgb="|"', '', cm)
    hex6 <- toupper(substring(rgb_val, nchar(rgb_val) - 5, nchar(rgb_val)))
    tm <- regmatches(rpr_str, regexpr('tint="([^"]+)"', rpr_str))
    if (length(tm) > 0 && nchar(tm) > 0) {
      hex6 <- apply_tint(hex6, as.numeric(gsub('tint="|"', '', tm)))
    }
    return(hex6)
  }
  # Try theme color
  th <- regmatches(rpr_str, regexpr('theme="(\\d+)"', rpr_str))
  if (length(th) > 0 && length(theme_colors) > 0) {
    idx <- as.integer(gsub('theme="|"', '', th)) + 1  # 0-indexed → 1-indexed
    if (idx >= 1 && idx <= length(theme_colors)) {
      hex6 <- theme_colors[idx]
      tm <- regmatches(rpr_str, regexpr('tint="([^"]+)"', rpr_str))
      if (length(tm) > 0 && nchar(tm) > 0) {
        hex6 <- apply_tint(hex6, as.numeric(gsub('tint="|"', '', tm)))
      }
      return(hex6)
    }
  }
  return(NULL)
}

if (!file.exists(ss_path)) {
  cat("  No sharedStrings.xml (normal for programmatically generated xlsx).\n")
  cat("  Partial rich text step skipped; whole-cell formatting will be used instead.\n")
  rich_text_map <- list()
  fmt_col_names    <- c(text_cols, help_cols)
  fmt_col_positions <- match(fmt_col_names, col_names_vec)
} else {

# Parse sharedStrings.xml using regex (no xml2 dependency)
ss_raw <- readLines(ss_path, encoding = "UTF-8", warn = FALSE)
ss_str <- paste(ss_raw, collapse = "\n")

# Extract all <si>...</si> blocks
si_blocks <- regmatches(ss_str, gregexpr('(?s)<si>(.*?)</si>', ss_str, perl = TRUE))[[1]]

# Helper: parse formatting from an <rPr> block
rpr_to_html <- function(rpr_str, text) {
  if (is.na(text) || text == "") return(text)
  open_tags <- ""; close_tags <- ""
  if (nchar(rpr_str) > 0) {
    is_bold <- grepl('<b[ />]|<b>', rpr_str)
    is_italic <- grepl('<i[ />]|<i>', rpr_str)
    is_underline <- grepl('<u[ />]|<u>', rpr_str)
    color_hex <- NULL
    hex6 <- resolve_theme_color(rpr_str)
    if (!is.null(hex6) && hex6 != "000000" && hex6 != "FFFFFF" && hex6 != "") color_hex <- paste0("#", hex6)
    if (!is.null(color_hex)) {
      color_style <- paste0(" style='color:", color_hex, "'")
    } else {
      color_style <- ""
    }
    # Build tags: merge color style onto outermost formatting tag
    # so LimeSurvey doesn't strip nested spans
    if (is_bold || is_italic || is_underline) {
      # Outermost tag gets the color style attribute
      if (is_bold) {
        open_tags <- paste0(open_tags, "<strong", color_style, ">")
        close_tags <- paste0("</strong>", close_tags)
        color_style <- ""  # consumed
      }
      if (is_italic) {
        open_tags <- paste0(open_tags, "<em", color_style, ">")
        close_tags <- paste0("</em>", close_tags)
        color_style <- ""
      }
      if (is_underline) {
        open_tags <- paste0(open_tags, "<u", color_style, ">")
        close_tags <- paste0("</u>", close_tags)
        color_style <- ""
      }
    } else if (!is.null(color_hex)) {
      # Color only - use span
      open_tags <- paste0(open_tags, "<span style='color:", color_hex, "'>")
      close_tags <- paste0("</span>", close_tags)
    }
  }
  paste0(open_tags, text, close_tags)
}

rich_text_map <- list()
for (i in seq_along(si_blocks)) {
  si <- si_blocks[i]
  # Check for <r> runs (rich text)
  runs <- regmatches(si, gregexpr('(?s)<r>(.*?)</r>', si, perl = TRUE))[[1]]
  if (length(runs) > 0) {
    html_parts <- character(0)
    has_fmt <- FALSE
    for (run in runs) {
      # Extract text
      t_match <- regmatches(run, regexpr('(?s)<t[^>]*>(.*?)</t>', run, perl = TRUE))
      run_text <- if (length(t_match) > 0) gsub('<t[^>]*>|</t>', '', t_match) else ""
      # Decode XML entities
      run_text <- gsub("&amp;", "&", run_text)
      run_text <- gsub("&lt;", "<", run_text)
      run_text <- gsub("&gt;", ">", run_text)
      # Extract rPr
      rpr_match <- regmatches(run, regexpr('(?s)<rPr>(.*?)</rPr>', run, perl = TRUE))
      rpr_str <- if (length(rpr_match) > 0) rpr_match else ""
      html_run <- rpr_to_html(rpr_str, run_text)
      if (html_run != run_text) has_fmt <- TRUE
      html_parts <- c(html_parts, html_run)
    }
    if (has_fmt) rich_text_map[[as.character(i - 1)]] <- paste0(html_parts, collapse = "")
  }
}
cat(sprintf("  Found %d strings with partial rich text\n", length(rich_text_map)))

# Map rich text to cells
fmt_col_names <- c(text_cols, help_cols)
fmt_col_positions <- match(fmt_col_names, col_names_vec)

if (length(rich_text_map) > 0) {
  cat("Mapping rich text to cells...\n")

  # Find sheet path for Survey Design
  wb_path_ss <- file.path(tmp_dir, "xl", "workbook.xml")
  survey_idx <- 1
  if (file.exists(wb_path_ss)) {
    wb_raw_ss <- readLines(wb_path_ss, encoding = "UTF-8", warn = FALSE)
    wb_str_ss <- paste(wb_raw_ss, collapse = "\n")
    sn_all <- regmatches(wb_str_ss, gregexpr('name="([^"]+)"', wb_str_ss))[[1]]
    sn_all <- gsub('name="|"', '', sn_all)
    si_match <- which(sn_all == "Survey Design")
    if (length(si_match) > 0) survey_idx <- si_match[1]
  }

  sheet_path <- file.path(tmp_dir, "xl", "worksheets",
                          paste0("sheet", survey_idx, ".xml"))
  if (file.exists(sheet_path)) {
    sheet_raw_ss <- readLines(sheet_path, encoding = "UTF-8", warn = FALSE)
    sheet_str_ss <- paste(sheet_raw_ss, collapse = "\n")

    col_letter_to_num <- function(letters) {
      chars <- strsplit(letters, "")[[1]]
      num <- 0
      for (ch in chars) num <- num * 26 + (utf8ToInt(ch) - utf8ToInt("A") + 1)
      num
    }

    # Find cells with type="s" (shared string reference)
    cell_matches <- gregexpr('(?s)<c r="([A-Z]+)(\\d+)"[^>]*t="s"[^>]*>\\s*<v>(\\d+)</v>', sheet_str_ss, perl = TRUE)
    cell_strs <- regmatches(sheet_str_ss, cell_matches)[[1]]

    n_rich <- 0
    for (cs in cell_strs) {
      ref_m <- regmatches(cs, regexpr('r="([A-Z]+)(\\d+)"', cs))
      col_letter <- gsub('[^A-Z]', '', gsub('r="', '', ref_m))
      row_num <- as.integer(gsub('[^0-9]', '', ref_m))
      col_num <- col_letter_to_num(col_letter)
      if (!(col_num %in% fmt_col_positions)) next
      if (row_num <= 1) next
      v_m <- regmatches(cs, regexpr('<v>(\\d+)</v>', cs))
      ss_index <- gsub('<v>|</v>', '', v_m)
      if (ss_index %in% names(rich_text_map)) {
        data_row_candidates <- which(excel_row_map == row_num)
        if (length(data_row_candidates) == 0) next
        data_row <- data_row_candidates[1]
        col_name <- col_names_vec[col_num]
        if (!is.na(df$class[data_row]) && df$class[data_row] == "S") next
        df[[col_name]][data_row] <- rich_text_map[[ss_index]]
        n_rich <- n_rich + 1
      }
    }
    cat(sprintf("  Applied rich text to %d cells\n", n_rich))
  }
}
} # end else (sharedStrings.xml exists)

# ==============================================================================
# STEP 3b: Inline rich text from sheet XML (openxlsx2-generated files)
# ==============================================================================
# openxlsx2 stores rich text as inline strings (<is><r><rPr>...</rPr><t>...</t></r>)
# rather than shared strings. tidyxl doesn't read these. Parse them directly.
cat("\nChecking for inline rich text...\n")
n_inline <- 0

sheet_xml_path <- NULL
wb_xml_path <- file.path(tmp_dir, "xl", "workbook.xml")
if (file.exists(wb_xml_path)) {
  wb_raw <- readLines(wb_xml_path, encoding = "UTF-8", warn = FALSE)
  wb_str <- paste(wb_raw, collapse = "\n")
  # Extract sheet names from <sheet name="..." />
  sheet_names_all <- regmatches(wb_str,
    gregexpr('name="([^"]+)"', wb_str))[[1]]
  sheet_names_all <- gsub('name="|"', '', sheet_names_all)
  s_idx <- which(sheet_names_all == "Survey Design")
  if (length(s_idx) > 0) {
    sheet_xml_path <- file.path(tmp_dir, "xl", "worksheets",
                                paste0("sheet", s_idx[1], ".xml"))
    if (!file.exists(sheet_xml_path)) sheet_xml_path <- NULL
  }
}

if (!is.null(sheet_xml_path)) {
  sheet_raw <- readLines(sheet_xml_path, encoding = "UTF-8", warn = FALSE)
  sheet_str <- paste(sheet_raw, collapse = "\n")

  # Extract cells with inline rich text runs containing formatting
  # Pattern: <c r="XX" ...><is><r><rPr>...</rPr><t>...</t></r>...</is></c>
  inline_cells <- gregexpr(
    '<c r="([A-Z]+)(\\d+)"[^>]*>\\s*<is>(.*?)</is>\\s*</c>',
    sheet_str, perl = TRUE)
  matches <- regmatches(sheet_str, inline_cells)[[1]]

  if (length(matches) > 0) {
    for (m in matches) {
      # Skip cells without <rPr> containing actual formatting
      if (!grepl("<rPr>.*?(color|<b|<i|<u)", m)) next

      # Extract cell reference
      ref <- regmatches(m, regexpr('r="([A-Z]+)(\\d+)"', m))
      col_letter <- gsub('[^A-Z]', '', gsub('r="', '', ref))
      row_num    <- as.integer(gsub('[^0-9]', '', ref))

      col_num <- 0
      for (ch in strsplit(col_letter, "")[[1]])
        col_num <- col_num * 26 + (utf8ToInt(ch) - utf8ToInt("A") + 1)

      if (!(col_num %in% fmt_col_positions)) next
      if (row_num <= 1) next

      # Map excel row to df row using the maintained mapping
      data_row_candidates <- which(excel_row_map == row_num)
      if (length(data_row_candidates) == 0) next
      data_row <- data_row_candidates[1]
      if (data_row > nrow(df)) next

      col_name <- col_names_vec[col_num]
      if (!is.na(df$class[data_row]) && df$class[data_row] == "S") next

      # Parse <r> runs into HTML
      runs <- gregexpr('(?s)<r>(.*?)</r>', m, perl = TRUE)
      run_strs <- regmatches(m, runs)[[1]]

      html_parts <- character(0)
      has_fmt <- FALSE

      for (run_str in run_strs) {
        # Extract text
        t_match <- regmatches(run_str, regexpr('(?s)<t[^>]*>(.*?)</t>', run_str, perl = TRUE))
        run_text <- if (length(t_match) > 0) gsub('<t[^>]*>|</t>', '', t_match) else ""

        # Extract formatting
        rpr_match <- regmatches(run_str, regexpr('(?s)<rPr>(.*?)</rPr>', run_str, perl = TRUE))
        rpr <- if (length(rpr_match) > 0) rpr_match else ""

        open_tags <- ""
        close_tags <- ""

        if (nchar(rpr) > 0) {
          # Color (supports both rgb and theme colors)
          color_hex_inline <- NULL
          hex6_i <- resolve_theme_color(rpr)
          if (!is.null(hex6_i) && hex6_i != "000000" && hex6_i != "") color_hex_inline <- hex6_i
          # Formatting flags
          has_bold_i <- grepl('<b[ /]', rpr) || grepl('<b>', rpr)
          has_italic_i <- grepl('<i[ /]', rpr) || grepl('<i>', rpr)
          has_underline_i <- grepl('<u[ /]', rpr) || grepl('<u>', rpr)

          color_style <- if (!is.null(color_hex_inline)) paste0(" style='color:#", color_hex_inline, "'") else ""

          if (has_bold_i || has_italic_i || has_underline_i) {
            if (has_bold_i) {
              open_tags <- paste0(open_tags, "<strong", color_style, ">")
              close_tags <- paste0("</strong>", close_tags)
              color_style <- ""; has_fmt <- TRUE
            }
            if (has_italic_i) {
              open_tags <- paste0(open_tags, "<em", color_style, ">")
              close_tags <- paste0("</em>", close_tags)
              color_style <- ""; has_fmt <- TRUE
            }
            if (has_underline_i) {
              open_tags <- paste0(open_tags, "<u", color_style, ">")
              close_tags <- paste0("</u>", close_tags)
              color_style <- ""; has_fmt <- TRUE
            }
          } else if (!is.null(color_hex_inline)) {
            open_tags <- paste0(open_tags, "<span style='color:#", color_hex_inline, "'>")
            close_tags <- paste0("</span>", close_tags)
            has_fmt <- TRUE
          }
        }

        html_parts <- c(html_parts, paste0(open_tags, run_text, close_tags))
      }

      if (has_fmt && length(html_parts) > 0) {
        html_result <- paste0(html_parts, collapse = "")
        # Decode XML entities
        html_result <- gsub("&amp;", "&", html_result)
        html_result <- gsub("&lt;", "<", html_result)
        html_result <- gsub("&gt;", ">", html_result)
        df[[col_name]][data_row] <- html_result
        n_inline <- n_inline + 1
      }
    }
  }
}
cat(sprintf("  Inline rich text: %d cells converted\n", n_inline))

# ==============================================================================
# STEP 4: Whole-cell formatting (tidyxl)
# ==============================================================================
cat("\nChecking whole-cell formatting...\n")
tidyxl_ok <- FALSE

# tidyxl can segfault on xlsx files saved by openpyxl or similar tools.
# Run in a forked subprocess so a crash doesn't kill the main R process.
if (requireNamespace("parallel", quietly = TRUE) && .Platform$OS.type == "unix") {
  tmp_rds <- tempfile(fileext = ".rds")
  result <- parallel::mcparallel({
    tryCatch({
      ct <- tidyxl::xlsx_cells(input_file, sheets = "Survey Design")
      fm <- tidyxl::xlsx_formats(input_file)
      saveRDS(list(cells = ct, formats = fm), tmp_rds)
      TRUE
    }, error = function(e) FALSE)
  })
  collected <- parallel::mccollect(result, wait = TRUE, timeout = 30)
  if (!is.null(collected) && isTRUE(collected[[1]]) && file.exists(tmp_rds)) {
    td <- readRDS(tmp_rds)
    cells_tidy <- td$cells; formats <- td$formats; tidyxl_ok <- TRUE
  } else {
    cat("  WARNING: tidyxl crashed. Whole-cell formatting skipped.\n")
    cat("  (Normal for xlsx files created by openpyxl. Re-save in Excel/LibreOffice to fix.)\n")
  }
  unlink(tmp_rds)
} else {
  tryCatch({
    cells_tidy <- xlsx_cells(input_file, sheets = "Survey Design")
    formats <- xlsx_formats(input_file)
    tidyxl_ok <- TRUE
  }, error = function(e) {
    cat(sprintf("  WARNING: tidyxl failed (%s). Whole-cell formatting skipped.\n", e$message))
  })
}

n_whole <- 0
if (!tidyxl_ok) {
  cat("  Whole-cell formatting: skipped\n")
} else {
for (col_name in fmt_col_names) {
  col_pos <- match(col_name, col_names_vec)
  if (is.na(col_pos)) next
  for (row_i in seq_len(nrow(df))) {
    # Skip rows injected by auto-insertion (no Excel counterpart)
    if (is.na(excel_row_map[row_i])) next
    # Skip HTML conversion for S (settings) rows -- plain text only
    if (!is.na(df$class[row_i]) && df$class[row_i] == "S") {
      # Warn if cell has formatting that will be ignored
      cell_text_s <- df[[col_name]][row_i]
      if (!is.na(cell_text_s) && cell_text_s != "") {
        excel_row_s <- excel_row_map[row_i]
        cell_info_s <- cells_tidy[cells_tidy$row == excel_row_s & cells_tidy$col == col_pos, ]
        if (nrow(cell_info_s) > 0) {
          fmt_id_s <- cell_info_s$local_format_id[1]
          if (!is.na(fmt_id_s)) {
            has_fmt_s <- isTRUE(formats$local$font$bold[fmt_id_s]) ||
                         isTRUE(formats$local$font$italic[fmt_id_s]) ||
                         (!is.na(formats$local$font$color$rgb[fmt_id_s]) &&
                          formats$local$font$color$rgb[fmt_id_s] != "")
            if (has_fmt_s) {
              cat(sprintf("  WARNING: Rich text formatting on S row (name='%s', col=%s) -- ignored. Check for auto-formatted hyperlinks in LibreOffice.\n",
                          ifelse(is.na(df$name[row_i]), "", df$name[row_i]), col_name))
            }
          }
        }
      }
      next
    }
    cell_text <- df[[col_name]][row_i]
    if (is.na(cell_text) || cell_text == "") next
    has_html <- grepl("<[a-zA-Z][^>]*>", cell_text)
    if (is.na(excel_row_map[row_i])) next
    excel_row <- excel_row_map[row_i]
    cell_info <- cells_tidy[cells_tidy$row == excel_row & cells_tidy$col == col_pos, ]
    if (nrow(cell_info) == 0) next
    fmt_id <- cell_info$local_format_id[1]
    if (is.na(fmt_id)) next
    is_bold <- isTRUE(formats$local$font$bold[fmt_id])
    is_italic <- isTRUE(formats$local$font$italic[fmt_id])
    is_underline <- FALSE
    ul_val <- formats$local$font$underline[fmt_id]
    if (!is.na(ul_val) && ul_val != "" && ul_val != "none") is_underline <- TRUE
    color_hex <- NULL
    # Only detect color if cell doesn't already have color from inline parser
    if (!has_html || !grepl("color:", cell_text)) {
      font_color_rgb <- formats$local$font$color$rgb[fmt_id]
      font_color_tint <- formats$local$font$color$tint[fmt_id]
      if (!is.na(font_color_rgb) && font_color_rgb != "") {
        hex6 <- toupper(substring(font_color_rgb, 3, 8))
        if (!is.na(font_color_tint) && font_color_tint != 0) {
          hex6 <- apply_tint(hex6, font_color_tint)
        }
        if (hex6 != "000000" && hex6 != "FFFFFF" && hex6 != "" && nchar(hex6) == 6) color_hex <- paste0("#", hex6)
      }
    }
    # Skip bold/italic/underline if already present in HTML
    if (has_html) {
      if (grepl("<strong", cell_text)) is_bold <- FALSE
      if (grepl("<em", cell_text)) is_italic <- FALSE
      if (grepl("<u[ >]", cell_text)) is_underline <- FALSE
    }
    if (is_bold || is_italic || is_underline || !is.null(color_hex)) {
      result <- cell_text
      color_style <- if (!is.null(color_hex)) paste0(" style='color:", color_hex, "'") else ""
      if (is_bold || is_italic || is_underline) {
        if (is_underline) { result <- paste0("<u", color_style, ">", result, "</u>"); color_style <- "" }
        if (is_italic) { result <- paste0("<em", color_style, ">", result, "</em>"); color_style <- "" }
        if (is_bold) { result <- paste0("<strong", color_style, ">", result, "</strong>"); color_style <- "" }
      } else if (!is.null(color_hex)) {
        result <- paste0("<span style='color:", color_hex, "'>", result, "</span>")
      }
      if (result != cell_text) {
        df[[col_name]][row_i] <- result
        n_whole <- n_whole + 1
      }
    }
  }
}
cat(sprintf("Whole-cell formatting: %d cells converted\n", n_whole))
} # end if (tidyxl_ok)

# ==============================================================================
# STEP 5: Line breaks and <p> wrapping
# ==============================================================================
cat("\nConverting line breaks and wrapping HTML...\n")

has_block_tag <- function(txt) {
  grepl("^\\s*<(p|div|h[1-6]|ul|ol|table)", txt, ignore.case = TRUE)
}

sl_wrap_fields <- c("surveyls_description", "surveyls_welcometext", "surveyls_endtext")
n_wrapped <- 0

for (col_name in fmt_col_names) {
  for (row_i in seq_len(nrow(df))) {
    cls <- df$class[row_i]
    if (is.na(cls)) next
    is_text_col <- grepl("^text_", col_name)
    is_help_col <- grepl("^help_", col_name)
    txt <- df[[col_name]][row_i]
    if (is.na(txt) || txt == "" || has_block_tag(txt)) next
    has_nl <- grepl("\n", txt, fixed = TRUE)

    if (is_text_col) {
      mode <- "none"
      if (cls %in% c("Q", "G")) {
        mode <- "p"
      } else if (cls == "SL" && !is.na(df$name[row_i]) &&
                 df$name[row_i] %in% sl_wrap_fields) {
        mode <- "p"
      } else if (cls %in% c("SQ", "A") && has_nl) {
        mode <- "br"
      }
      if (mode == "p") {
        if (has_nl) {
          lines <- strsplit(txt, "\n")[[1]]
          lines <- trimws(lines)
          lines <- lines[lines != ""]
          df[[col_name]][row_i] <- paste0("<p>", lines, "</p>", collapse = "")
        } else {
          df[[col_name]][row_i] <- paste0("<p>", txt, "</p>")
        }
        n_wrapped <- n_wrapped + 1
      } else if (mode == "br") {
        df[[col_name]][row_i] <- gsub("\n", "<br>", txt)
        n_wrapped <- n_wrapped + 1
      }
    } else if (is_help_col && cls %in% c("Q", "G")) {
      if (has_nl) {
        lines <- strsplit(hlp, "\n")[[1]]
        lines <- trimws(lines)
        lines <- lines[lines != ""]
        df[[col_name]][row_i] <- paste0("<p>", lines, "</p>", collapse = "")
      } else {
        df[[col_name]][row_i] <- paste0("<p>", txt, "</p>")
      }
      n_wrapped <- n_wrapped + 1
    }
  }
}
cat(sprintf("HTML wrapping: %d fields processed\n", n_wrapped))

# ==============================================================================
# STEP 6: Expand rows into multi-language TSV format
# ==============================================================================
# LimeSurvey requires ALL G/Q/SQ/A rows for EVERY language declared.
# If text_xx is empty for a row, we fall back to the base language text.
# This ensures LimeSurvey receives a complete set of rows per language.
# ==============================================================================

cat("\nExpanding to multi-language TSV...\n")

shared_cols <- setdiff(col_names_vec, c(text_cols, help_cols))

standard_shared <- c("id", "related_id", "class", "type/scale", "name",
                      "relevance", "validation", "mandatory", "other",
                      "default", "same_default")
adv_cols <- setdiff(shared_cols, standard_shared)
adv_cols_used <- character(0)
for (col in adv_cols) {
  if (any(!is.na(df[[col]]) & df[[col]] != "")) {
    adv_cols_used <- c(adv_cols_used, col)
  }
}
adv_cols_used <- sort(adv_cols_used)

tsv_cols <- c("id", "related_id", "class", "type/scale", "name",
              "relevance", "text", "help", "language",
              "validation", "mandatory", "other", "default", "same_default",
              adv_cols_used)

if (length(adv_cols_used) > 0) {
  cat("Advanced attributes:", paste(adv_cols_used, collapse = ", "), "\n")
}

# Helper: get text for a language with cascading fallback
# Tries: requested language -> base language -> any other language with content
get_text <- function(row_i, lang, type = "text") {
  col_lang <- paste0(type, "_", lang)
  val <- NA_character_
  if (col_lang %in% names(df)) val <- df[[col_lang]][row_i]
  if (!is.na(val) && val != "") return(val)

  # Fallback 1: base language
  if (lang != base_lang) {
    col_base <- paste0(type, "_", base_lang)
    if (col_base %in% names(df)) val <- df[[col_base]][row_i]
    if (!is.na(val) && val != "") return(val)
  }

  # Fallback 2: first available language that has content
  for (fl in lang_order) {
    if (fl == lang || fl == base_lang) next
    col_fl <- paste0(type, "_", fl)
    if (col_fl %in% names(df)) val <- df[[col_fl]][row_i]
    if (!is.na(val) && val != "") return(val)
  }

  return("")
}

out_rows <- list()
n_fallback <- 0

for (row_i in seq_len(nrow(df))) {
  cls <- df$class[row_i]
  if (is.na(cls)) next

  shared <- list()
  for (col in c(standard_shared, adv_cols_used)) {
    shared[[col]] <- if (col %in% names(df)) {
      val <- df[[col]][row_i]; if (is.na(val)) "" else val
    } else ""
  }

  if (cls == "S") {
    # S rows: single row, text from base language, no language column
    r <- shared
    r[["text"]] <- get_text(row_i, base_lang, "text")
    r[["help"]] <- ""
    r[["language"]] <- ""
    out_rows[[length(out_rows) + 1]] <- r

  } else if (cls == "SL") {
    # SL rows: one row per language
    for (lang in lang_order) {
      txt <- get_text(row_i, lang, "text")
      hlp <- get_text(row_i, lang, "help")
      if (txt == "" && hlp == "") next
      r <- shared
      r[["text"]] <- txt
      r[["help"]] <- hlp
      r[["language"]] <- lang
      out_rows[[length(out_rows) + 1]] <- r
    }

  } else {
    # G, Q, SQ, A: base language row
    r <- shared
    r[["text"]] <- get_text(row_i, base_lang, "text")
    r[["help"]] <- get_text(row_i, base_lang, "help")
    r[["language"]] <- base_lang
    out_rows[[length(out_rows) + 1]] <- r
  }
}

# Second pass: additional language rows for ALL G/Q/SQ/A
for (lang in other_langs) {
  text_col_lang <- paste0("text_", lang)
  for (row_i in seq_len(nrow(df))) {
    cls <- df$class[row_i]
    if (is.na(cls) || cls %in% c("S", "SL")) next

    # Get translated text, fall back to base language if empty
    txt <- get_text(row_i, lang, "text")
    hlp <- get_text(row_i, lang, "help")

    # Track fallbacks
    lang_col <- paste0("text_", lang)
    lang_val <- if (lang_col %in% names(df)) df[[lang_col]][row_i] else NA_character_
    if (is.na(lang_val) || lang_val == "") n_fallback <- n_fallback + 1

    shared <- list()
    for (col in c(standard_shared, adv_cols_used)) {
      shared[[col]] <- if (col %in% names(df)) {
        val <- df[[col]][row_i]; if (is.na(val)) "" else val
      } else ""
    }

    r <- shared
    r[["text"]] <- txt
    r[["help"]] <- hlp
    r[["language"]] <- lang
    # Clear attributes for translated rows (LimeSurvey uses base language only)
    for (k in c("relevance", "validation", "mandatory", "default", "same_default")) {
      r[[k]] <- ""
    }
    for (adv in adv_cols_used) r[[adv]] <- ""

    out_rows[[length(out_rows) + 1]] <- r
  }
}

# Convert to data frame
df_out <- as.data.frame(
  do.call(rbind, lapply(out_rows, function(r) {
    sapply(tsv_cols, function(col) {
      val <- r[[col]]
      if (is.null(val) || is.na(val)) "" else as.character(val)
    })
  })),
  stringsAsFactors = FALSE
)
names(df_out) <- tsv_cols

# ==============================================================================
# STEP 6b: Read Quotas sheet and generate QTA/QTALS/QTAM rows
# ==============================================================================
# Flat format: one row per quota in the Quotas sheet.
# Columns: quota_name, quota_limit, active, quota_action, autoload_url,
#          message_<lang>..., question_code_1, answer_code_1, question_code_2, ...
#
# LimeSurvey TSV mapping (from manual):
#   QTA:   mandatory=limit, other=action, default=active, same_default=autoload_url
#   QTALS: relevance=message, text=url, help=url_description, language=lang
#   QTAM:  name=answer_code  (placed after the question it relates to)
# ==============================================================================
cat("\nChecking for Quotas sheet...\n")

quota_sheets <- excel_sheets(input_file)
if ("Quotas" %in% quota_sheets) {
  qdf <- read_excel(input_file, sheet = "Quotas", col_types = "text", .name_repair = "minimal")
  qdf <- qdf[rowSums(!is.na(qdf) & qdf != "") > 0, ]

  if (nrow(qdf) > 0 && "quota_name" %in% names(qdf)) {
    qdf <- qdf[!is.na(qdf$quota_name) & qdf$quota_name != "", ]
  }

  if (nrow(qdf) > 0) {
    cat(sprintf("  Found %d quotas in Quotas sheet\n", nrow(qdf)))

    tval <- function(x) {
      if (is.null(x)) return("")
      v <- x[[1]]
      if (is.na(v) || v == "") return("") else return(as.character(v))
    }

    # Detect question_code_N / answer_code_N column pairs
    qc_cols <- grep("^question_code_\\d+$", names(qdf), value = TRUE)
    member_nums <- sort(unique(as.integer(sub("^question_code_", "", qc_cols))))
    cat(sprintf("  Member column pairs detected: %d\n", length(member_nums)))

    quota_id_counter <- 0
    end_rows <- list()
    qtam_by_question <- list()

    for (row_i in seq_len(nrow(qdf))) {
      r <- qdf[row_i, ]
      qname <- tval(r$quota_name)
      if (qname == "") next

      # QTA row
      quota_id_counter <- quota_id_counter + 1
      qta_id <- quota_id_counter

      qta_row <- setNames(rep("", length(tsv_cols)), tsv_cols)
      qta_row["id"]           <- as.character(qta_id)
      qta_row["class"]        <- "QTA"
      qta_row["name"]         <- qname
      qta_row["mandatory"]    <- tval(r$quota_limit)
      qta_row["other"]        <- if (tval(r$quota_action) != "") tval(r$quota_action) else "1"
      active_val <- toupper(tval(r$active))
      qta_row["default"]      <- if (active_val %in% c("Y", "1")) "1" else "0"
      qta_row["same_default"] <- if (tval(r$autoload_url) != "") tval(r$autoload_url) else "0"
      end_rows[[length(end_rows) + 1]] <- qta_row

      # QTALS rows: one per language
      for (lang in lang_order) {
        quota_id_counter <- quota_id_counter + 1
        qtals_row <- setNames(rep("", length(tsv_cols)), tsv_cols)
        qtals_row["id"]         <- as.character(quota_id_counter)
        qtals_row["related_id"] <- as.character(qta_id)
        qtals_row["class"]      <- "QTALS"
        msg <- ""
        msg_col <- paste0("message_", lang)
        if (msg_col %in% names(r)) {
          msg <- tval(r[[msg_col]])
          if (msg == "") {
            msg_base <- paste0("message_", base_lang)
            if (msg_base %in% names(r)) msg <- tval(r[[msg_base]])
          }
        }
        qtals_row["relevance"] <- msg
        qtals_row["language"]  <- lang
        end_rows[[length(end_rows) + 1]] <- qtals_row
      }

      # QTAM rows: one per question_code_N / answer_code_N pair
      for (n in member_nums) {
        qc_col <- paste0("question_code_", n)
        ac_col <- paste0("answer_code_", n)
        qcode <- if (qc_col %in% names(r)) tval(r[[qc_col]]) else ""
        acode <- if (ac_col %in% names(r)) tval(r[[ac_col]]) else ""
        if (qcode == "" || acode == "") next

        quota_id_counter <- quota_id_counter + 1
        qtam_row <- setNames(rep("", length(tsv_cols)), tsv_cols)
        qtam_row["id"]         <- as.character(quota_id_counter)
        qtam_row["related_id"] <- as.character(qta_id)
        qtam_row["class"]      <- "QTAM"
        qtam_row["name"]       <- acode

        if (is.null(qtam_by_question[[qcode]])) qtam_by_question[[qcode]] <- list()
        qtam_by_question[[qcode]][[length(qtam_by_question[[qcode]]) + 1]] <- qtam_row
      }
    }

    # Inject QTAM rows after each referenced question's SQ/A block
    if (length(qtam_by_question) > 0) {
      insert_positions <- list()
      for (qcode in names(qtam_by_question)) {
        q_pos <- which(df_out$class == "Q" & df_out$name == qcode & df_out$language == base_lang)
        if (length(q_pos) == 0) {
          warning(sprintf("Quota references question '%s' not found in survey", qcode))
          next
        }
        q_pos <- q_pos[1]
        insert_after <- q_pos
        if (q_pos < nrow(df_out)) {
          for (j in (q_pos + 1):nrow(df_out)) {
            if (df_out$class[j] %in% c("G", "Q", "S", "SL", "QTA", "QTALS")) break
            if (df_out$class[j] %in% c("SQ", "A")) {
              insert_after <- j
            } else {
              break
            }
          }
        }
        insert_positions[[length(insert_positions) + 1]] <- list(
          pos = insert_after, rows = qtam_by_question[[qcode]])
      }

      positions <- sapply(insert_positions, function(x) x$pos)
      insert_positions <- insert_positions[order(positions, decreasing = TRUE)]

      for (ip in insert_positions) {
        qtam_df <- as.data.frame(do.call(rbind, ip$rows), stringsAsFactors = FALSE)
        names(qtam_df) <- tsv_cols
        if (ip$pos == nrow(df_out)) {
          df_out <- rbind(df_out, qtam_df)
        } else {
          df_out <- rbind(df_out[1:ip$pos, ], qtam_df, df_out[(ip$pos + 1):nrow(df_out), ])
        }
      }
      n_qtam <- sum(sapply(insert_positions, function(x) length(x$rows)))
      cat(sprintf("  Inserted %d QTAM rows after their respective questions\n", n_qtam))
    }

    # Append QTA + QTALS at end of file
    if (length(end_rows) > 0) {
      end_df <- as.data.frame(do.call(rbind, end_rows), stringsAsFactors = FALSE)
      names(end_df) <- tsv_cols
      df_out <- rbind(df_out, end_df)
      n_qta <- sum(sapply(end_rows, function(r) r["class"] == "QTA"))
      n_qtals <- sum(sapply(end_rows, function(r) r["class"] == "QTALS"))
      cat(sprintf("  Appended %d QTA + %d QTALS rows at end of file\n", n_qta, n_qtals))
    }

  } else {
    cat("  Quotas sheet is empty -- skipping.\n")
  }
} else {
  cat("  No Quotas sheet found -- skipping.\n")
}

# Count by language
lang_counts <- table(df_out$language[df_out$language != ""])
cat(sprintf("TSV rows: %d total\n", nrow(df_out)))
for (l in names(lang_counts)) {
  cat(sprintf("  %s: %d rows\n", l, lang_counts[[l]]))
}
if (n_fallback > 0) {
  cat(sprintf("  (%d rows used base language as fallback for missing translations)\n", n_fallback))
}

# ==============================================================================
# STEP 7: Validation
# ==============================================================================
cat("\n--- Validation ---\n")
valid_classes <- c("S","SL","G","Q","SQ","A","AS","QTA","QTALS","QTAM","C")
invalid <- setdiff(unique(df_out$class[df_out$class != ""]), valid_classes)
if (length(invalid) > 0) warning("Invalid class values: ", paste(invalid, collapse = ", "))

for (cls in c("S","SL","G","Q","SQ","A","QTA","QTALS","QTAM")) {
  cat(sprintf("  %-6s %d rows\n", paste0(cls, ":"), sum(df_out$class == cls)))
}

q_rows <- df_out[df_out$class == "Q", ]
if (nrow(q_rows) > 0) {
  q_check <- q_rows[q_rows$name != "" & q_rows$language != "", c("name","language")]
  for (lang in unique(q_check$language)) {
    dupes <- q_check$name[q_check$language == lang]
    dupes <- dupes[duplicated(dupes)]
    if (length(dupes) > 0) {
      warning("Duplicate question names in '", lang, "': ",
              paste(unique(dupes), collapse = ", "))
    }
  }
  q_with_underscore <- q_rows$name[grepl("_", q_rows$name, fixed = TRUE)]
  if (length(q_with_underscore) > 0) {
    warning("Question codes with underscores (LimeSurvey will strip them!): ",
            paste(unique(q_with_underscore), collapse = ", "))
  }
}

# ==============================================================================
# STEP 8: Write TSV with UTF-8 BOM
# ==============================================================================
cat("\n--- Writing output ---\n")

# TSV quoting: fields containing tabs, newlines, or double quotes
# must be wrapped in double quotes, with inner quotes doubled.
quote_tsv_field <- function(x) {
  if (is.na(x)) return("")
  if (grepl("[\t\n\"]", x)) {
    return(paste0("\"", gsub("\"", "\"\"", x), "\""))
  }
  x
}

header_line <- paste(tsv_cols, collapse = "\t")
data_lines  <- apply(df_out, 1, function(row) {
  paste(sapply(row, quote_tsv_field), collapse = "\t")
})
tsv_content <- paste(c(header_line, data_lines), collapse = "\n")

con <- file(output_file, open = "wb")
writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
writeBin(charToRaw(tsv_content), con)
close(con)

cat(sprintf("Output:  %s\n", output_file))
cat(sprintf("Rows:    %d data + 1 header\n", nrow(df_out)))
cat(sprintf("Columns: %d (%d standard + %d advanced)\n",
            length(tsv_cols), 14, length(adv_cols_used)))
cat(sprintf("Languages: %s\n", paste(lang_order, collapse = ", ")))
cat("\n--- Import into LimeSurvey ---\n")
cat("Create Survey > Import > select", output_file, "\n")
cat("Done!\n")

# Cleanup
unlink(tmp_dir, recursive = TRUE)
