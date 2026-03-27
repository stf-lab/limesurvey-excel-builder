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
library(xml2)

# ==============================================================================
# STEP 1: Read plain text data
# ==============================================================================
cat("Reading data:", input_file, "\n")
df <- read_excel(input_file, sheet = "Survey Design", col_types = "text", .name_repair = "minimal")
df <- df[rowSums(!is.na(df) & df != "") > 0, ]
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
    if (insert_after == 0) {
      df <- rbind(new_df, df)
    } else if (insert_after == nrow(df)) {
      df <- rbind(df[1:insert_after, ], new_df)
    } else {
      df <- rbind(df[1:insert_after, ], new_df, df[(insert_after + 1):nrow(df), ])
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
if (!file.exists(ss_path)) {
  cat("  No sharedStrings.xml (normal for programmatically generated xlsx).\n")
  cat("  Partial rich text step skipped; whole-cell formatting will be used instead.\n")
  rich_text_map <- list()
  fmt_col_names    <- c(text_cols, help_cols)
  fmt_col_positions <- match(fmt_col_names, col_names_vec)
} else {

ss_xml <- read_xml(ss_path)
ns <- xml_ns(ss_xml)
si_nodes <- xml_find_all(ss_xml, ".//d1:si", ns)

run_to_html <- function(text, rpr_node, ns) {
  if (is.na(text) || text == "") return(text)
  open_tags <- ""
  close_tags <- ""
  if (!is.null(rpr_node)) {
    is_bold <- length(xml_find_all(rpr_node, ".//d1:b", ns)) > 0
    is_italic <- length(xml_find_all(rpr_node, ".//d1:i", ns)) > 0
    is_underline <- length(xml_find_all(rpr_node, ".//d1:u", ns)) > 0
    color_node <- xml_find_first(rpr_node, ".//d1:color", ns)
    color_hex <- NULL
    if (!is.na(color_node)) {
      rgb_val <- xml_attr(color_node, "rgb")
      tint_val <- xml_attr(color_node, "tint")
      if (!is.na(rgb_val) && nchar(rgb_val) >= 6) {
        hex6 <- toupper(substring(rgb_val, nchar(rgb_val) - 5, nchar(rgb_val)))
        # Apply tint if present (theme colors with shade/tint modification)
        if (!is.na(tint_val) && tint_val != "") {
          hex6 <- apply_tint(hex6, as.numeric(tint_val))
        }
        if (hex6 != "000000" && hex6 != "") color_hex <- paste0("#", hex6)
      }
    }
    if (!is.null(color_hex)) {
      open_tags <- paste0(open_tags, "<span style='color:", color_hex, "'>")
      close_tags <- paste0("</span>", close_tags)
    }
    if (is_underline) {
      open_tags <- paste0(open_tags, "<u>"); close_tags <- paste0("</u>", close_tags)
    }
    if (is_italic) {
      open_tags <- paste0(open_tags, "<em>"); close_tags <- paste0("</em>", close_tags)
    }
    if (is_bold) {
      open_tags <- paste0(open_tags, "<strong>"); close_tags <- paste0("</strong>", close_tags)
    }
  }
  paste0(open_tags, text, close_tags)
}

rich_text_map <- list()
for (i in seq_along(si_nodes)) {
  si <- si_nodes[[i]]
  r_nodes <- xml_find_all(si, "./d1:r", ns)
  if (length(r_nodes) > 0) {
    html_parts <- character(0)
    has_fmt <- FALSE
    for (r_node in r_nodes) {
      t_node <- xml_find_first(r_node, "./d1:t", ns)
      rpr_node <- xml_find_first(r_node, "./d1:rPr", ns)
      run_text <- if (!is.na(t_node)) xml_text(t_node) else ""
      html_run <- run_to_html(run_text, if (is.na(rpr_node)) NULL else rpr_node, ns)
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

  wb_xml <- read_xml(file.path(tmp_dir, "xl", "workbook.xml"))
  wb_ns <- xml_ns(wb_xml)
  sheet_nodes <- xml_find_all(wb_xml, ".//d1:sheets/d1:sheet", wb_ns)
  sheet_names <- xml_attr(sheet_nodes, "name")
  survey_idx <- which(sheet_names == "Survey Design")

  sheet_path <- file.path(tmp_dir, "xl", "worksheets",
                          paste0("sheet", survey_idx, ".xml"))
  if (file.exists(sheet_path)) {
    sheet_xml <- read_xml(sheet_path)
    sheet_ns <- xml_ns(sheet_xml)
    all_cells <- xml_find_all(sheet_xml, ".//d1:c", sheet_ns)

    col_letter_to_num <- function(letters) {
      chars <- strsplit(letters, "")[[1]]
      num <- 0
      for (ch in chars) num <- num * 26 + (utf8ToInt(ch) - utf8ToInt("A") + 1)
      num
    }

    n_rich <- 0
    for (cell_node in all_cells) {
      cell_type <- xml_attr(cell_node, "t")
      if (is.na(cell_type) || cell_type != "s") next
      cell_ref <- xml_attr(cell_node, "r")
      col_letter <- gsub("[0-9]", "", cell_ref)
      row_num <- as.integer(gsub("[A-Z]", "", cell_ref))
      col_num <- col_letter_to_num(col_letter)
      if (!(col_num %in% fmt_col_positions)) next
      if (row_num <= 1) next
      v_node <- xml_find_first(cell_node, "./d1:v", sheet_ns)
      if (is.na(v_node)) next
      ss_index <- xml_text(v_node)
      if (ss_index %in% names(rich_text_map)) {
        data_row <- row_num - 1
        col_name <- col_names_vec[col_num]
        # Skip HTML conversion for S (settings) rows -- plain text only
        if (!is.na(df$class[data_row]) && df$class[data_row] == "S") {
          cat(sprintf("  [%s] Row %d: S row -- rich text formatting ignored\n", col_name, row_num))
          next
        }
        df[[col_name]][data_row] <- rich_text_map[[ss_index]]
        n_rich <- n_rich + 1
        cat(sprintf("  [%s] Row %d: rich text applied\n", col_name, row_num))
      }
    }
    cat(sprintf("  Applied rich text to %d cells\n", n_rich))
  }
}
} # end else (sharedStrings.xml exists)

# ==============================================================================
# STEP 4: Whole-cell formatting (tidyxl)
# ==============================================================================
cat("\nChecking whole-cell formatting...\n")
cells_tidy <- xlsx_cells(input_file, sheets = "Survey Design")
formats <- xlsx_formats(input_file)

n_whole <- 0
for (col_name in fmt_col_names) {
  col_pos <- match(col_name, col_names_vec)
  if (is.na(col_pos)) next
  for (row_i in seq_len(nrow(df))) {
    # Skip HTML conversion for S (settings) rows -- plain text only
    if (!is.na(df$class[row_i]) && df$class[row_i] == "S") {
      # Warn if cell has formatting that will be ignored
      cell_text_s <- df[[col_name]][row_i]
      if (!is.na(cell_text_s) && cell_text_s != "") {
        excel_row_s <- row_i + 1
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
    if (grepl("<[a-zA-Z][^>]*>", cell_text)) next
    excel_row <- row_i + 1
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
    font_color_rgb <- formats$local$font$color$rgb[fmt_id]
    font_color_tint <- formats$local$font$color$tint[fmt_id]
    if (!is.na(font_color_rgb) && font_color_rgb != "") {
      hex6 <- toupper(substring(font_color_rgb, 3, 8))
      # Apply tint if present (theme colors with shade/tint modification)
      if (!is.na(font_color_tint) && font_color_tint != 0) {
        hex6 <- apply_tint(hex6, font_color_tint)
      }
      if (hex6 != "000000" && hex6 != "" && nchar(hex6) == 6) color_hex <- paste0("#", hex6)
    }
    if (is_bold || is_italic || is_underline || !is.null(color_hex)) {
      result <- cell_text
      if (!is.null(color_hex)) result <- paste0("<span style='color:", color_hex, "'>", result, "</span>")
      if (is_underline) result <- paste0("<u>", result, "</u>")
      if (is_italic) result <- paste0("<em>", result, "</em>")
      if (is_bold) result <- paste0("<strong>", result, "</strong>")
      if (result != cell_text) {
        df[[col_name]][row_i] <- result
        n_whole <- n_whole + 1
      }
    }
  }
}
cat(sprintf("Whole-cell formatting: %d cells converted\n", n_whole))

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

for (cls in c("S","SL","G","Q","SQ","A")) {
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
