# Title: LimeSurvey TSV to Excel Builder (Inverse Converter)
# Author: Stefan Savin
# Date: 2026-03-27
## ==============================================================================
## limesurvey_tsv_to_xlsx.R
## Convert a LimeSurvey TSV export (.txt) back into the Excel Builder format
## (.xlsx) so it can be edited in Excel and re-exported with xlsx_to_limesurvey_tsv.R.
##
## WHAT IT DOES:
##   - Reads a LimeSurvey TSV export file (.txt) with multiple languages
##   - Collapses multi-language rows into side-by-side text_xx / help_xx columns
##   - Drops server-specific settings and empty attribute columns
##   - Converts HTML formatting (bold, italic, underline, color) to Excel rich text
##   - Produces an .xlsx file with the same structure as the master template:
##       conditional formatting, data validations, frozen header, reference sheets
##
## USAGE:
##   1. In LimeSurvey: Surveys > Export > Tab Separated Values (.txt)
##   2. Set input_file below to the exported .txt filename
##   3. Run this script in RStudio
##   4. Open the resulting .xlsx in Excel, edit, then run xlsx_to_limesurvey_tsv.R
##
## Dependencies: openxlsx2 (auto-installed if missing)
##
## NOTE: If you also use xlsx_to_limesurvey_tsv.R (which loads xml2),
##       restart R before running this script. The xml2 and openxlsx2
##       packages both define xml_ns() and conflict with each other.
##       In RStudio: Session > Restart R (Ctrl+Shift+F10)
## ==============================================================================

# --- Set working directory to script location ---
if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
} else {
  args <- commandArgs(trailingOnly = FALSE)
  script_arg <- grep("--file=", args, value = TRUE)
  if (length(script_arg) > 0) {
    setwd(dirname(normalizePath(sub("--file=", "", script_arg))))
  }
}

# --- Configuration ---
input_file  <- "limesurvey_survey_builder.txt"

# Output: same base name with .xlsx extension; adds _1, _2, ... if file exists
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
output_file <- safe_filename(sub("\\.[^.]+$", "", input_file), "xlsx")

# --- Dependencies ---
for (pkg in c("openxlsx2")) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}
library(openxlsx2)

# ==============================================================================
# STEP 1: Read TSV file
# ==============================================================================
cat("Reading TSV:", input_file, "\n")
raw <- readLines(input_file, encoding = "UTF-8", warn = FALSE)
if (length(raw) > 0 && grepl("^\uFEFF", raw[1])) raw[1] <- sub("^\uFEFF", "", raw[1])
tsv <- read.delim(textConnection(paste(raw, collapse = "\n")),
                  sep = "\t", header = TRUE, quote = "\"",
                  stringsAsFactors = FALSE, check.names = FALSE,
                  colClasses = "character", na.strings = character(0))
tsv[is.na(tsv)] <- ""
tsv[is.na(tsv)] <- ""
cat(sprintf("  Rows: %d, Columns: %d\n", nrow(tsv), ncol(tsv)))

# ==============================================================================
# STEP 2: Identify languages and base language
# ==============================================================================
# Find base language from S row name=language
base_lang <- "en"
for (i in seq_len(nrow(tsv))) {
  if (tsv$class[i] == "S" && tsv$name[i] == "language" && tsv$text[i] != "") {
    base_lang <- tsv$text[i]
    break
  }
}

# Collect all languages from the language column
all_langs <- sort(unique(tsv$language[tsv$language != "" & !is.na(tsv$language)]))
other_langs <- setdiff(all_langs, base_lang)
lang_order <- c(base_lang, other_langs)

cat(sprintf("  Base language: %s\n", base_lang))
cat(sprintf("  All languages: %s\n", paste(lang_order, collapse = ", ")))

# ==============================================================================
# STEP 3: Filter S rows -- keep essential settings, drop server-specific ones
# ==============================================================================
cat("\nFiltering settings...\n")

# Settings to DROP (server-specific, timestamps, tokens, etc.)
drop_s <- c("gsid", "template", "savetimings", "lastmodified",
            "tokenencryptionoptions", "access_mode",
            "usecookie", "allowregister", "autonumber_start",
            "autoredirect", "ipanonymize", "refurl",
            "showsurveypolicynotice", "publicstatistics", "publicgraphs",
            "listpublic", "htmlemail", "sendconfirmation",
            "tokenanswerspersistence", "assessments", "usecaptcha",
            "usetokens", "bounce_email", "tokenlength",
            "bounceprocessing", "navigationdelay", "nokeyboard",
            "alloweditaftercompletion", "showxquestions",
            "shownoanswer", "showwelcome", "questionindex",
            # language/additional_languages auto-generated by forward script
            "language", "additional_languages")

s_rows <- tsv[tsv$class == "S", ]
s_kept <- s_rows[!s_rows$name %in% drop_s, ]
s_dropped <- s_rows[s_rows$name %in% drop_s, ]
cat(sprintf("  S rows: %d kept, %d dropped\n", nrow(s_kept), nrow(s_dropped)))

# ==============================================================================
# STEP 4: Filter SL rows -- keep essential 4 fields only
# ==============================================================================
essential_sl <- c("surveyls_title", "surveyls_description",
                  "surveyls_welcometext", "surveyls_endtext")

sl_rows <- tsv[tsv$class == "SL", ]
sl_kept <- sl_rows[sl_rows$name %in% essential_sl, ]
cat(sprintf("  SL rows: %d kept (from %d)\n",
            length(unique(sl_kept$name)), length(unique(sl_rows$name))))

# ==============================================================================
# STEP 5: Extract base-language G/Q/SQ/A rows with attributes
# ==============================================================================
content_rows <- tsv[tsv$class %in% c("G", "Q", "SQ", "A") & tsv$language == base_lang, ]
cat(sprintf("  Content rows (base language): %d\n", nrow(content_rows)))

# ==============================================================================
# STEP 6: Determine which advanced attribute columns to keep
# ==============================================================================
# Standard columns (first 14 in the TSV spec)
standard_cols <- c("id", "related_id", "class", "type/scale", "name",
                   "relevance", "text", "help", "language",
                   "validation", "mandatory", "encrypted", "other",
                   "default", "same_default")

# Columns to always DROP (server-specific, theme-specific, or causes import issues)
always_drop <- c("id", "related_id", "encrypted",
                 "question_theme_name", "same_script", "parent_order")

# Find advanced columns that have at least one non-empty value in base-language rows
all_tsv_cols <- names(tsv)
potential_adv <- setdiff(all_tsv_cols, c(standard_cols, always_drop))

adv_keep <- character(0)
for (col in potential_adv) {
  vals <- content_rows[[col]]
  # Keep if any non-empty, non-zero, non-default value exists
  non_empty <- vals[!is.na(vals) & vals != "" & vals != "0"]
  if (length(non_empty) > 0) {
    adv_keep <- c(adv_keep, col)
  }
}
adv_keep <- sort(adv_keep)

if (length(adv_keep) > 0) {
  cat(sprintf("  Advanced attributes kept: %s\n", paste(adv_keep, collapse = ", ")))
} else {
  cat("  No advanced attributes with non-default values\n")
}

# ==============================================================================
# STEP 7: Build the collapsed output dataframe
# ==============================================================================
cat("\nBuilding Excel structure...\n")

# Output columns: standard (without id/related_id/encrypted/language/text/help)
# + text_xx/help_xx per language + advanced attributes
out_std <- c("class", "type/scale", "name", "relevance",
             "validation", "mandatory", "other", "default", "same_default")

# Language columns in order
lang_cols <- character(0)
for (lang in lang_order) {
  lang_cols <- c(lang_cols, paste0("text_", lang), paste0("help_", lang))
}

out_cols <- c(out_std[1:4],  # class, type/scale, name, relevance
              lang_cols,      # text_en, help_en, text_fr, help_fr, ...
              out_std[5:9],  # validation, mandatory, other, default, same_default
              adv_keep)

# Initialize output list
out_list <- list()

# --- S rows (one per setting) ---
for (i in seq_len(nrow(s_kept))) {
  row <- setNames(rep("", length(out_cols)), out_cols)
  row["class"] <- "S"
  row["name"] <- s_kept$name[i]
  # S row text goes into text_<base_lang> only
  base_text_col <- paste0("text_", base_lang)
  if (base_text_col %in% out_cols) {
    row[base_text_col] <- if (is.na(s_kept$text[i])) "" else s_kept$text[i]
  }
  out_list[[length(out_list) + 1]] <- row
}

# --- SL rows (one per field, text spread across language columns) ---
for (field in essential_sl) {
  row <- setNames(rep("", length(out_cols)), out_cols)
  row["class"] <- "SL"
  row["name"] <- field
  # Fill each language column
  for (lang in lang_order) {
    sl_match <- sl_kept[sl_kept$name == field & sl_kept$language == lang, ]
    if (nrow(sl_match) > 0) {
      tc <- paste0("text_", lang)
      hc <- paste0("help_", lang)
      if (tc %in% out_cols) row[tc] <- if (is.na(sl_match$text[1])) "" else sl_match$text[1]
      if (hc %in% out_cols) row[hc] <- if (is.na(sl_match$help[1])) "" else sl_match$help[1]
    }
  }
  out_list[[length(out_list) + 1]] <- row
}

# --- G/Q/SQ/A rows (structure from base language, translations from others) ---
# Index translated rows by language (positional matching)
trans_rows <- list()
for (lang in other_langs) {
  lang_content <- tsv[tsv$class %in% c("G", "Q", "SQ", "A") & tsv$language == lang, ]
  trans_rows[[lang]] <- lang_content
}

# Verify same row count across languages
for (lang in other_langs) {
  n_trans <- nrow(trans_rows[[lang]])
  if (n_trans != nrow(content_rows)) {
    cat(sprintf("  WARNING: %s has %d rows vs base %d rows\n",
                lang, n_trans, nrow(content_rows)))
  }
}

for (i in seq_len(nrow(content_rows))) {
  row <- setNames(rep("", length(out_cols)), out_cols)

  # Standard fields from base language
  for (col in out_std) {
    if (col %in% names(content_rows)) {
      val <- content_rows[[col]][i]
      row[col] <- if (is.na(val)) "" else val
    }
  }

  # Base language text/help
  base_tc <- paste0("text_", base_lang)
  base_hc <- paste0("help_", base_lang)
  if (base_tc %in% out_cols) {
    row[base_tc] <- if (is.na(content_rows$text[i])) "" else content_rows$text[i]
  }
  if (base_hc %in% out_cols) {
    row[base_hc] <- if (is.na(content_rows$help[i])) "" else content_rows$help[i]
  }

  # Translated text/help from other languages (positional match)
  for (lang in other_langs) {
    lang_df <- trans_rows[[lang]]
    if (i <= nrow(lang_df)) {
      tc <- paste0("text_", lang)
      hc <- paste0("help_", lang)
      if (tc %in% out_cols) {
        row[tc] <- if (is.na(lang_df$text[i])) "" else lang_df$text[i]
      }
      if (hc %in% out_cols) {
        row[hc] <- if (is.na(lang_df$help[i])) "" else lang_df$help[i]
      }
    }
  }

  # Advanced attributes
  for (col in adv_keep) {
    if (col %in% names(content_rows)) {
      val <- content_rows[[col]][i]
      row[col] <- if (is.na(val)) "" else val
    }
  }

  out_list[[length(out_list) + 1]] <- row
}

# Convert to dataframe
df_out <- as.data.frame(do.call(rbind, out_list), stringsAsFactors = FALSE)
names(df_out) <- out_cols

cat(sprintf("  Output rows: %d (S=%d, SL=%d, G/Q/SQ/A=%d)\n",
            nrow(df_out), nrow(s_kept),
            length(essential_sl), nrow(content_rows)))

# Count by class
for (cls in c("S", "SL", "G", "Q", "SQ", "A")) {
  n <- sum(df_out$class == cls)
  if (n > 0) cat(sprintf("    %s: %d\n", cls, n))
}

# ==============================================================================
# STEP 7.5: HTML to Excel rich text conversion
# ==============================================================================
cat("\nConverting HTML to Excel rich text...\n")

#' Convert an HTML string (from LimeSurvey) to an openxlsx2 fmt_txt object.
#' Handles: <strong>/<b>, <em>/<i>, <u>, <span style='color:#HEX'>,
#'          <p>...</p> (paragraphs), <br>/<br/> (line breaks).
#' Returns a plain string if no formatting tags are present,
#' or an fmt_txt object chain if formatting is found.
html_to_rich_text <- function(html_str) {
  if (is.na(html_str) || html_str == "") return("")

  # If no HTML tags at all, return plain text
  if (!grepl("<[^>]+>", html_str)) return(html_str)

  s <- html_str

  # --- Step A: Normalize block elements to newlines ---
  s <- gsub("<br\\s*/?>", "\n", s, ignore.case = TRUE)
  s <- gsub("</p>\\s*<p[^>]*>", "\n", s, ignore.case = TRUE)
  s <- gsub("^\\s*<p[^>]*>", "", s, ignore.case = TRUE)
  s <- gsub("</p>\\s*$", "", s, ignore.case = TRUE)
  s <- gsub("</?p[^>]*>", "", s, ignore.case = TRUE)

  # --- Step B: If no inline formatting tags remain, return plain text ---
  if (!grepl("<(strong|em|b|i|u|span)[^>]*>", s, ignore.case = TRUE)) {
    return(trimws(s))
  }

  # --- Step C: Tokenize into text segments and tags ---
  tag_pattern <- "<[^>]+>"
  tag_locs <- gregexpr(tag_pattern, s)[[1]]
  if (tag_locs[1] == -1) return(s)

  tag_lengths <- attr(tag_locs, "match.length")

  tokens <- list()
  pos <- 1
  for (j in seq_along(tag_locs)) {
    # Text before this tag
    if (tag_locs[j] > pos) {
      txt <- substr(s, pos, tag_locs[j] - 1)
      if (nchar(txt) > 0) tokens[[length(tokens) + 1]] <- list(type = "text", value = txt)
    }
    # The tag itself
    tag_str <- substr(s, tag_locs[j], tag_locs[j] + tag_lengths[j] - 1)
    tokens[[length(tokens) + 1]] <- list(type = "tag", value = tag_str)
    pos <- tag_locs[j] + tag_lengths[j]
  }
  # Text after last tag
  if (pos <= nchar(s)) {
    txt <- substr(s, pos, nchar(s))
    if (nchar(txt) > 0) tokens[[length(tokens) + 1]] <- list(type = "text", value = txt)
  }

  # --- Step D: Walk tokens, tracking formatting state, building fmt_txt segments ---
  segments <- list()
  bold <- FALSE
  italic <- FALSE
  underline <- FALSE
  color_stack <- list()

  for (tok in tokens) {
    if (tok$type == "text") {
      txt_val <- tok$value
      if (nchar(txt_val) == 0) next

      cur_color <- NULL
      if (length(color_stack) > 0) {
        # Find the last non-NULL color in the stack
        for (ci in rev(seq_along(color_stack))) {
          if (!is.null(color_stack[[ci]])) {
            cur_color <- color_stack[[ci]]
            break
          }
        }
      }

      if (!bold && !italic && !underline && is.null(cur_color)) {
        seg <- fmt_txt(txt_val)
      } else {
        args <- list(txt_val)
        if (bold) args$bold <- "true"
        if (italic) args$italic <- "true"
        if (underline) args$underline <- "single"
        if (!is.null(cur_color)) args$color <- wb_color(cur_color)
        seg <- do.call(fmt_txt, args)
      }
      segments[[length(segments) + 1]] <- seg

    } else {
      tag <- tok$value
      tag_lower <- tolower(tag)

      if (grepl("^<(strong|b)(\\s|>)", tag_lower))        bold <- TRUE
      else if (grepl("^</(strong|b)>", tag_lower))         bold <- FALSE
      else if (grepl("^<(em|i)(\\s|>)", tag_lower))        italic <- TRUE
      else if (grepl("^</(em|i)>", tag_lower))             italic <- FALSE
      else if (grepl("^<u(\\s|>)", tag_lower))             underline <- TRUE
      else if (grepl("^</u>", tag_lower))                  underline <- FALSE
      else if (grepl("^<span", tag_lower)) {
        m <- regmatches(tag, regexpr("#[0-9A-Fa-f]{6}", tag))
        if (length(m) > 0 && nchar(m) == 7) {
          hex <- toupper(sub("#", "", m))
          if (hex != "000000") {
            color_stack[[length(color_stack) + 1]] <- paste0("FF", hex)
          } else {
            color_stack[[length(color_stack) + 1]] <- NULL
          }
        } else {
          color_stack[[length(color_stack) + 1]] <- NULL
        }
      }
      else if (grepl("^</span>", tag_lower)) {
        if (length(color_stack) > 0) {
          color_stack[[length(color_stack)]] <- NULL
        }
      }
    }
  }

  # --- Step E: Combine segments ---
  if (length(segments) == 0) return(trimws(s))
  if (length(segments) == 1) return(segments[[1]])

  result <- segments[[1]]
  for (k in 2:length(segments)) {
    result <- result + segments[[k]]
  }
  return(result)
}

# --- Apply HTML-to-rich-text conversion to all text/help columns ---
rich_text_count <- 0
# Store rich text objects: key = "row,col" -> fmt_txt object
rich_text_cells <- list()

text_help_col_indices <- which(grepl("^(text|help)_", out_cols))

for (col_idx in text_help_col_indices) {
  col_name <- out_cols[col_idx]
  for (row_idx in seq_len(nrow(df_out))) {
    val <- df_out[[col_name]][row_idx]
    if (is.na(val) || val == "") next
    # Skip S rows (settings are plain text, never HTML)
    if (df_out$class[row_idx] == "S") next

    if (grepl("<[^>]+>", val)) {
      rt <- tryCatch(html_to_rich_text(val), error = function(e) NULL)
      if (is.null(rt)) next
      if (inherits(rt, "fmt_txt")) {
        key <- paste(row_idx, col_idx, sep = ",")
        rich_text_cells[[key]] <- rt
        rich_text_count <- rich_text_count + 1
      } else if (is.character(rt)) {
        # Block-level HTML was stripped, update plain text
        df_out[[col_name]][row_idx] <- rt
      }
    }
  }
}
cat(sprintf("  Rich text cells converted: %d\n", rich_text_count))

# ==============================================================================
# STEP 8: Create Excel workbook with formatting (openxlsx2)
# ==============================================================================
cat("\nCreating Excel workbook...\n")

wb <- wb_workbook()

ws_name <- "Survey Design"
wb$add_worksheet(sheet = ws_name)

# Write plain text data first
wb$add_data(sheet = ws_name, x = df_out, start_row = 1, start_col = 1)

ncol_out <- ncol(df_out)
nrow_out <- nrow(df_out)

# --- Write rich text cells (overwrite HTML strings with formatted objects) ---
if (length(rich_text_cells) > 0) {
  cat(sprintf("  Writing %d rich text cells...\n", length(rich_text_cells)))
  for (key in names(rich_text_cells)) {
    parts <- as.integer(strsplit(key, ",")[[1]])
    r <- parts[1]
    c_idx <- parts[2]
    # Data starts at Excel row 2 (row 1 is header)
    cell_dims <- wb_dims(rows = r + 1, cols = c_idx)
    wb$add_data(sheet = ws_name, x = rich_text_cells[[key]], dims = cell_dims)
  }
}

# --- Header styling ---
header_dims <- wb_dims(rows = 1, cols = 1:ncol_out)
wb$add_font(sheet = ws_name, dims = header_dims,
            name = "Arial", size = "10",
            color = wb_color("FFFFFFFF"), bold = "true")
wb$add_fill(sheet = ws_name, dims = header_dims,
            color = wb_color("FF2C3E50"))
wb$add_cell_style(sheet = ws_name, dims = header_dims,
                  wrap_text = "true", horizontal = "left", vertical = "center")

# --- Data font (Arial 10 for all data cells) ---
if (nrow_out > 0) {
  data_dims <- wb_dims(rows = 2:(nrow_out + 1), cols = 1:ncol_out)
  wb$add_font(sheet = ws_name, dims = data_dims,
              name = "Arial", size = "10")
  wb$add_cell_style(sheet = ws_name, dims = data_dims, vertical = "center")
}

# --- Conditional formatting (colors matching master template) ---
max_row <- nrow_out + 1000
class_col_idx <- which(out_cols == "class")
class_col_letter <- int2col(class_col_idx)
cf_dims <- wb_dims(rows = 2:max_row, cols = 1:ncol_out)

# Colors from master template (project document reference)
cf_colors <- list(
  S  = "FFD5D0C4",   # light tan/beige
  SL = "FFCDB5D6",   # light lavender
  G  = "FF93CDDD",   # light blue
  Q  = "FFFFFFFF",   # white
  SQ = "FFB9F1D1",   # light green
  A  = "FFFFF2AE"    # light yellow
)

# Register each color as a named dxf style (ensures unique dxfId per class)
for (cls in names(cf_colors)) {
  wb$add_dxfs_style(
    name    = paste0("cf_", cls),
    bg_fill = wb_color(cf_colors[[cls]])
  )
}

for (cls in names(cf_colors)) {
  wb$add_conditional_formatting(
    sheet = ws_name,
    dims = cf_dims,
    rule = paste0("$", class_col_letter, "2=\"", cls, "\""),
    style = paste0("cf_", cls),
    type = "expression"
  )
}

# --- Data validations ---
class_val_dims <- wb_dims(rows = 2:100000, cols = class_col_idx)
wb$add_data_validation(
  sheet = ws_name,
  dims = class_val_dims,
  type = "list",
  value = '"Q,SQ,A,G,S,SL"'
)

mand_col <- which(out_cols == "mandatory")
if (length(mand_col) > 0) {
  wb$add_data_validation(
    sheet = ws_name,
    dims = wb_dims(rows = 2:100000, cols = mand_col),
    type = "list",
    value = '"Y,N,"'
  )
}

other_col <- which(out_cols == "other")
if (length(other_col) > 0) {
  wb$add_data_validation(
    sheet = ws_name,
    dims = wb_dims(rows = 2:100000, cols = other_col),
    type = "list",
    value = '"Y,N,"'
  )
}

# --- Column widths ---
for (i in seq_along(out_cols)) {
  col_name <- out_cols[i]
  w <- if (col_name == "class") 7
       else if (col_name == "type/scale") 10
       else if (col_name == "name") 18
       else if (col_name == "relevance") 30
       else if (grepl("^text_", col_name)) 40
       else if (grepl("^help_", col_name)) 30
       else if (col_name %in% c("validation", "mandatory", "other")) 10
       else if (col_name %in% c("default", "same_default")) 12
       else 14
  wb$set_col_widths(sheet = ws_name, cols = i, widths = w)
}

# Freeze top row
wb$freeze_pane(sheet = ws_name, first_row = TRUE)

# ==============================================================================
# Reference Sheets
# ==============================================================================

# --- Question Types Reference ---
wb$add_worksheet(sheet = "Question Types Reference")
qt_data <- data.frame(
  Code = c("L","!","O","M","P","F","H","1","A","B","C","E",
           "N","K","S","T","U","D","R","|","*","X","Y",";",":"),
  Type = c("List (Radio)","List (Dropdown)","List with comment",
           "Multiple choice","Multiple choice with comments",
           "Array","Array (Column)","Array dual scale",
           "Array (5 point)","Array (10 point)","Array (Yes/No/Uncertain)",
           "Array (Increase/Same/Decrease)",
           "Numerical","Multiple numerical","Short text","Long text",
           "Huge text","Date/Time","Ranking","File upload",
           "Equation","Boilerplate","Yes/No",
           "Array (Texts)","Array (Numbers)"),
  SubQ = c("No","No","No","Yes","Yes","Yes","Yes","Yes",
           "Yes","Yes","Yes","Yes","No","Yes","No","No",
           "No","No","No","No","No","No","No","Yes","Yes"),
  Answers = c("Yes","Yes","Yes","No","No","Yes","Yes","Yes",
              "No","No","No","No","No","No","No","No",
              "No","No","Yes","No","No","No","No","Yes","Yes"),
  Notes = c("Single select radio buttons","Single select dropdown",
            "List + comment text box","Checkboxes","Checkboxes + comment per item",
            "Rows=SQ, Cols=A (Likert)","Like F but columns are SQ",
            "Two scales per row","Fixed 1-5 scale","Fixed 1-10 scale",
            "Fixed Y/N/? scale","Fixed Inc/Same/Dec scale",
            "Single number input","Multiple number inputs (one per SQ)",
            "Single line text","Multi-line text","Very large text area",
            "Date/time picker","Drag-and-drop ranking (uses A rows)",
            "File upload control","Calculated field (ExpressionScript)",
            "Display text only (no input)","Simple Yes/No radio",
            "Text input grid (SQ=rows, A=cols)",
            "Number input grid (SQ=rows, A=cols)"),
  stringsAsFactors = FALSE
)
wb$add_data(sheet = "Question Types Reference", x = qt_data)
qt_header_dims <- wb_dims(rows = 1, cols = 1:5)
wb$add_font(sheet = "Question Types Reference", dims = qt_header_dims,
            name = "Arial", size = "10", color = wb_color("FFFFFFFF"), bold = "true")
wb$add_fill(sheet = "Question Types Reference", dims = qt_header_dims,
            color = wb_color("FF34495E"))
wb$set_col_widths(sheet = "Question Types Reference",
                  cols = 1:5, widths = c(8, 28, 8, 10, 55))

# --- Survey Settings Reference ---
wb$add_worksheet(sheet = "Survey Settings Reference")
ss_data <- data.frame(
  Setting = c("sid","admin","adminemail","anonymized","format",
              "datestamp","ipaddr","showprogress","allowprev",
              "printanswers","showgroupinfo","showqnumcode","allowsave"),
  Default = c("(auto)","Survey Admin","admin@example.com","N","G",
              "Y","N","Y","Y","Y","D","X","Y"),
  Description = c("Survey ID (auto-assigned on import)",
                   "Administrator name","Administrator email",
                   "Anonymize responses (Y/N)",
                   "Format: G=group-by-group, S=question-by-question, A=all-in-one",
                   "Save date stamp with responses","Save IP address",
                   "Show progress bar","Allow backward navigation",
                   "Allow print answers at end",
                   "Show group info: B=both, D=description, N=none",
                   "Show question number/code: B=both, N=number, C=code, X=none",
                   "Allow participants to save and resume later"),
  stringsAsFactors = FALSE
)
wb$add_data(sheet = "Survey Settings Reference", x = ss_data)
ss_header_dims <- wb_dims(rows = 1, cols = 1:3)
wb$add_font(sheet = "Survey Settings Reference", dims = ss_header_dims,
            name = "Arial", size = "10", color = wb_color("FFFFFFFF"), bold = "true")
wb$add_fill(sheet = "Survey Settings Reference", dims = ss_header_dims,
            color = wb_color("FF34495E"))
wb$set_col_widths(sheet = "Survey Settings Reference",
                  cols = 1:3, widths = c(22, 15, 60))

# --- Instructions ---
wb$add_worksheet(sheet = "Instructions")
instr <- data.frame(
  Topic = c("Workflow",
            "", "",
            "Languages",
            "", "",
            "Row Classes",
            "", "", "", "", "",
            "Key Rules",
            "", "", "", ""),
  Detail = c("This file was generated from a LimeSurvey TSV export.",
             "Edit the Survey Design sheet, then run xlsx_to_limesurvey_tsv.R to convert back to .txt.",
             "Import the .txt file in LimeSurvey: Create Survey > Import.",
             "Each language has text_xx and help_xx columns side by side.",
             "The forward script auto-detects languages from these columns.",
             "Untranslated cells fall back to the base language automatically.",
             "S = Survey setting (text goes in base language column only)",
             "SL = Survey language text (title, description, welcome, end)",
             "G = Question group (section header)",
             "Q = Question (with type code, name, relevance, mandatory)",
             "SQ = Subquestion (row/column in arrays, items in multiple choice)",
             "A = Answer option (choices in lists, scales; ranking uses A not SQ)",
             "Row order matters: S > SL > G > Q > SQ/A > Q > SQ/A > G > ...",
             "Question codes (name column): alphanumeric only, no underscores.",
             "S rows: leave text_xx empty for non-base languages.",
             "SL rows: translate all 4 fields per language.",
             "G/Q/SQ/A: translate text_xx; help_xx is optional."),
  stringsAsFactors = FALSE
)
wb$add_data(sheet = "Instructions", x = instr)
instr_header_dims <- wb_dims(rows = 1, cols = 1:2)
wb$add_font(sheet = "Instructions", dims = instr_header_dims,
            name = "Arial", size = "10", color = wb_color("FFFFFFFF"), bold = "true")
wb$add_fill(sheet = "Instructions", dims = instr_header_dims,
            color = wb_color("FF34495E"))
wb$set_col_widths(sheet = "Instructions", cols = 1:2, widths = c(15, 80))

# ==============================================================================
# STEP 9: Validation summary
# ==============================================================================
cat("\n--- Validation ---\n")
q_names <- df_out$name[df_out$class == "Q" & df_out$name != ""]
dupes <- q_names[duplicated(q_names)]
if (length(dupes) > 0) {
  cat(sprintf("  WARNING: Duplicate question codes: %s\n",
              paste(unique(dupes), collapse = ", ")))
}
q_underscore <- q_names[grepl("_", q_names, fixed = TRUE)]
if (length(q_underscore) > 0) {
  cat(sprintf("  WARNING: Question codes with underscores: %s\n",
              paste(unique(q_underscore), collapse = ", ")))
}

cat(sprintf("  Languages: %s\n", paste(lang_order, collapse = ", ")))
cat(sprintf("  Columns: %d (%d standard + %d language + %d advanced)\n",
            length(out_cols),
            length(out_std),
            length(lang_cols),
            length(adv_keep)))

# ==============================================================================
# STEP 10: Save workbook
# ==============================================================================
cat("\n--- Writing output ---\n")
wb$save(output_file, overwrite = TRUE)

cat(sprintf("Output:   %s\n", output_file))
cat(sprintf("Rows:     %d data + 1 header\n", nrow(df_out)))
cat(sprintf("Columns:  %d\n", ncol(df_out)))
cat(sprintf("Rich text: %d cells with HTML converted to Excel formatting\n", rich_text_count))
cat(sprintf("Languages: %s (base: %s)\n",
            paste(lang_order, collapse = ", "), base_lang))
cat(sprintf("\nNext step: edit in Excel, then run xlsx_to_limesurvey_tsv.R\n"))
cat("Done!\n")
