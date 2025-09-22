# =========================
# 03_compatibilita_trasfusioni.R
# Outputs per macro (RBC / Plasma / Platelets):
#  - PNG/CSV table of counts: Donors (rows) × Recipients (columns) + Totals
#  - PNG/CSV row-wise percentages (rows sum to 100%), without Totals
# Folders:
#   <out_root>/<macro>/{png, csv}
# =========================

suppressPackageStartupMessages({
  pkgs <- c("tidyverse","janitor","lubridate","stringr","glue","fs","scales","gt","readr")
  for (p in pkgs) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(tidyverse); library(janitor); library(lubridate); library(stringr)
  library(glue); library(fs); library(scales); library(gt); library(readr)
})

# -----------------
# PATHS
# -----------------
dir_trasfuso <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized"
file_ulss7   <- file.path(dir_trasfuso, "ULSS7.csv")
file_ulss8   <- file.path(dir_trasfuso, "ULSS8.csv")

out_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi complessive/Compatibilita"

# -----------------
# WORD/A4 SIZING (PNG allineate alla larghezza testo; tabelle più strette)
# -----------------
page_width_in   <- 8.27    # A4 portrait
left_margin_pt  <- 3       # 3 pt per lato
right_margin_pt <- 3
target_ppi      <- 96      # Word posiziona alla misura esatta

content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72
content_width_px <- round(content_width_in * target_ppi)

# Quanta parte della larghezza testo occupa la tabella (equilibrata)
table_pct <- 80  # consigliato 76–84

# -----------------
# HELPERS
# -----------------
log_msg <- function(...) cat(format(Sys.time(), "[%Y-%m-%d %H:%M:%S]"), ..., "\n")

safe_read <- function(path, guess_max = 200000) {
  stopifnot(!is.na(path), file.exists(path))
  readr::read_csv(path, guess_max = guess_max, show_col_types = FALSE) %>% clean_names()
}

normalize_group <- function(x) {
  x <- as.character(x)
  x[x %in% c("", "NA", "na")] <- NA_character_
  x <- toupper(trimws(x))
  x <- str_replace_all(x, "POSITIVO", "+")
  x <- str_replace_all(x, "NEGATIVO", "-")
  x <- str_replace_all(x, "\\s+", "")
  case_when(
    x %in% c("O-","O+","A-","A+","B-","B+","AB-","AB+") ~ x,
    TRUE ~ NA_character_
  )
}

# Component recognizers
is_rbc      <- function(x){ x0<-tolower(x); str_detect(x0,"emaz|eritro|\\brbc\\b|\\b250(4)?\\b|leucodeplet") & !str_detect(x0,"plasm|\\bplt\\b|piastr|buffy") }
is_plasma   <- function(x){ x0<-tolower(x); str_detect(x0,"plasm") & !str_detect(x0,"\\bplt\\b|piastr|buffy|4305") }
is_platelet <- function(x){ x0<-tolower(x); str_detect(x0,"\\b4305\\b|\\bplt\\b|piastr|platelet|buffy") }

macro_from_component <- function(x) {
  case_when(
    is_rbc(x)      ~ "RBC",
    is_plasma(x)   ~ "Plasma",
    is_platelet(x) ~ "Platelets",
    TRUE ~ NA_character_
  )
}

# Trova colonne gruppo (dopo clean_names)
resolve_group_cols <- function(df) {
  nms <- names(df)
  donor_cand <- nms[ str_detect(nms, "gruppo") & str_detect(nms, "rh") & !str_detect(nms, "rich") ]
  recip_cand <- nms[ str_detect(nms, "gruppo") & str_detect(nms, "rh") &  str_detect(nms, "rich") ]
  list(donor = donor_cand[1], recip = recip_cand[1])
}

# Trova colonna data
resolve_date_col <- function(df) {
  nms <- names(df)
  cand <- nms[str_detect(nms, "data|date")]
  cand[1]
}

# Sceglie una colonna componente fra candidati
get_component_vector <- function(df, candidates) {
  for (nm in candidates) if (nm %in% names(df)) return(as.character(df[[nm]]))
  rep("", nrow(df))
}

# -----------------
# STILE TABELL E EXPORT (gt)
# -----------------
gt_style_thesis <- function(gt_tbl,
                            title_px = 13, base_px = 11, label_px = 11,
                            font_family = c("Times New Roman", "Liberation Serif", "serif"),
                            table_pct_local = table_pct) {
  gt_tbl %>%
    tab_options(
      table.align = "center",
      table.width = pct(table_pct_local),
      table.font.names = font_family,
      table.font.size  = px(base_px),
      heading.title.font.size = px(title_px),
      data_row.padding = px(6),
      column_labels.border.top.width    = px(1),
      column_labels.border.bottom.width = px(1),
      column_labels.vlines.width        = px(0),
      table.border.top.width            = px(0),
      table.border.bottom.width         = px(0),
      column_labels.background.color    = "white",
      heading.background.color          = "white"
    ) %>%
    tab_style(
      style = cell_text(weight = "bold", size = px(label_px)),
      locations = cells_column_labels(everything())
    ) %>%
    cols_align(align = "left",  columns = where(~ !is.numeric(.))) %>%
    cols_align(align = "right", columns = where(is.numeric)) %>%
    opt_row_striping()
}

save_gt_png <- function(gt_tbl, path) {
  fs::dir_create(dirname(path), recurse = TRUE)
  gt::gtsave(
    gt_tbl,
    filename = path,
    vwidth   = content_width_px,  # canvas = larghezza del testo
    expand   = 0                  # nessuno spazio bianco esterno
  )
}

# Costruisce tabella COUNT (Donors × Recipients + Totals) in gt
build_counts_gt <- function(counts_df, groups_order, title_txt) {
  counts_df %>%
    gt(rowname_col = "donor") %>%
    tab_header(title = title_txt) %>%
    tab_stubhead(label = "Donors") %>%
    tab_spanner(label = "Recipients", columns = all_of(groups_order)) %>%
    cols_label(.list = setNames(groups_order, groups_order)) %>%
    fmt_number(columns = all_of(groups_order), decimals = 0, use_seps = TRUE) %>%
    fmt_number(columns = "Total", decimals = 0, use_seps = TRUE) %>%
    tab_style(
      style = cell_text(weight = "bold"),
      locations = list(
        cells_stub(rows = donor == "Total"),
        cells_body(rows = donor == "Total", columns = everything())
      )
    ) %>%
    gt_style_thesis()
}

# Costruisce tabella ROW % (senza Totals) in gt
build_pct_gt <- function(pct_df_chr, groups_order, title_txt) {
  # per PNG usiamo numerico + formattazione %, ma lasciamo il CSV in stringa com'è
  pct_num <- pct_df_chr %>%
    mutate(across(all_of(groups_order), ~ suppressWarnings(readr::parse_number(.x)))) 
  pct_num %>%
    gt(rowname_col = "donor") %>%
    tab_header(title = title_txt) %>%
    tab_stubhead(label = "Donors") %>%
    tab_spanner(label = "Recipients", columns = all_of(groups_order)) %>%
    cols_label(.list = setNames(groups_order, groups_order)) %>%
    fmt_number(columns = all_of(groups_order), decimals = 1, pattern = "{x}%") %>%
    gt_style_thesis()
}

# -----------------
# BUILD COUNT / PCT TABLES
# -----------------
make_counts_table <- function(df, groups_order) {
  xt <- df %>%
    count(donor, recipient, name = "n") %>%
    tidyr::pivot_wider(names_from = recipient, values_from = n, values_fill = 0)
  
  for (g in groups_order) if (!g %in% names(xt)) xt[[g]] <- 0L
  
  xt <- xt %>%
    select(donor, all_of(groups_order)) %>%
    arrange(factor(donor, levels = groups_order, ordered = TRUE))
  
  col_tot <- c("Total", colSums(select(xt, -donor), na.rm = TRUE))
  col_tot <- as_tibble_row(setNames(as.list(col_tot), c("donor", groups_order)))
  
  xt$Total <- rowSums(select(xt, -donor), na.rm = TRUE)
  bind_rows(xt, col_tot)
}

make_pct_table <- function(counts_no_tot, groups_order, digits = 1) {
  m <- counts_no_tot %>% select(donor, all_of(groups_order))
  num <- as.matrix(m[, groups_order, drop = FALSE]); storage.mode(num) <- "double"
  row_sum <- rowSums(num, na.rm = TRUE); row_sum[row_sum == 0] <- NA_real_
  pct <- sweep(num, 1, row_sum, `/`) * 100
  pct_df <- as_tibble(pct) %>% mutate(donor = m$donor, .before = 1)
  pct_df %>% mutate(across(all_of(groups_order), ~ ifelse(is.na(.x), "", sprintf(paste0("%.",digits,"f%%"), .x))))
}

# -----------------
# LOAD & PREP
# -----------------
stopifnot(file.exists(file_ulss7), file.exists(file_ulss8))
log_msg("Reading transfusion data: ULSS7 + ULSS8")
ulss7 <- safe_read(file_ulss7) %>% mutate(source = "ULSS7")
ulss8 <- safe_read(file_ulss8) %>% mutate(source = "ULSS8")
trf_raw <- bind_rows(ulss7, ulss8)

cols <- resolve_group_cols(trf_raw)
if (is.null(cols$donor) || is.null(cols$recip)) stop("Cannot identify donor/recipient group columns after clean_names().")
date_col <- resolve_date_col(trf_raw)
if (is.na(date_col) || is.null(date_col)) stop("Cannot identify a date column (expecting 'data' or 'date').")

# Time window (inclusive): 2023-01-01 .. 2025-05-31
date_min <- as.Date("2023-01-01")
date_max <- as.Date("2025-05-31")

component_vec <- get_component_vector(trf_raw, c("emc_emolife", "emocomponente"))

trf <- trf_raw %>%
  mutate(
    trx_date = as.Date(suppressWarnings(lubridate::as_date(.data[[date_col]]))),
    donor_group_raw = .data[[cols$donor]],
    recip_group_raw = .data[[cols$recip]],
    donor = normalize_group(donor_group_raw),
    recipient = normalize_group(recip_group_raw),
    macro = macro_from_component(coalesce(.env$component_vec, ""))
  ) %>%
  filter(!is.na(macro), !is.na(donor), !is.na(recipient),
         !is.na(trx_date), trx_date >= date_min, trx_date <= date_max)

# -----------------
# OUTPUT BY MACRO
# -----------------
macro_levels <- c("RBC","Plasma","Platelets")
groups_order <- c("O-","O+","A-","A+","B-","B+","AB-","AB+")

for (mc in macro_levels) {
  log_msg("== Macro:", mc, "==")
  out_macro_dir <- file.path(out_root, mc)
  dir_png <- file.path(out_macro_dir, "png")
  dir_csv <- file.path(out_macro_dir, "csv")
  lapply(c(dir_png, dir_csv), fs::dir_create, recurse = TRUE)
  
  df <- trf %>% filter(macro == mc)
  if (nrow(df) == 0) {
    log_msg("[", mc, "] No compatibility rows in the selected period: skipping.")
    next
  }
  
  # ---- COUNTS (with Totals) ----
  counts_core <- df %>%
    count(donor, recipient, name = "n") %>%
    tidyr::pivot_wider(names_from = recipient, values_from = n, values_fill = 0) %>%
    { . -> xt; for (g in groups_order) if (!g %in% names(xt)) xt[[g]] <- 0L; xt } %>%
    select(donor, all_of(groups_order)) %>%
    arrange(factor(donor, levels = groups_order, ordered = TRUE))
  
  counts_with_tot <- counts_core %>% mutate(Total = rowSums(across(all_of(groups_order)), na.rm = TRUE))
  
  col_tot <- tibble(donor = "Total")
  for (g in groups_order) col_tot[[g]] <- sum(counts_core[[g]], na.rm = TRUE)
  col_tot$Total <- sum(counts_with_tot$Total, na.rm = TRUE)
  
  counts_full <- bind_rows(counts_with_tot, col_tot)
  
  # CSV (counts)
  csv_counts <- file.path(dir_csv, glue("{mc}_compat_counts.csv"))
  readr::write_csv(counts_full, csv_counts)
  
  # PNG (counts) — stile thesis come nello script precedente
  counts_gt <- build_counts_gt(counts_full, groups_order, title_txt = glue("{mc} — Compatibility analysis"))
  png_counts <- file.path(dir_png, glue("{mc}_compat_counts.png"))
  save_gt_png(counts_gt, png_counts)
  
  # ---- ROW PERCENTAGES (no Totals) ----
  pct_tbl <- make_pct_table(counts_core, groups_order, digits = 1)
  
  # CSV (row %)
  csv_pct <- file.path(dir_csv, glue("{mc}_compat_rowpct.csv"))
  readr::write_csv(pct_tbl, csv_pct)
  
  # PNG (row %) — stile thesis, numerico + formattazione %
  pct_gt <- build_pct_gt(pct_tbl, groups_order, title_txt = glue("{mc} — Compatibility analysis (row percentages)"))
  png_pct <- file.path(dir_png, glue("{mc}_compat_rowpct.png"))
  save_gt_png(pct_gt, png_pct)
  
  log_msg("[", mc, "] saved counts + row percentages.")
}

log_msg("DONE. Outputs in: ", out_root)
