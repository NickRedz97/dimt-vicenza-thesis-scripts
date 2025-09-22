# ────────────────────────────────────────────────────────────────────────────────
# Unified TO-BE script — RBCs / Plasma / Platelets — ULSS7 + ULSS8
# Inputs:
#   C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized/{ULSS7.csv, ULSS8.csv}
# Per-output (by component, year, facility):
#   - model_{facility}_{year}.csv
#   - summary_{facility}_{year}.png  (thesis-style table)
# Output folders:
#   C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi_inventory_management/{Emazie|Plasma|Piastrine}/TO-BE/<Year>/<Facility>/
# ────────────────────────────────────────────────────────────────────────────────

suppressPackageStartupMessages({
  library(dplyr);      library(lubridate);  library(stringr);  library(readr)
  library(zoo);        library(fs);         library(glue);     library(tidyr)
  library(gridExtra);  library(ggplot2);    library(grid);     library(gt)
})

# -----------------------------
# 1) Paths and constants
# -----------------------------
input_root  <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized"
input_files <- c("ULSS7.csv", "ULSS8.csv")

base_output <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi_inventory_management"
out_dirs    <- list(
  EMAZ = fs::path(base_output, "Emazie",    "TO-BE"),
  PLAS = fs::path(base_output, "Plasma",    "TO-BE"),
  PIA  = fs::path(base_output, "Piastrine", "TO-BE")
)

# Model parameters
LT_days   <- 5
LT_weeks  <- LT_days / 7
GCS_level <- 0.95
z_value   <- qnorm(GCS_level)

# Directory creator (fs first; fallback to base)
ensure_dir <- function(p) {
  ok <- TRUE
  tryCatch(fs::dir_create(p), error = function(e) { ok <<- FALSE })
  if (!ok && !dir.exists(p)) dir.create(p, recursive = TRUE, showWarnings = FALSE)
  invisible(p)
}

# -----------------------------
# 2) Read & merge sources
# -----------------------------
read_one <- function(fname){
  readr::read_csv(fs::path(input_root, fname), col_types = cols(.default = "c")) %>%
    mutate(.file = fname)
}
df_all <- dplyr::bind_rows(lapply(input_files, read_one))

# -----------------------------
# 3) Facility normalization
# -----------------------------
norm_facility <- function(ospedale, reparto) {
  o <- stringr::str_to_upper(coalesce(ospedale, ""))
  r <- stringr::str_to_upper(coalesce(reparto,   ""))
  dplyr::case_when(
    str_detect(o, "SANTORSO") | str_detect(o, "ALTO\\s*VICENTINO") |
      str_detect(o, "ULSS\\s*4") | str_detect(o, "SCHIO") | str_detect(o, "THIENE") |
      str_detect(r, "\\bSOAP\\b") | str_detect(r, "\\bCORSIA\\b") ~ "SANTORSO",
    (str_detect(o, "BASSANO") | (str_detect(r, "BASSANO") & !str_detect(r, "ESTERNI"))) ~ "BASSANO",
    str_detect(o, "VICENZA") | str_detect(o, "ULSS\\s*8") | str_detect(o, "AZIENDA\\s*ULSS\\s*8") |
      str_detect(r, "\\bVIC\\.?\\b") ~ "VICENZA",
    str_detect(o, "ARZIGNANO") | str_detect(r, "\\bARZ\\.?|ARZIG") ~ "ARZIGNANO",
    str_detect(o, "VALDAGNO")  | str_detect(r, "\\bVALD\\.?")      ~ "VALDAGNO",
    TRUE ~ NA_character_
  )
}

df_all <- df_all %>%
  dplyr::rename(Group = `Gruppo AB0 - Rh`, TxDate = `Data Trasfusione`) %>%
  dplyr::mutate(
    TxDate         = lubridate::ymd(TxDate),
    Year           = lubridate::year(TxDate),
    Week           = lubridate::isoweek(TxDate),
    Facility_clean = norm_facility(Ospedale, Reparto)
  ) %>%
  dplyr::filter(!is.na(Facility_clean)) %>%
  dplyr::select(CDM, Year, Week, Facility_clean, Group, `EMC Emolife`, Reparto, .file)

facilities_ulss7 <- c("SANTORSO", "BASSANO")
facilities_ulss8 <- c("VICENZA", "ARZIGNANO", "VALDAGNO")

# -----------------------------
# 4) Weekly aggregation and table helpers
# -----------------------------
# Weekly counts per group (rolling window statistics computed on Transfusions)
agg_weekly <- function(df_sub){
  df_sub %>%
    dplyr::group_by(Group, Year, Week) %>%
    dplyr::summarise(Transfusions = dplyr::n(), .groups = "drop") %>%
    dplyr::arrange(Group, Week) %>%
    dplyr::group_by(Group) %>%
    dplyr::mutate(
      MA5     = zoo::rollapply(Transfusions, width = 5, FUN = mean, align = "right", fill = NA, na.rm = TRUE),
      SD_week = zoo::rollapply(Transfusions, width = 5, FUN = sd,   align = "right", fill = NA, na.rm = TRUE)
    ) %>% dplyr::ungroup()
}

# A4 layout proxy and PNG width control for thesis-like balance
.page_width_in    <- 8.27
.left_margin_pt   <- 3
.right_margin_pt  <- 3
.target_ppi       <- 96
.content_width_in <- .page_width_in - (.left_margin_pt + .right_margin_pt)/72
.content_width_px <- round(.content_width_in * .target_ppi)
.table_pct_width  <- 80  # percent of content width

# Thesis-style gt theme (no version-specific helpers)
gt_style_thesis <- function(gt_tbl,
                            title_px = 13, base_px = 11, label_px = 11,
                            font_family = c("Times New Roman","Liberation Serif","serif"),
                            table_pct_local = .table_pct_width) {
  gt_tbl %>%
    gt::tab_options(
      table.align = "center",
      table.width = gt::pct(table_pct_local),
      table.font.names = font_family,
      table.font.size  = gt::px(base_px),
      heading.title.font.size = gt::px(title_px),
      data_row.padding = gt::px(6),
      column_labels.border.top.width    = gt::px(1),
      column_labels.border.bottom.width = gt::px(1),
      column_labels.vlines.width        = gt::px(0),
      table.border.top.width            = gt::px(0),
      table.border.bottom.width         = gt::px(0),
      column_labels.background.color    = "white",
      heading.background.color          = "white"
    ) %>%
    gt::cols_align(align = "right", columns = gt::everything()) %>%
    gt::cols_align(align = "left",  columns = 1) %>%
    gt::tab_style(
      style = gt::cell_text(weight = "bold", size = gt::px(label_px)),
      locations = gt::cells_column_labels(gt::everything())
    ) %>%
    gt::opt_row_striping()
}

save_gt_png <- function(gt_tbl, path_png) {
  ensure_dir(dirname(path_png))
  gt::gtsave(gt_tbl, filename = path_png, vwidth = .content_width_px, expand = 0)
}

# Build a thesis-style table PNG. Input must have first column named "Metric".
thesis_table_png <- function(df_wide, title, out_png) {
  num_cols <- names(df_wide)[names(df_wide) != "Metric"]
  suppressWarnings({
    df_fmt <- df_wide
    for (cc in num_cols) {
      if (!is.numeric(df_fmt[[cc]])) {
        as_num <- suppressWarnings(as.numeric(df_fmt[[cc]]))
        if (!all(is.na(as_num))) df_fmt[[cc]] <- as_num
      }
    }
  })
  gt_tbl <- df_fmt %>%
    gt::gt() %>%
    gt::tab_header(title = title) %>%
    gt::fmt_number(columns = -Metric, rows = Metric %in% c("High","Medium","Low"),
                   decimals = 0, use_seps = TRUE) %>%
    gt::fmt_number(columns = -Metric, rows = Metric %in% c("ME(s)","RMSE(s)","ME(S)","RMSE(S)"),
                   decimals = 2, use_seps = TRUE) %>%
    gt_style_thesis()
  save_gt_png(gt_tbl, out_png)
}

# -----------------------------
# 5) Emitters (thesis PNG + CSV)
# -----------------------------
emit_rbc_ulss8 <- function(weekly, out_dir, facility, year_sel){
  weekly <- weekly %>%
    dplyr::mutate(
      ss       = ceiling(z_value * SD_week * sqrt(LT_weeks)),
      MA5      = ceiling(MA5),
      SD_week  = ceiling(SD_week),
      s        = ceiling(MA5 * LT_weeks + ss),
      S        = ceiling(MA5 * LT_weeks + z_value * SD_week * sqrt(LT_weeks * 2)),
      Crit1    = Transfusions > (MA5 + SD_week),
      Crit2    = Transfusions > S,
      Severity = dplyr::case_when(Crit1 & Crit2 ~ "High", Crit1 | Crit2 ~ "Medium", TRUE ~ "Low"),
      err_s    = Transfusions - s
    )
  ensure_dir(out_dir)
  readr::write_csv(
    weekly %>% dplyr::select(Year, Week, Group, Transfusions, MA5, SD_week, ss, s, S, Severity, err_s),
    fs::path(out_dir, glue::glue("model_{facility}_{year_sel}.csv"))
  )
  summary_tbl <- weekly %>%
    dplyr::group_by(Group) %>%
    dplyr::summarise(
      High   = sum(Severity == "High"),
      Medium = sum(Severity == "Medium"),
      Low    = sum(Severity == "Low"),
      `ME(s)`   = round(mean(err_s, na.rm = TRUE), 2),
      `RMSE(s)` = round(sqrt(mean(err_s^2, na.rm = TRUE)), 2), .groups = "drop"
    ) %>%
    tidyr::pivot_longer(-Group, names_to = "Metric", values_to = "Value") %>%
    tidyr::pivot_wider(names_from = Group, values_from = Value)
  thesis_table_png(summary_tbl,
                   glue::glue("TO-BE — {facility} — {year_sel}"),
                   fs::path(out_dir, glue::glue("summary_{facility}_{year_sel}.png")))
}

emit_rbc_ulss7 <- function(weekly, out_dir, facility, year_sel){
  weekly <- weekly %>%
    dplyr::mutate(
      LT_days  = LT_days, LT_weeks = LT_weeks, GCS = GCS_level, z = z_value,
      ss       = ceiling(z * SD_week * sqrt(LT_weeks)),
      MA5      = ceiling(MA5), SD_week = ceiling(SD_week),
      s        = ceiling(MA5 * LT_weeks + ss),
      S        = ceiling(MA5 * LT_weeks + z * SD_week * sqrt(LT_weeks * 2)),
      Criterion1 = Transfusions > (MA5 + SD_week),
      Criterion2 = Transfusions > S,
      Severity   = dplyr::case_when(Criterion1 & Criterion2 ~ "High", Criterion1 | Criterion2 ~ "Medium", TRUE ~ "Low"),
      err_S      = Transfusions - S
    )
  ensure_dir(out_dir)
  readr::write_csv(
    weekly %>% dplyr::select(Year, Week, Group, Transfusions, MA5, SD_week, LT_days, LT_weeks, GCS, z, ss, s, S, Criterion1, Criterion2, Severity, err_S),
    fs::path(out_dir, glue::glue("model_{facility}_{year_sel}.csv"))
  )
  metrics_raw <- weekly %>% dplyr::group_by(Group) %>%
    dplyr::summarise(
      High   = sum(Severity == "High",  na.rm = TRUE),
      Medium = sum(Severity == "Medium",na.rm = TRUE),
      Low    = sum(Severity == "Low",   na.rm = TRUE),
      `ME(S)`   = round(mean(err_S,    na.rm = TRUE), 2),
      `RMSE(S)` = round(sqrt(mean(err_S^2, na.rm = TRUE)), 2), .groups = "drop"
    )
  summary_table <- metrics_raw %>%
    tidyr::pivot_longer(-Group, names_to = "Metric", values_to = "Value") %>%
    tidyr::pivot_wider(names_from = Group, values_from = Value)
  thesis_table_png(summary_table,
                   glue::glue("TO-BE — {facility} — {year_sel}"),
                   fs::path(out_dir, glue::glue("summary_{facility}_{year_sel}.png")))
}

# -----------------------------
# 6) Batch processing
# -----------------------------
process_component <- function(component_code){
  if (component_code == "EMAZ") {
    out_root <- out_dirs$EMAZ
    filter_rx <- regex("emaz", ignore_case = TRUE)
  } else if (component_code == "PLAS") {
    out_root <- out_dirs$PLAS
    filter_rx <- regex("PLASMA", ignore_case = TRUE)
  } else if (component_code == "PIA") {
    out_root <- out_dirs$PIA
    filter_rx <- regex("PLT|PIASTRIN", ignore_case = TRUE)
  } else stop("Invalid component code")
  
  df_c <- df_all %>% dplyr::filter(stringr::str_detect(`EMC Emolife`, filter_rx))
  years <- sort(unique(df_c$Year))
  
  # ULSS8 facilities
  for (a in years) for (s in facilities_ulss8) {
    sub <- df_c %>% dplyr::filter(Year == a, Facility_clean == s)
    if (nrow(sub) == 0) next
    out_dir <- fs::path(out_root, as.character(a), s); ensure_dir(out_dir)
    weekly <- agg_weekly(sub)
    emit_rbc_ulss8(weekly, out_dir, s, a)
  }
  
  # ULSS7 facilities
  for (a in years) for (s in facilities_ulss7) {
    sub <- df_c %>% dplyr::filter(Year == a, Facility_clean == s)
    if (nrow(sub) == 0) next
    out_dir <- fs::path(out_root, as.character(a), s); ensure_dir(out_dir)
    weekly <- agg_weekly(sub)
    emit_rbc_ulss7(weekly, out_dir, s, a)
  }
}

# Ensure top-level TO-BE roots exist
ensure_dir(out_dirs$EMAZ)
ensure_dir(out_dirs$PLAS)
ensure_dir(out_dirs$PIA)

# Run
message(">>> TO-BE — RBCs")
process_component("EMAZ")
message(">>> TO-BE — Plasma")
process_component("PLAS")
message(">>> TO-BE — Platelets")
process_component("PIA")

message("TO-BE processing completed (RBC model applied to all components).")
