# ────────────────────────────────────────────────────────────────────────────────
# RBCs ULSS7+ULSS8 — Thesis PNG table (weekly summary + ME/RMSE) + control CSV
# ────────────────────────────────────────────────────────────────────────────────

suppressPackageStartupMessages({
  library(dplyr); library(lubridate); library(stringr); library(readr)
  library(fs); library(glue); library(tibble); library(tidyr); library(rlang)
  library(gt)
})

# ---------- Thesis PNG geometry ----------
page_width_in   <- 8.27
left_margin_pt  <- 3
right_margin_pt <- 3
target_ppi      <- 96
content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72
content_width_px <- round(content_width_in * target_ppi)
table_pct_width  <- 80

gt_style_thesis <- function(gt_tbl,
                            title_px = 13, base_px = 11, label_px = 11,
                            font_family = c("Times New Roman","Liberation Serif","serif"),
                            table_pct_local = table_pct_width) {
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
    cols_align(align = "left",  columns = 1) %>%
    cols_align(align = "right", columns = 2:last_col()) %>%
    opt_row_striping()
}
save_gt_png <- function(gt_tbl, path_png) {
  fs::dir_create(dirname(path_png), recurse = TRUE)
  gt::gtsave(gt_tbl, filename = path_png, vwidth = content_width_px, expand = 0)
}

# ---------- Parameters ----------
input_root  <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized"
output_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi_inventory_management/Emazie/AS-IS"
input_files <- c("ULSS7.csv", "ULSS8.csv")
facilities  <- c("SANTORSO","BASSANO","VICENZA","ARZIGNANO","VALDAGNO")
years_sel   <- c(2023,2024,2025)
groups_order <- c("O-","O+","A-","A+","B-","B+","AB-","AB+")

# Lookup (unchanged)
pct_scorta_tbl <- tribble(
  ~Facility,   ~Group, ~PctStockVsMedian,
  "VICENZA","A-",117.65, "VICENZA","A+",75.47, "VICENZA","B+",68.97, "VICENZA","O-",86.21,  "VICENZA","O+",69.44,
  "ARZIGNANO","A-",100.00, "ARZIGNANO","A+",81.25, "ARZIGNANO","B+",75.00, "ARZIGNANO","O-",87.50,  "ARZIGNANO","O+",90.91,
  "VALDAGNO","A-",133.33,  "VALDAGNO","A+",80.00,  "VALDAGNO","B+",100.00, "VALDAGNO","O-",75.00,   "VALDAGNO","O+",71.43,
  "SANTORSO","A-",71.43,   "SANTORSO","A+",75.86,  "SANTORSO","B+",100.00,  "SANTORSO","O-",76.92,   "SANTORSO","O+",76.92,
  "BASSANO","A-",85.71,    "BASSANO","A+",83.33,   "BASSANO","B+",83.33,    "BASSANO","O-",75.76,    "BASSANO","O+",70.59
)

`%||%` <- function(a,b) if (is.null(a)) b else a

# ---------- Read & filter RBCs (identical selection logic) ----------
read_one <- function(f){
  readr::read_csv(fs::path(input_root, f), col_types = cols(.default = "c")) %>%
    rename(Group = `Gruppo AB0 - Rh`) %>%
    mutate(
      Tx_Date = lubridate::ymd(`Data Trasfusione`),
      Year    = lubridate::year(Tx_Date),
      Facility_clean = dplyr::case_when(
        str_detect(Ospedale, regex("OSPEDALE\\s+ALTOVICENTINO\\s+SANTORSO", TRUE)) |
          str_detect(Ospedale, regex("AZIENDA\\s+ULSS\\s*4\\s+ALTO\\s+VICENTINO.*SCHIO.*THIENE", TRUE)) ~ "SANTORSO",
        str_detect(Reparto %||% "", regex("BASSANO", TRUE)) &
          !str_detect(Reparto %||% "", regex("ESTERNI", TRUE)) ~ "BASSANO",
        str_detect(Ospedale, regex("ULSS\\s*6", TRUE)) ~ "VICENZA",
        str_detect(Ospedale, regex("ARZIGNANO", TRUE)) ~ "ARZIGNANO",
        str_detect(Ospedale, regex("VALDAGNO",  TRUE)) ~ "VALDAGNO",
        TRUE ~ NA_character_
      )
    ) %>%
    filter(
      !is.na(Facility_clean),
      str_detect(`EMC Emolife`, regex("(\\bemaz|\\beritr|\\bRBC\\b|\\b25\\d{2}\\b)", TRUE)) &
        !str_detect(`EMC Emolife`, regex("plasm|\\bplt\\b|piastr|buffy", TRUE))
    )
}
df <- bind_rows(lapply(input_files, read_one))

# ---------- Build matrix + export ----------
for (yy in years_sel) {
  for (fac in facilities) {
    sub_df <- df %>% filter(Year == yy, Facility_clean == fac)
    if (nrow(sub_df) == 0) next
    
    out_dir <- fs::path(output_root, as.character(yy), fac)
    fs::dir_create(out_dir, recurse = TRUE)
    
    # Weekly aggregation (observed ISO weeks)
    weekly_df <- sub_df %>%
      mutate(Date = as.Date(Tx_Date),
             Year = year(Date),
             Week = isoweek(Date)) %>%
      group_by(Facility = Facility_clean, Group, Year, Week) %>%
      summarise(Transfusions_Total = n(), .groups = "drop")
    
    # Core summary — FIX: avoid name masking by using a different temporary name,
    # then rename to 'Transfusions_Total' afterwards.
    summary_df <- weekly_df %>%
      filter(Year == yy, Facility == fac) %>%
      group_by(Facility, Group) %>%
      summarise(
        Weekly_Total_sum    = sum(Transfusions_Total),                # <- somma settimanali
        Weekly_Mean         = round(mean(Transfusions_Total), 1),     # <- media settimanale
        Weekly_Median       = median(Transfusions_Total),             # <- mediana settimanale
        .groups = "drop"
      ) %>%
      rename(Transfusions_Total = Weekly_Total_sum) %>%
      left_join(pct_scorta_tbl, by = c("Facility","Group")) %>%
      mutate(
        Stock_5d  = round(Weekly_Median / 7 * 5, 1),
        Min_Stock = if_else(!is.na(PctStockVsMedian),
                            ceiling(Weekly_Median * PctStockVsMedian / 100),
                            NA_real_)
      ) %>%
      select(Group, Transfusions_Total, Weekly_Mean, Weekly_Median, Stock_5d, Min_Stock)
    
    if (nrow(summary_df) == 0) next
    
    # Total column: only the sum of Transfusions_Total; other metrics blank
    total_row <- tibble(
      Group = "Total",
      Transfusions_Total = sum(summary_df$Transfusions_Total, na.rm = TRUE),
      Weekly_Mean   = NA_real_,
      Weekly_Median = NA_real_,
      Stock_5d      = NA_real_,
      Min_Stock     = NA_real_
    )
    
    # Weekly model for err_s (unchanged)
    median_tbl <- weekly_df %>%
      group_by(Facility, Group) %>%
      summarise(Weekly_Median = median(Transfusions_Total), .groups = "drop")
    
    model_weekly <- weekly_df %>%
      left_join(median_tbl,     by = c("Facility","Group")) %>%
      left_join(pct_scorta_tbl, by = c("Facility","Group")) %>%
      mutate(
        Min_Stock = ceiling(Weekly_Median * PctStockVsMedian / 100),
        err_s     = Transfusions_Total - Min_Stock
      ) %>%
      filter(Facility == fac, Year == yy)
    
    err_stats <- model_weekly %>%
      group_by(Group) %>%
      summarise(
        ME_s   = mean(err_s, na.rm = TRUE),
        RMSE_s = sqrt(mean(err_s^2, na.rm = TRUE)),
        .groups = "drop"
      )
    
    # Assemble matrix (rows = metrics; cols = groups + Total)
    mtx <- bind_rows(summary_df, total_row) %>%
      column_to_rownames("Group") %>%
      t() %>% as.data.frame(check.names = FALSE)
    
    # Reorder columns to canonical order + Total at the end
    col_order <- c(intersect(groups_order, colnames(mtx)),
                   setdiff(colnames(mtx), groups_order))
    mtx <- mtx[, col_order, drop = FALSE]
    
    # Append ME_s / RMSE_s (Total blank)
    me_row   <- setNames(rep(NA_real_, ncol(mtx)), colnames(mtx))
    rmse_row <- me_row
    for (g in err_stats$Group) if (g %in% colnames(mtx)) {
      me_row[g]   <- err_stats$ME_s[err_stats$Group == g]
      rmse_row[g] <- err_stats$RMSE_s[err_stats$Group == g]
    }
    mtx2 <- rbind(mtx, `ME_s` = me_row, `RMSE_s` = rmse_row)
    
    # ---- Control CSV (exact matrix used for the PNG) ----
    csv_df <- tibble(Metric = rownames(mtx2)) %>%
      bind_cols(as_tibble(mtx2, .name_repair = "minimal"))
    readr::write_csv(csv_df, fs::path(out_dir, glue("weekly_summary_{fac}_{yy}.csv")))
    
    # ---- Thesis PNG ----
    tab_gt <- csv_df %>%
      gt() %>%
      tab_header(title = glue("{fac} — {yy}: Weekly summary by group")) %>%
      fmt_number(columns = 2:last_col(),
                 rows = Metric %in% c("Transfusions_Total","Weekly_Median","Min_Stock"),
                 decimals = 0, use_seps = TRUE) %>%
      fmt_number(columns = 2:last_col(),
                 rows = Metric %in% c("Weekly_Mean","Stock_5d"),
                 decimals = 1, use_seps = TRUE) %>%
      fmt_number(columns = 2:last_col(),
                 rows = Metric %in% c("ME_s","RMSE_s"),
                 decimals = 2, use_seps = TRUE) %>%
      gt_style_thesis()
    
    save_gt_png(tab_gt, fs::path(out_dir, glue("weekly_summary_{fac}_{yy}.png")))
  }
}

message("Done: thesis-style PNG + control CSV created per Facility × Year.")
