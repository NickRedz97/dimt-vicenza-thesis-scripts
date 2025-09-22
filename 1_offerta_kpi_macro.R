# =============================================================================
# Script name : 01_offerta_kpi_macro.R
# Purpose     : KPI dashboard by macro (Emazie / Plasma / Piastrine), overall
#               scope (no ULSS split). Produces monthly plots and annual tables.
# Outputs     : For each macro:
#                 - Monthly flows (stacked)
#                 - Monthly flows (time series)
#                 - Flows by component (stacked, facet by flow)
#                 - Wastage rate (%)
#               Plus annual PNG/CSV tables of collection by blood group
#               (rows = months; 2025 up to May).
# Notes       : Plasma excludes PLT/4305/"piastr*"; Platelets includes 4305/PLT/buffy;
#               macro "Altro" removed; plasma-like labels excluded from Emazie/Piastrine plots.
# =============================================================================

suppressPackageStartupMessages({
  pkgs <- c("tidyverse","lubridate","janitor","scales","zoo","fs","glue","stringr","rlang","gt")
  for(p in pkgs) if(!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(tidyverse); library(lubridate); library(janitor)
  library(scales); library(zoo); library(fs); library(glue)
  library(stringr); library(rlang); library(gt)
})

# -----------------------------------------------------------------------------
# PATHS
# -----------------------------------------------------------------------------
dir_raccolta <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/raccolta/normalized"
dir_trasfuso <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized"
dir_scartato <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/scartato/normalized"
dir_cessioni <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/cessioni/normalized"
file_cessioni <- file.path(dir_cessioni, "cessioni_2023_2025.csv")

out_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi complessive/Offerta_KPI_Macro"

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
log_msg <- function(...) cat(format(Sys.time(), "[%Y-%m-%d %H:%M:%S]"), ..., "\n")

ensure_dir <- function(p) { if (!dir.exists(p)) dir.create(p, recursive = TRUE, showWarnings = FALSE); invisible(p) }

safe_read <- function(path, guess_max = 200000) {
  stopifnot(!is.na(path), file.exists(path))
  readr::read_csv(path, guess_max = guess_max, show_col_types = FALSE) %>% clean_names()
}

ensure_date <- function(x) {
  if(inherits(x,"Date")) return(x)
  if(inherits(x,"POSIXct")||inherits(x,"POSIXt")) return(as.Date(x))
  suppressWarnings(as.Date(x))
}

normalize_group <- function(x) {
  x <- as.character(x)
  x[x %in% c("", "NA", "na")] <- NA_character_
  x <- toupper(trimws(x))
  x <- str_replace_all(x, "POSITIVO", "+")
  x <- str_replace_all(x, "NEGATIVO", "-")
  x <- str_replace_all(x, "\\s+", "")
  case_when(
    x %in% c("A+","A-","B+","B-","O+","O-","AB+","AB-") ~ x,
    TRUE ~ NA_character_
  )
}

std_emocomp <- function(x) {
  xp <- as.character(x)
  code <- str_match(xp, "^(\\d{1,4})\\s*-")[,2]
  desc <- xp %>% toupper() %>% str_replace("^\\s*\\d{1,4}\\s*-\\s*", "") %>% str_squish()
  out  <- ifelse(!is.na(code), paste(code, "-", desc), desc)
  trimws(out)
}

# Recognizers
is_whole_blood <- function(x) { x0 <- tolower(x); str_detect(x0, "sangue\\s*intero") }
is_rbc <- function(x) { x0 <- tolower(x); str_detect(x0, "emaz|eritro|\\brbc\\b|\\b250(4)?\\b|leucodeplet") & !str_detect(x0, "plasm|\\bplt\\b|piastr|buffy") }
is_plasma <- function(x) { x0 <- tolower(x); str_detect(x0, "plasm") & !str_detect(x0, "\\bplt\\b|piastr|buffy|\\b4305\\b") }
is_platelet <- function(x) { x0 <- tolower(x); str_detect(x0, "\\b4305\\b|\\bplt\\b|piastr|platelet|buffy") }

is_plasma_std <- function(x_std) { xs <- toupper(as.character(x_std)); str_detect(xs, "\\bPLASMA\\b|DEPLASM") }

macro_from_component <- function(x) {
  case_when(
    is_rbc(x)        ~ "Emazie",
    is_plasma(x)     ~ "Plasma",
    is_platelet(x)   ~ "Piastrine",
    TRUE             ~ NA_character_
  )
}

legend_compact <- theme(
  legend.position = "bottom",
  legend.text = element_text(size = 7),
  legend.key.width = grid::unit(0.65, "lines"),
  legend.key.height = grid::unit(0.65, "lines"),
  legend.box.margin = margin(t=2, r=2, b=2, l=2),
  plot.margin = margin(10, 26, 10, 10)
)

shorten_label <- function(s, max_chars = 26) {
  s <- as.character(s)
  ifelse(nchar(s) > max_chars, paste0(substr(s,1,max_chars-1),"…"), s)
}

# -----------------------------------------------------------------------------
# WORD/A4 SIZING 
# -----------------------------------------------------------------------------
page_width_in   <- 8.27   # A4 portrait width
left_margin_pt  <- 3      # side margins in points
right_margin_pt <- 3
target_ppi      <- 96     # Word inserts at exact size at 96 ppi
content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72

# Suggested heights (inches) for readability at content width
h_stacked_in   <- 5.8
h_timeseries_in<- 5.8
h_wastage_in   <- 4.8
h_bycomp_in    <- 8.6

# Thesis table style (gt)
table_pct <- 80  # table body occupies 80% of canvas width (balanced lateral breathing room)

gt_style_thesis <- function(gt_tbl,
                            title_px = 13, base_px = 11, label_px = 11,
                            font_family = c("Times New Roman","Liberation Serif","serif"),
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
  ensure_dir(dirname(path))
  gt::gtsave(
    gt_tbl,
    file   = path,
    vwidth = round(content_width_in * target_ppi),  # canvas = text width
    expand = 0                                      # no external whitespace
  )
}

# -----------------------------------------------------------------------------
# LOAD FILES
# -----------------------------------------------------------------------------
file_raccolta <- file.path(dir_raccolta, "raccolta_merged_clean_with_group.csv")
file_ulss7    <- file.path(dir_trasfuso, "ULSS7.csv")
file_ulss8    <- file.path(dir_trasfuso, "ULSS8.csv")
file_scartato <- file.path(dir_scartato, "scartato_merged_cleaned.csv")
stopifnot(file.exists(file_raccolta), file.exists(file_ulss7), file.exists(file_ulss8),
          file.exists(file_scartato), file.exists(file_cessioni))

log_msg("Reading raccolta: ", file_raccolta)
raccolta <- safe_read(file_raccolta)

log_msg("Reading trasfuso: ULSS7 + ULSS8")
ulss7 <- safe_read(file_ulss7) %>% mutate(source = "ULSS7")
ulss8 <- safe_read(file_ulss8) %>% mutate(source = "ULSS8")
trasfuso <- bind_rows(ulss7, ulss8)

log_msg("Reading scartato: ", file_scartato)
scartato <- safe_read(file_scartato)

log_msg("Reading cessioni: ", file_cessioni)
cessioni <- safe_read(file_cessioni)

# -----------------------------------------------------------------------------
# COLUMN DETECTION (SCARTATO)
# -----------------------------------------------------------------------------
date_candidates_sct <- c(
  "data_consumo_std","data_consumo","data","data_movimento",
  "data_cons","data_consumi","data_prelievo","data_prestazione_date"
)
qty_candidates_sct <- c(
  "numero_emocomponenti","numero_unita","quantita","qta","pezzi","unita",
  "n_componenti","n","numero_sacche","sacche"
)
grp_candidates_sct <- c("gruppo_ab0_rh_std","gruppo_ab0_rh")

col_date_sct <- intersect(date_candidates_sct, names(scartato))
if (length(col_date_sct) == 0) stop("No date column found in 'scartato'.")
date_sym_sct <- rlang::sym(col_date_sct[1])

col_qty_sct <- intersect(qty_candidates_sct, names(scartato))
qty_sym_sct <- if (length(col_qty_sct) == 0) NULL else rlang::sym(col_qty_sct[1])

col_grp_sct <- intersect(grp_candidates_sct, names(scartato))
grp_sym_sct <- if (length(col_grp_sct) == 0) NULL else rlang::sym(col_grp_sct[1])

# -----------------------------------------------------------------------------
# DATA PREP
# -----------------------------------------------------------------------------
rac <- raccolta %>%
  mutate(
    data_prestazione_date = ensure_date(data_prestazione_date),
    numero_emocomponenti  = suppressWarnings(as.numeric(numero_emocomponenti)),
    gruppo_ab0_rh         = normalize_group(gruppo_ab0_rh),
    emocomponente_std     = std_emocomp(emocomponente),
    month                 = floor_date(data_prestazione_date, "month")
  ) %>%
  filter(toupper(stato_donazione) == "CONCLUSA")

trf <- trasfuso %>%
  mutate(
    data_trasfusione  = ensure_date(data_trasfusione),
    gruppo_ab0_rh     = normalize_group(gruppo_ab0_rh),
    macro             = macro_from_component(emc_emolife),
    emocomponente_std = std_emocomp(emc_emolife),
    month             = floor_date(data_trasfusione, "month")
  ) %>% filter(!is.na(macro))

sct <- scartato %>%
  mutate(
    data_consumo_std  = ensure_date(!!date_sym_sct),
    sct_qty           = if (is.null(qty_sym_sct)) 1 else suppressWarnings(as.numeric(!!qty_sym_sct)),
    gruppo_ab0_rh     = if (is.null(grp_sym_sct)) NA_character_ else normalize_group(as.character(!!grp_sym_sct)),
    macro             = macro_from_component(emocomponente),
    emocomponente_std = std_emocomp(emocomponente),
    month             = floor_date(data_consumo_std, "month")
  ) %>% filter(!is.na(macro))

css <- cessioni %>%
  mutate(
    data              = ensure_date(data),
    numero_unita_cedute = suppressWarnings(as.numeric(numero_unita_cedute)),
    macro             = macro_from_component(emocomponente),
    emocomponente_std = std_emocomp(emocomponente),
    month             = floor_date(data, "month")
  ) %>% filter(!is.na(macro))

# -----------------------------------------------------------------------------
# MONTHLY KPI BY MACRO
# -----------------------------------------------------------------------------
macro_levels <- c("Emazie","Plasma","Piastrine")

months_all <- seq.Date(
  from = min(c(rac %>% pull(month), trf$month, sct$month, css$month), na.rm = TRUE),
  to   = max(c(rac %>% pull(month), trf$month, sct$month, css$month), na.rm = TRUE),
  by = "month"
)
months_df <- tibble(month = months_all)

# Annual tables (collection by group × month)
save_group_tables_for_year <- function(mc, year_sel, rac_sel, dir_csv, dir_tab) {
  months_year <- tibble(month = seq.Date(as.Date(paste0(year_sel,"-01-01")),
                                         as.Date(paste0(year_sel,"-12-01")), by = "month"))
  if (year_sel == 2025) months_year <- months_year %>% filter(month < as.Date("2025-06-01"))
  
  rac_grp <- rac_sel %>%
    filter(month %in% months_year$month) %>%
    group_by(gruppo_ab0_rh, month) %>%
    summarise(units = sum(numero_emocomponenti, na.rm = TRUE), .groups = "drop")
  
  df_wide <- months_year %>%
    left_join(rac_grp, by = "month") %>%
    mutate(gruppo_ab0_rh = ifelse(is.na(gruppo_ab0_rh), "NA", gruppo_ab0_rh)) %>%
    group_by(month, gruppo_ab0_rh) %>%
    summarise(units = sum(units, na.rm = TRUE), .groups = "drop") %>%
    mutate(Mese = format(month, "%Y-%m"),
           mese_num = month(month)) %>%
    select(Mese, mese_num, gruppo_ab0_rh, units) %>%
    tidyr::pivot_wider(names_from = gruppo_ab0_rh, values_from = units, values_fill = list(units = 0)) %>%
    arrange(mese_num, Mese) %>%
    select(-mese_num)
  
  group_order <- c("O-","O+","A-","A+","B-","B+","AB-","AB+","NA")
  existing <- intersect(group_order, names(df_wide))
  other    <- setdiff(names(df_wide), c("Mese", existing))
  df_wide  <- df_wide %>% select(Mese, all_of(existing), all_of(other))
  
  ensure_dir(dir_csv); ensure_dir(dir_tab)
  csv_path <- file.path(dir_csv, glue("{mc}_by_group_{year_sel}.csv"))
  png_path <- file.path(dir_tab, glue("{mc}_by_group_{year_sel}.png"))
  
  # CSV (Italian header 'Mese' as requested)
  readr::write_csv(df_wide, csv_path)
  
  # PNG (English 'Month' for captioning consistency)
  df_wide_png <- df_wide %>% rename(Month = Mese)
  tab <- df_wide_png %>%
    gt() %>%
    tab_header(title = glue("{mc} — Monthly collection by ABO/Rh — {year_sel}")) %>%
    fmt_number(columns = where(is.numeric), decimals = 0, use_seps = TRUE) %>%
    gt_style_thesis()
  save_gt_png(tab, png_path)
  
  log_msg(glue("[{mc}][{year_sel}] group tables saved (rows={nrow(df_wide_png)})"))
}

for (mc in macro_levels) {
  log_msg("== Macro: ", mc, " ==")
  out_macro_dir <- file.path(out_root, mc)
  dir_csv   <- file.path(out_macro_dir, "file csv")
  dir_plot  <- file.path(out_macro_dir, "grafici")
  dir_tab   <- file.path(out_macro_dir, "tabelle png")
  invisible(lapply(c(dir_csv, dir_plot, dir_tab), ensure_dir))
  
  # RACCOLTA selection per macro
  rac_m <- rac %>%
    filter(
      (mc == "Emazie"    & (is_whole_blood(emocomponente) | is_rbc(emocomponente))) |
        (mc == "Plasma"    & (is_whole_blood(emocomponente) | is_plasma(emocomponente))) |
        (mc == "Piastrine" & (is_whole_blood(emocomponente) | is_platelet(emocomponente)))
    )
  
  trf_m <- trf %>% filter(macro == mc)
  sct_m <- sct %>% filter(macro == mc)
  css_m <- css %>% filter(macro == mc)
  
  log_msg(glue("[{mc}] Rows -> rac:{nrow(rac_m)} trf:{nrow(trf_m)} sct:{nrow(sct_m)} css:{nrow(css_m)}"))
  
  # Monthly aggregations
  agg_rac <- rac_m %>%
    group_by(month) %>%
    summarise(
      components_collected = sum(numero_emocomponenti, na.rm = TRUE),
      sacs_collected = n_distinct(cdm),
      .groups = "drop"
    )
  
  agg_trf <- trf_m %>%
    group_by(month) %>%
    summarise(
      transf_rows = n(),
      sacs_transfused = n_distinct(cdm),
      .groups = "drop"
    )
  
  has_cdm_sct <- "cdm" %in% names(sct_m)
  agg_sct <- sct_m %>%
    group_by(month) %>%
    summarise(
      components_wasted = sum(sct_qty, na.rm = TRUE),
      sacs_wasted = if (has_cdm_sct) n_distinct(cdm) else NA_integer_,
      .groups = "drop"
    )
  
  agg_css <- css_m %>%
    group_by(month) %>%
    summarise(
      units_ceded = sum(numero_unita_cedute, na.rm = TRUE),
      .groups = "drop"
    )
  
  monthly <- months_df %>%
    left_join(agg_rac, by = "month") %>%
    left_join(agg_sct, by = "month") %>%
    left_join(agg_css, by = "month") %>%
    left_join(agg_trf, by = "month") %>%
    replace_na(list(
      components_collected = 0, sacs_collected = 0,
      components_wasted = 0, sacs_wasted = 0,
      units_ceded = 0, transf_rows = 0, sacs_transfused = 0
    )) %>%
    arrange(month) %>%
    mutate(
      wastage_rate_pct = ifelse(components_collected > 0,
                                round(components_wasted / components_collected * 100, 2),
                                NA_real_)
    )
  
  # CSV export
  readr::write_csv(monthly, file.path(dir_csv, glue("{mc}_monthly_kpi.csv")))
  
  # ------
  # PLOTS 
  # ------
  serie_labels <- c(
    components_collected = "Collected (units)",
    transf_rows          = "Transfused (rows)",
    components_wasted    = "Wasted (units)",
    units_ceded          = "Ceded (units)"
  )
  pal <- c(
    "Collected (units)"  = "#1b9e77",
    "Transfused (rows)"  = "#7570b3",
    "Wasted (units)"     = "#d95f02",
    "Ceded (units)"      = "#e7298a"
  )
  
  # (1) Monthly flows — stacked
  df_flow <- monthly %>%
    select(month, components_collected, transf_rows, components_wasted, units_ceded) %>%
    pivot_longer(-month, names_to = "serie", values_to = "val") %>%
    mutate(serie = recode(serie, !!!serie_labels),
           tipo = factor(serie, levels = names(pal)))
  p_bar <- ggplot(df_flow, aes(x = month, y = val, fill = tipo)) +
    geom_col(position = "stack") +
    scale_fill_manual(values = pal, name = NULL, guide = guide_legend(nrow = 2)) +
    scale_y_continuous(labels = comma) +
    labs(title = glue("{mc} — Monthly flows (stacked)"),
         x = "Month", y = "Value") +
    theme_minimal() + legend_compact +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  ggsave(file.path(dir_plot, glue("{mc}_flows_stacked.png")), p_bar,
         width = content_width_in, height = h_stacked_in, units = "in", dpi = target_ppi)
  
  # (2) Monthly flows — time series
  df_ts <- monthly %>%
    pivot_longer(cols = c(components_collected, transf_rows, components_wasted, units_ceded),
                 names_to = "serie", values_to = "val") %>%
    mutate(serie = recode(serie, !!!serie_labels))
  p_ts <- ggplot(df_ts, aes(x = month, y = val, colour = serie)) +
    geom_line(linewidth = 1, na.rm = TRUE) +
    scale_colour_manual(values = pal, name = NULL, guide = guide_legend(nrow = 2)) +
    scale_y_continuous(labels = comma) +
    labs(title = glue("{mc} — Monthly flows (time series)"),
         x = "Month", y = "Value") +
    theme_minimal() + legend_compact +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  ggsave(file.path(dir_plot, glue("{mc}_timeseries.png")), p_ts,
         width = content_width_in, height = h_timeseries_in, units = "in", dpi = target_ppi)
  
  # (3) Wastage rate (%)
  p_w <- ggplot(monthly, aes(x = month, y = wastage_rate_pct)) +
    geom_col(fill = "#d95f02") +
    labs(title = glue("{mc} — Wastage rate (%) by month"),
         x = "Month", y = "%") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.margin = margin(10, 26, 10, 10))
  ggsave(file.path(dir_plot, glue("{mc}_wastage_pct.png")), p_w,
         width = content_width_in, height = h_wastage_in, units = "in", dpi = target_ppi)
  
  # (4) Flows by component (stacked per component, facet by flow)
  flow_by_emc <- bind_rows(
    rac_m %>% count(month, emocomponente_std, wt = numero_emocomponenti, name = "Collected"),
    trf_m %>% count(month, emocomponente_std, name = "Transfused"),
    sct_m %>% count(month, emocomponente_std, wt = sct_qty,             name = "Wasted"),
    css_m %>% count(month, emocomponente_std, wt = numero_unita_cedute, name = "Ceded")
  ) %>%
    pivot_longer(cols = c(Collected, Transfused, Wasted, Ceded),
                 names_to = "flow", values_to = "val") %>%
    filter(!is.na(val)) %>%
    group_by(month, emocomponente_std, flow) %>%
    summarise(val = sum(val, na.rm = TRUE), .groups = "drop")
  
  if (mc != "Plasma") {
    flow_by_emc <- flow_by_emc %>%
      mutate(is_plasma_like = is_plasma_std(emocomponente_std)) %>%
      filter(!is_plasma_like) %>%
      select(-is_plasma_like)
  }
  
  if (nrow(flow_by_emc)) {
    flow_by_emc <- flow_by_emc %>% mutate(emc_short = shorten_label(emocomponente_std, 26))
    p_comp <- ggplot(flow_by_emc, aes(x = month, y = val, fill = emc_short)) +
      geom_col(position = "stack") +
      facet_wrap(~flow, ncol = 1, scales = "free_y") +
      labs(title = glue("{mc} — Flows by component (stacked)"),
           x = "Month", y = "Value", fill = "Component") +
      theme_minimal() +
      theme(legend.position = "bottom",
            legend.text = element_text(size = 6.5),
            legend.key.width = grid::unit(0.6, "lines"),
            legend.key.height = grid::unit(0.6, "lines"),
            legend.box.margin = margin(2,2,2,2),
            plot.margin = margin(10, 28, 10, 10)) +
      guides(fill = guide_legend(nrow = 3))
    ggsave(file.path(dir_plot, glue("{mc}_flows_by_component.png")), p_comp,
           width = content_width_in, height = h_bycomp_in, units = "in", dpi = target_ppi)
  }
  
  # Annual group tables (PNG + CSV)
  for (yy in c(2023, 2024, 2025)) {
    save_group_tables_for_year(mc, yy, rac_m, dir_csv, dir_tab)
  }
}

log_msg("DONE. Outputs in: ", out_root)
