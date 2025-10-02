# Script R per l'analisi degli scarti (Excel 2023–2025) — mantiene tutte le colonne originali
# Output: CSV canonici + grafici + tabelle PNG in stile "thesis" (uniforme con gli altri script)
# NOTE VISIVE: niente titoli nei grafici, legende e assi più leggibili, date leggibili (EN),
#              niente spazio vuoto sopra/sotto; dimensioni ottimizzate per Word (A4, margini 3 pt)

suppressPackageStartupMessages({
  library(readxl);  library(readr);  library(dplyr);  library(tidyr);  library(janitor)
  library(lubridate); library(ggplot2); library(gridExtra); library(grid); library(stringr)
  library(scales); library(gt)
})

# -------------------------------------------------------------------
# Percorsi
# -------------------------------------------------------------------
base_path      <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/scartato"
dataset_dir    <- file.path(base_path, "dataset")
outputs_dir    <- file.path(base_path, "outputs")
normalized_dir <- file.path(base_path, "normalized")
dir.create(outputs_dir,    recursive = TRUE, showWarnings = FALSE)
dir.create(normalized_dir, recursive = TRUE, showWarnings = FALSE)

# Trova gli Excel "scartato YYYY.xlsx"
files <- list.files(dataset_dir, pattern = "^scartato\\s+\\d{4}\\.xlsx$", full.names = TRUE)
if (length(files) == 0) stop("Nessun file 'scartato YYYY.xlsx' trovato in ", dataset_dir)
years <- as.integer(gsub("^.*\\s(\\d{4})\\.xlsx$", "\\1", basename(files)))
ord   <- order(years)
files <- files[ord]; years <- years[ord]
message("File trovati: ", paste(basename(files), collapse = ", "))

# -------------------------------------------------------------------
# Word/A4 sizing — PNG esattamente alla larghezza del testo (A4 portrait, margini 3 pt)
# -------------------------------------------------------------------
page_width_in   <- 8.27
left_margin_pt  <- 3
right_margin_pt <- 3
target_ppi      <- 96
content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72

# -------------------------------------------------------------------
# Helpers parsing/normalization
# -------------------------------------------------------------------
parse_date_robust <- function(x) {
  if (inherits(x, "Date"))    return(as.Date(x))
  if (inherits(x, "POSIXct") || inherits(x, "POSIXt")) return(as.Date(x))
  if (is.numeric(x))          return(as.Date(x, origin = "1899-12-30")) # seriale Excel
  x_chr <- as.character(x)
  dt <- suppressWarnings(lubridate::parse_date_time(
    x_chr,
    orders = c("Y-m-d","d/m/Y","m/d/Y","d-m-Y","Ymd","d/m/Y HMS","Y-m-d HMS","d-m-Y HMS"),
    tz = "UTC"
  ))
  as.Date(dt)
}
detect_col <- function(nms, candidates) {
  hit <- intersect(candidates, nms)
  if (length(hit) == 0) NA_character_ else hit[1]
}

# Arricchisce SENZA rimuovere colonne
enrich_keep_all_cols <- function(df_raw) {
  df <- df_raw %>% clean_names()
  nms <- names(df)
  date_col  <- detect_col(nms, c("data_consumo","data_cons","data","dataconsumo","data_movimento","data_prestazione","data_prelievo"))
  grp_col   <- detect_col(nms, c("gruppo_ab0_rh","gruppo_abo_rh","gruppo_ab0","gruppo_abo","ab0_rh","abo_rh","ab0","abo"))
  emc_col   <- detect_col(nms, c("emocomponente","emo_componente","emocomp","emc_emolife"))
  tipo_col  <- detect_col(nms, c("tipo_consumo","motivo_scarto","causa_scarto","tipo"))
  simt_col  <- detect_col(nms, c("simt","struttura","simt_origine","centro","reparto","simt_di_appartenenza","simt_di_provenienza"))
  
  if (!is.na(date_col)) {
    if ("data_consumo" %in% nms) df <- df %>% mutate(data_consumo = parse_date_robust(.data[["data_consumo"]]))
    else df <- df %>% mutate(data_consumo = parse_date_robust(.data[[date_col]]))
  } else df <- df %>% mutate(data_consumo = as.Date(NA))
  
  if (!is.na(grp_col)) {
    if ("gruppo_ab0_rh" %in% nms) df <- df %>% mutate(gruppo_ab0_rh = as.character(.data[["gruppo_ab0_rh"]]))
    else df <- df %>% mutate(gruppo_ab0_rh = as.character(.data[[grp_col]]))
  } else if (!("gruppo_ab0_rh" %in% nms)) df <- df %>% mutate(gruppo_ab0_rh = NA_character_)
  
  if (!is.na(emc_col)) {
    if ("emocomponente" %in% nms) df <- df %>% mutate(emocomponente = as.character(.data[["emocomponente"]]))
    else df <- df %>% mutate(emocomponente = as.character(.data[[emc_col]]))
  } else if (!("emocomponente" %in% nms)) df <- df %>% mutate(emocomponente = NA_character_
  )
  
  if (!is.na(tipo_col)) {
    if ("tipo_consumo" %in% nms) df <- df %>% mutate(tipo_consumo = as.character(.data[["tipo_consumo"]]))
    else df <- df %>% mutate(tipo_consumo = as.character(.data[[tipo_col]]))
  } else if (!("tipo_consumo" %in% nms)) df <- df %>% mutate(tipo_consumo = NA_character_)
  
  if (!is.na(simt_col)) {
    if ("simt" %in% nms) df <- df %>% mutate(simt = as.character(.data[["simt"]]))
    else df <- df %>% mutate(simt = as.character(.data[[simt_col]]))
  } else if (!("simt" %in% nms)) df <- df %>% mutate(simt = NA_character_)
  
  df <- df %>% mutate(
    gruppo_ab0_rh = ifelse(is.na(gruppo_ab0_rh), NA_character_, stringr::str_squish(as.character(gruppo_ab0_rh))),
    emocomponente = ifelse(is.na(emocomponente), NA_character_, stringr::str_squish(as.character(emocomponente))),
    tipo_consumo  = ifelse(is.na(tipo_consumo),  NA_character_, stringr::str_squish(as.character(tipo_consumo))),
    simt          = ifelse(is.na(simt),          NA_character_, stringr::str_squish(as.character(simt)))
  )
  df
}

# -------------------------------------------------------------------
# Lettura e unione
# -------------------------------------------------------------------
scarti_list <- vector("list", length(files))
for (i in seq_along(files)) {
  message("Leggo: ", files[i])
  raw <- read_excel(files[i], sheet = 1)
  tmp <- enrich_keep_all_cols(raw) %>% mutate(anno_file = years[i])
  scarti_list[[i]] <- tmp
}
df <- bind_rows(scarti_list)

# CSV canonici (tutte le colonne)
write_csv(df, file.path(normalized_dir, "scartato_merged_cleaned.csv"))
write_csv(df, file.path(normalized_dir, "scartato_2023_2025.csv"))
message("CSV canonici scritti in ", normalized_dir)

# -------------------------------------------------------------------
# Feature engineering (non rimuove colonne)
# -------------------------------------------------------------------
df <- df %>%
  mutate(
    year    = lubridate::year(data_consumo),
    week    = lubridate::isoweek(data_consumo),
    month   = floor_date(data_consumo, "month"),
    quarter = paste0(lubridate::year(data_consumo), "-Q", lubridate::quarter(data_consumo))
  )

# Weekly summary CSV (coerente con altri script)
summary_df <- df %>%
  group_by(year, week, gruppo_ab0_rh, emocomponente, tipo_consumo, simt) %>%
  summarise(count = n(), .groups = 'drop')
write_csv(summary_df, file.path(outputs_dir, "scarti_summary.csv"))

# -------------------------------------------------------------------
# STILI: grafici + TABELLE PNG THESIS (uniformati, senza titoli)
# -------------------------------------------------------------------
pal_macro <- c("RBCs"="#0072B2","Plasma"="#009E73","Platelets"="#D55E00","Other"="#999999")

theme_thesis <- function(base_size = 12) {
  theme_minimal(base_size = base_size) +
    theme(
      plot.title   = element_blank(),
      plot.subtitle= element_blank(),
      axis.title   = element_text(face = "bold", size = base_size+0),
      axis.text    = element_text(size = base_size),
      axis.text.x  = element_text(angle = 45, hjust = 1),
      panel.grid.minor = element_blank(),
      legend.position = "bottom",
      legend.title    = element_blank(),
      legend.text     = element_text(size = base_size+1),
      legend.key.width  = unit(1.0, "lines"),
      legend.key.height = unit(1.0, "lines"),
      legend.box.margin = margin(6, 20, 6, 20),
      plot.margin       = margin(2, 18, 2, 12)
    )
}
legend_compact <- theme(
  legend.key.width  = unit(0.9, "lines"),
  legend.key.height = unit(0.9, "lines")
)

shorten_label <- function(s, max_chars = 28) {
  s <- as.character(s)
  ifelse(nchar(s) > max_chars, paste0(substr(s,1,max_chars-1), "…"), s)
}

label_date_en <- function(fmt = "%b %Y") {
  function(d) {
    old <- try(Sys.getlocale("LC_TIME"), silent = TRUE)
    on.exit(try(Sys.setlocale("LC_TIME", old), silent = TRUE))
    ok <- try(Sys.setlocale("LC_TIME", "C"), silent = TRUE)
    if (inherits(ok, "try-error") || is.na(ok)) try(Sys.setlocale("LC_TIME", "English"), silent = TRUE)
    format(d, fmt)
  }
}

# -------------------------------------------------------------------
# MACRO mapping per visualizzazioni
# -------------------------------------------------------------------
is_rbc <- function(x){
  x0 <- tolower(x)
  (str_detect(x0, "emaz|eritro|\\brbc\\b|\\b25\\d{2}\\b|leucodeplet")) &
    !str_detect(x0, "plasm|\\bplt\\b|piastr|buffy")
}
is_plasma <- function(x){
  x0 <- tolower(x)
  str_detect(x0, "plasm|\\bffp\\b|\\bpfc\\b") & !str_detect(x0, "\\bplt\\b|piastr|buffy|\\b4305\\b")
}
is_platelet <- function(x){
  x0 <- tolower(x)
  str_detect(x0, "\\b4305\\b|\\bplt\\b|piastr|platelet|buffy|\\b2004\\b|\\b1904\\b")
}
macro_from_component <- function(x){
  case_when(
    is_rbc(x)      ~ "RBCs",
    is_plasma(x)   ~ "Plasma",
    is_platelet(x) ~ "Platelets",
    TRUE           ~ "Other"
  )
}

df_vis <- df %>%
  mutate(
    macro = macro_from_component(emocomponente),
    emocomponente_std = emocomponente %>% toupper() %>% str_squish(),
    emc_short = shorten_label(emocomponente_std)
  )

# -------------------------------------------------------------------
# Utility: ordinamento per facet SENZA dipendenze esterne
# -------------------------------------------------------------------
reorder_within <- function(x, by, within, fun = mean, sep = "___", ...) {
  x <- paste(x, within, sep = sep)
  stats::reorder(x, by, FUN = fun)
}
scale_x_reordered <- function(..., sep = "___") {
  ggplot2::scale_x_discrete(labels = function(x) gsub(paste0(sep, ".*$"), "", x), ...)
}

# -------------------------------------------------------------------
# G1 — Monthly wasted units by macro (stacked)
# -------------------------------------------------------------------
g1_df <- df_vis %>%
  filter(!is.na(month)) %>%
  count(month, macro, name = "wasted_units")

if (nrow(g1_df)) {
  min_m <- min(g1_df$month, na.rm = TRUE)
  max_m <- max(g1_df$month, na.rm = TRUE)
  quarter_breaks <- seq(floor_date(min_m, "quarter"), ceiling_date(max_m, "quarter"), by = "3 months")
  
  g1 <- ggplot(g1_df, aes(x = month, y = wasted_units, fill = macro)) +
    geom_col() +
    scale_fill_manual(values = pal_macro) +
    scale_y_continuous(labels = comma) +
    scale_x_date(breaks = quarter_breaks, labels = label_date_en("%b %Y"),
                 expand = expansion(mult = c(0.005, 0.02))) +
    labs(x = "Month", y = "Units") +
    theme_thesis() + legend_compact +
    coord_cartesian(clip = "off")
  
  ggsave(file.path(outputs_dir, "g1_waste_by_macro_month.png"),
         g1, width = content_width_in, height = 5.8, dpi = target_ppi, units = "in")
}

# -------------------------------------------------------------------
# G2 — Wastage reasons: share by macro (100% stacked)
# -------------------------------------------------------------------
g2_df <- df_vis %>%
  mutate(tipo_consumo = ifelse(is.na(tipo_consumo) | tipo_consumo == "", "Unknown", tipo_consumo)) %>%
  count(macro, tipo_consumo, name = "n") %>%
  group_by(macro) %>%
  mutate(perc = n / sum(n)) %>%
  ungroup()

if (nrow(g2_df)) {
  g2 <- ggplot(g2_df, aes(x = macro, y = perc, fill = tipo_consumo)) +
    geom_col(position = "fill") +
    scale_y_continuous(labels = percent_format(accuracy = 1)) +
    scale_fill_brewer(palette = "Set2", na.translate = FALSE) +
    labs(x = "Macro-component", y = "Share") +
    theme_thesis() +
    guides(fill = guide_legend(ncol = 2, byrow = TRUE))
  
  ggsave(file.path(outputs_dir, "g2_reasons_share_by_macro.png"),
         g2, width = content_width_in, height = 5.8, dpi = target_ppi, units = "in")
}

# -------------------------------------------------------------------
# G3 — Top 5 discarded components per macro (faceted, horizontal bars) — LOG SCALE
# -------------------------------------------------------------------
g3_df <- df_vis %>%
  filter(macro %in% c("RBCs","Plasma","Platelets")) %>%
  mutate(
    emc_short = ifelse(is.na(emc_short) | emc_short == "", "Other/Unknown", emc_short)
  ) %>%
  count(macro, emc_short, name = "n") %>%
  group_by(macro) %>%
  arrange(desc(n), .by_group = TRUE) %>%
  slice_head(n = 5) %>%
  ungroup() %>%
  filter(n > 0)   # evita problemi col log(0)

if (nrow(g3_df)) {
  # Ordine entro ciascun facet senza dipendenze esterne
  g3_df <- g3_df %>%
    group_by(macro) %>%
    arrange(n, .by_group = TRUE) %>%
    mutate(emc_key = paste(macro, emc_short, sep = "___"),
           ord_key = factor(emc_key, levels = emc_key)) %>%
    ungroup()
  
  g3 <- ggplot(g3_df,
               aes(x = ord_key, y = n, fill = macro)) +
    geom_col(show.legend = FALSE) +
    scale_fill_manual(values = pal_macro) +
    # ------- LOG SCALE SOLO QUI -------
  scale_y_log10(labels = comma, breaks = scales::breaks_log(n = 6, base = 10)) +
    coord_flip() +
    facet_wrap(~ macro, ncol = 1, scales = "free_y") +
    labs(x = "Component (shortened)", y = "Units (log scale)") +
    theme_thesis() +
    theme(
      strip.text   = element_text(face = "bold", size = 12),
      axis.text.y  = element_text(size = 11),
      plot.margin  = margin(6, 26, 6, 10)
    ) +
    scale_x_discrete(labels = function(x) sub("^.*___", "", x))
  
  ggsave(file.path(outputs_dir, "g3_top_components_by_macro.png"),
         g3, width = content_width_in, height = 9.0, dpi = target_ppi, units = "in")
}

# -------------------------------------------------------------------
# 3 TABELLE PNG (stile thesis, con CSV già salvati dove presenti)
# -------------------------------------------------------------------
.content_width_px <- round(content_width_in * target_ppi)
.table_pct_width  <- 80

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

thesis_table_png_df <- function(df, title, out_png) {
  num_cols <- names(df)[-1]
  df_fmt <- df
  for (cc in num_cols) {
    if (!is.numeric(df_fmt[[cc]])) {
      as_num <- suppressWarnings(as.numeric(df_fmt[[cc]]))
      if (!all(is.na(as_num))) df_fmt[[cc]] <- as_num
    }
  }
  gt_tbl <- df_fmt %>%
    gt::gt() %>%
    gt::tab_header(title = title) %>%
    gt::fmt_number(columns = num_cols, decimals = 0, use_seps = TRUE) %>%
    gt_style_thesis()
  dir.create(dirname(out_png), recursive = TRUE, showWarnings = FALSE)
  gt::gtsave(gt_tbl, filename = out_png, vwidth = .content_width_px, expand = 0)
}

# T1 — Year × Macro (counts)
t1_wide <- df_vis %>%
  filter(!is.na(year)) %>%
  mutate(macro = ifelse(macro %in% c("RBCs","Plasma","Platelets"), macro, "Other")) %>%
  count(year, macro, name = "units") %>%
  tidyr::pivot_wider(names_from = macro, values_from = units, values_fill = 0) %>%
  arrange(year)
readr::write_csv(t1_wide, file.path(outputs_dir, "t1_year_by_macro.csv"))
thesis_table_png_df(t1_wide, "Annual wasted units by macro", file.path(outputs_dir, "t1_year_by_macro.png"))

# T2 — Year × ABO/Rh group (counts)
t2_wide <- df_vis %>%
  filter(!is.na(year)) %>%
  mutate(gruppo_ab0_rh = ifelse(is.na(gruppo_ab0_rh) | gruppo_ab0_rh == "", "NA", gruppo_ab0_rh)) %>%
  count(year, gruppo_ab0_rh, name = "units") %>%
  tidyr::pivot_wider(names_from = gruppo_ab0_rh, values_from = units, values_fill = 0) %>%
  arrange(year)
readr::write_csv(t2_wide, file.path(outputs_dir, "t2_year_by_group.csv"))
thesis_table_png_df(t2_wide, "Annual wasted units by ABO/Rh group", file.path(outputs_dir, "t2_year_by_group.png"))

# T3 — Top 15 SIMT by waste volume (overall)
t3_tbl <- df_vis %>%
  mutate(simt = ifelse(is.na(simt) | simt == "", "NA", simt)) %>%
  count(simt, name = "units") %>%
  arrange(desc(units)) %>%
  slice_head(n = 15) %>%
  rename(SIMT = simt, `Wasted units` = units)
readr::write_csv(t3_tbl, file.path(outputs_dir, "t3_top_simt.csv"))
thesis_table_png_df(t3_tbl, "Top 15 SIMT by wasted units (overall)", file.path(outputs_dir, "t3_top_simt.png"))

message("SCARTATO OK. CSV canonici in ", normalized_dir, " — Grafici/Tabelle in ", outputs_dir)
