# =============================================================================
# Script name : 02_centri_raccolta_zona.R
# Purpose     : Analyze RACCOLTA by centers aggregated into zones (ZONA).
#               For each zone, produce:
#                 - PNG: Collection time series (1–2 panels, one per center)
#                 - PNG: Weekday share (1–2 panels, one per center)
#                 - CSV: Monthly units by family & center
#                 - CSV: Weekday share by center
# Data field  : `punto_prelievo` (from raccolta)
# Families    : Whole blood / RBC / Plasma / Platelets
# Weekday     : Share of donations (n_distinct(cdm)) Mon→Sun
# =============================================================================

suppressPackageStartupMessages({
  pkgs <- c("tidyverse","lubridate","janitor","scales","fs","glue","gridExtra","grid","stringr")
  for(p in pkgs) if(!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(tidyverse); library(lubridate); library(janitor); library(scales)
  library(fs); library(glue); library(gridExtra); library(grid); library(stringr)
})

# -----------------------------------------------------------------------------
# A4/Word sizing — PNGs at text width (A4 portrait, side margins 3 pt)
# -----------------------------------------------------------------------------
page_width_in   <- 8.27      # A4 width (in)
left_margin_pt  <- 3         # Word side margins (pt)
right_margin_pt <- 3
target_ppi      <- 96        # ensures 1:1 size when pasted in Word

content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72

# Panel heights tuned for readability at content width
ts_panel_height_in <- 4.6     # per-center time series panel height
wd_panel_height_in <- 4.2     # per-center weekday panel height

# -----------------------------------------------------------------------------
# PATHS
# -----------------------------------------------------------------------------
dir_raccolta  <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/raccolta/normalized"
file_raccolta <- file.path(dir_raccolta, "raccolta_merged_clean_with_group.csv")

out_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi complessive/Offerta_Centri"
dir_plot <- file.path(out_root, "grafici"); dir_csv <- file.path(out_root, "file csv")
fs::dir_create(c(dir_plot, dir_csv), recurse = TRUE)

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
log_msg <- function(...) cat(format(Sys.time(),"[%Y-%m-%d %H:%M:%S]"), ..., "\n")

safe_read <- function(path, guess_max = 300000) {
  stopifnot(file.exists(path))
  readr::read_csv(path, guess_max = guess_max, show_col_types = FALSE) %>% clean_names()
}

ensure_date <- function(x) {
  if(inherits(x,"Date")) return(x)
  if(inherits(x,"POSIXct")||inherits(x,"POSIXt")) return(as.Date(x))
  suppressWarnings(as.Date(x))
}

# Normalize strings (uppercase, no diacritics/punctuation)
strip_accents <- function(x) {
  tryCatch(iconv(x, from = "", to = "ASCII//TRANSLIT"), error = function(e) x)
}
norm_str <- function(x) {
  x %>%
    as.character() %>%
    strip_accents() %>%
    toupper() %>%
    str_replace_all("[^A-Z0-9 ]", " ") %>%
    str_squish()
}

# Component family recognizers (RACCOLTA only)
is_whole_blood <- function(x) {
  x0 <- tolower(x); str_detect(x0, "sangue\\s*intero")
}
is_rbc <- function(x) {
  x0 <- tolower(x)
  str_detect(x0, "emaz|eritro|\\brbc\\b|leucodeplet|\\b250(4)?\\b") &
    !is_whole_blood(x) & !str_detect(x0, "plasm|piastr|\\bplt\\b|buffy")
}
is_plasma <- function(x) {
  x0 <- tolower(x); str_detect(x0, "plasm") & !str_detect(x0, "piastr|\\bplt\\b|buffy")
}
is_platelet <- function(x) {
  x0 <- tolower(x); str_detect(x0, "piastr|\\bplt\\b|platelet|buffy|\\b4305\\b")
}
component_family <- function(x) {
  case_when(
    is_whole_blood(x) ~ "Whole blood",
    is_rbc(x)         ~ "RBC",
    is_plasma(x)      ~ "Plasma",
    is_platelet(x)    ~ "Platelets",
    TRUE              ~ "Other"
  )
}

# Map `punto_prelievo` to canonical center keys
std_center <- function(punto_prelievo) {
  s <- norm_str(punto_prelievo)
  case_when(
    str_detect(s, "\\bVALDAGNO\\b")                         ~ "VALDAGNO",
    str_detect(s, "\\bMONTECCHIO(\\s+MAGGIORE)?\\b")        ~ "MONTECCHIO",
    str_detect(s, "\\bNOVENTA(\\s+VICENTINA)?\\b")          ~ "NOVENTA",
    str_detect(s, "\\bLONIGO\\b")                           ~ "LONIGO",
    str_detect(s, "\\bVICENZA\\b|\\bS\\.? BORTOLO\\b")      ~ "VICENZA",
    str_detect(s, "\\bSANTORSO\\b")                         ~ "SANTORSO",
    str_detect(s, "\\bTHIENE\\b")                           ~ "THIENE",
    str_detect(s, "\\bSCHIO\\b")                            ~ "SCHIO",
    str_detect(s, "\\bSANDRIGO\\b")                         ~ "SANDRIGO",
    str_detect(s, "\\bMAROSTICA\\b")                        ~ "MAROSTICA",
    str_detect(s, "\\bBASSANO(\\s+DEL\\s+GRAPPA)?\\b")      ~ "BASSANO",
    TRUE ~ NA_character_
  )
}

# Palette and thesis theme
fam_levels <- c("Whole blood","RBC","Plasma","Platelets","Other")
fam_pal <- c("Whole blood"="#8e44ad","RBC"="#e41a1c","Plasma"="#377eb8","Platelets"="#4daf4a","Other"="#999999")

theme_thesis <- theme_minimal(base_family = "serif", base_size = 12) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
    axis.title = element_text(face = "bold"),
    panel.grid.minor = element_blank(),
    legend.title = element_text(face = "bold"),
    plot.margin = margin(6, 10, 6, 10)
  )

theme_legend_compact <- theme(
  legend.position = "bottom",
  legend.text = element_text(size = 8),
  legend.key.width = unit(0.8,"lines"),
  legend.key.height = unit(0.8,"lines"),
  legend.box.margin = margin(2,2,2,2)
)

# -----------------------------------------------------------------------------
# LOAD RACCOLTA
# -----------------------------------------------------------------------------
log_msg("Reading raccolta: ", file_raccolta)
rac <- safe_read(file_raccolta) %>%
  mutate(
    data_prestazione_date = ensure_date(data_prestazione_date),
    month = floor_date(data_prestazione_date, "month"),
    punto_prelievo_std = std_center(punto_prelievo),
    family = component_family(emocomponente),
    numero_emocomponenti = suppressWarnings(as.numeric(numero_emocomponenti)),
    cdm = as.character(cdm)
  ) %>%
  filter(toupper(stato_donazione) == "CONCLUSA") %>%
  filter(!is.na(month), !is.na(punto_prelievo_std))

# Debug: centers found
centri_map <- rac %>% count(punto_prelievo_std, sort = TRUE)
log_msg("Mapped centers: ", paste(centri_map$punto_prelievo_std, collapse = ", "))

# -----------------------------------------------------------------------------
# ZONES
# -----------------------------------------------------------------------------
zone <- list(
  Valdagno_Montecchio = c("VALDAGNO","MONTECCHIO"),
  Noventa_Lonigo     = c("NOVENTA","LONIGO"),
  Vicenza            = c("VICENZA"),
  Santorso           = c("SANTORSO"),
  Thiene_Schio       = c("THIENE","SCHIO"),
  Sandrigo_Marostica = c("SANDRIGO","MAROSTICA"),
  Bassano            = c("BASSANO")
)

# -----------------------------------------------------------------------------
# PLOT BUILDERS (thesis style; date axis with quarterly ticks)
# -----------------------------------------------------------------------------
plot_timeseries_center <- function(df_center, center_name) {
  fams_present <- df_center %>% distinct(family) %>% pull(family)
  pal_use <- fam_pal[names(fam_pal) %in% fams_present]
  ggplot(df_center, aes(x = month, y = units, colour = family)) +
    geom_line(linewidth = 1, na.rm = TRUE) +
    scale_colour_manual(values = pal_use, name = NULL) +
    scale_y_continuous(labels = comma) +
    scale_x_date(date_breaks = "3 months", date_labels = "%b %Y") +
    labs(title = glue("{center_name} — Collection time series"),
         x = "Month", y = "Units collected") +
    theme_thesis + theme_legend_compact +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
}

plot_weekday_share_center <- function(df_center_days, center_name) {
  ymax <- max(100, max(df_center_days$share_pct, na.rm = TRUE) + 5)
  ggplot(df_center_days, aes(x = weekday, y = share_pct)) +
    geom_col() +
    geom_text(aes(label = paste0(round(share_pct,1),"%")), vjust = -0.35, size = 3) +
    scale_y_continuous(labels = function(x) paste0(x, "%"), limits = c(0, ymax), expand = expansion(mult = c(0.02, 0.08))) +
    labs(title = glue("{center_name} — Weekday share of donations"),
         x = "Weekday", y = "Share of donations (%)") +
    theme_thesis
}

# Utilities for weekday labels
weekday_order <- c("Mon","Tue","Wed","Thu","Fri","Sat","Sun")
weekday_label <- function(d) {
  i <- lubridate::wday(d, week_start = 1) # 1 = Monday
  labs <- weekday_order
  labs[i]
}

# -----------------------------------------------------------------------------
# ELABORATION BY ZONE
# -----------------------------------------------------------------------------
for (zn in names(zone)) {
  centers <- zone[[zn]]
  log_msg("== Zone:", zn, "==")
  log_msg("Expected centers:", paste(centers, collapse = ", "))
  
  df_zone <- rac %>% filter(punto_prelievo_std %in% centers)
  log_msg(glue("Zone rows: {nrow(df_zone)}; Centers present: {paste(sort(unique(df_zone$punto_prelievo_std)), collapse = ', ')}"))
  if (nrow(df_zone) == 0) {
    log_msg(glue("[{zn}] No data: skipping outputs."))
    next
  }
  
  # ---- Monthly units by family & center (CSV) ----
  monthly_units <- df_zone %>%
    mutate(family = factor(family, levels = fam_levels)) %>%
    group_by(punto_prelievo_std, family, month) %>%
    summarise(units = sum(replace_na(numero_emocomponenti, 0), na.rm = TRUE), .groups = "drop") %>%
    arrange(punto_prelievo_std, family, month)
  
  csv_monthly_path <- file.path(dir_csv, glue("{zn}_monthly_units_by_family_center.csv"))
  readr::write_csv(monthly_units, csv_monthly_path)
  
  # ---- Time series PNG (1–2 panels, one per center) ----
  plots_ts <- list()
  for (ct in centers) {
    dct <- monthly_units %>% filter(punto_prelievo_std == ct)
    if (nrow(dct) == 0) next
    plots_ts[[ct]] <- plot_timeseries_center(dct, ct)
  }
  if (length(plots_ts) > 0) {
    g_ts <- do.call(gridExtra::arrangeGrob, c(plots_ts, ncol = 1))
    height_in <- ts_panel_height_in * length(plots_ts)
    ggsave(file.path(dir_plot, glue("{zn}_timeseries.png")), g_ts,
           width = content_width_in, height = height_in, units = "in", dpi = target_ppi)
    log_msg(glue("[{zn}] Saved time series PNG."))
  } else {
    log_msg(glue("[{zn}] No TS panels to save."))
  }
  
  # ---- Weekday share by center (CSV) ----
  wkd <- df_zone %>%
    filter(!is.na(cdm), !is.na(data_prestazione_date)) %>%
    mutate(weekday = weekday_label(data_prestazione_date)) %>%
    filter(!is.na(weekday))
  
  weekday_share <- wkd %>%
    group_by(punto_prelievo_std, weekday) %>%
    summarise(don = n_distinct(cdm), .groups = "drop") %>%
    group_by(punto_prelievo_std) %>%
    complete(weekday = weekday_order, fill = list(don = 0)) %>%
    mutate(total = sum(don), share_pct = ifelse(total > 0, 100*don/total, 0)) %>%
    ungroup() %>%
    mutate(weekday = factor(weekday, levels = weekday_order))
  
  csv_weekday_path <- file.path(dir_csv, glue("{zn}_weekday_share_by_center.csv"))
  readr::write_csv(weekday_share, csv_weekday_path)
  
  # ---- Weekday share PNG (1–2 panels, one per center) ----
  plots_wd <- list()
  for (ct in centers) {
    dct <- weekday_share %>% filter(punto_prelievo_std == ct)
    if (nrow(dct) == 0) next
    plots_wd[[ct]] <- plot_weekday_share_center(dct, ct)
  }
  if (length(plots_wd) > 0) {
    g_wd <- do.call(gridExtra::arrangeGrob, c(plots_wd, ncol = 1))
    height_in <- wd_panel_height_in * length(plots_wd)
    ggsave(file.path(dir_plot, glue("{zn}_weekday_share.png")), g_wd,
           width = content_width_in, height = height_in, units = "in", dpi = target_ppi)
    log_msg(glue("[{zn}] Saved weekday share PNG."))
  } else {
    log_msg(glue("[{zn}] No weekday panels to save."))
  }
}

log_msg("DONE. Outputs in: ", out_root)
