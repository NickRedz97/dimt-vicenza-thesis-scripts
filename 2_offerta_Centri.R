# =============================================================================
# Script name : 02_centri_raccolta_zona.R  (per-centre version)
# Purpose     : Analyze RACCOLTA by individual centres (no grouping).
#               For each centre, produce:
#                 - PNG: Collection time series (single panel)
#                 - PNG: Weekday share (single panel)
#                 - CSV: Monthly units by family (for that centre)
#                 - CSV: Weekday share (for that centre)
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
page_width_in   <- 8.27
left_margin_pt  <- 3
right_margin_pt <- 3
target_ppi      <- 96

content_width_in   <- page_width_in - (left_margin_pt + right_margin_pt)/72
ts_panel_height_in <- 4.6
wd_panel_height_in <- 4.2

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
is_whole_blood <- function(x) { x0 <- tolower(x); str_detect(x0, "sangue\\s*intero") }
is_rbc <- function(x) {
  x0 <- tolower(x)
  str_detect(x0, "emaz|eritro|\\brbc\\b|leucodeplet|\\b250(4)?\\b") &
    !is_whole_blood(x) & !str_detect(x0, "plasm|piastr|\\bplt\\b|buffy")
}
is_plasma   <- function(x) { x0 <- tolower(x); str_detect(x0, "plasm") & !str_detect(x0, "piastr|\\bplt\\b|buffy") }
is_platelet <- function(x) { x0 <- tolower(x); str_detect(x0, "piastr|\\bplt\\b|platelet|buffy|\\b4305\\b") }
component_family <- function(x) {
  case_when(
    is_whole_blood(x) ~ "Whole blood",
    is_rbc(x)         ~ "RBC",
    is_plasma(x)      ~ "Plasma",
    is_platelet(x)    ~ "Platelets",
    TRUE              ~ "Other"
  )
}

# Map `punto_prelievo` to canonical centre keys
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
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    axis.title = element_text(face = "bold", size = 12),
    axis.text  = element_text(size = 11),
    panel.grid.minor = element_blank(),
    legend.title = element_text(face = "bold"),
    plot.margin = margin(6, 10, 6, 10)
  )

# Bigger legend & axes for time series / weekday
theme_legend_big <- theme(
  legend.position = "bottom",
  legend.text = element_text(size = 11),
  legend.key.width = unit(1.0,"lines"),
  legend.key.height = unit(1.0,"lines"),
  legend.box.margin = margin(4,6,4,6)
)
theme_axes_big <- theme(
  axis.text.x  = element_text(size = 11, angle = 45, hjust = 1),
  axis.text.y  = element_text(size = 11),
  axis.title.x = element_text(size = 12),
  axis.title.y = element_text(size = 12),
  plot.title   = element_text(size = 16, face = "bold", hjust = 0.5)
)

# English month labels for x-axis (independent from system locale)
label_date_en <- function(fmt = "%b %Y") {
  function(d) {
    old <- try(Sys.getlocale("LC_TIME"), silent = TRUE)
    on.exit(try(Sys.setlocale("LC_TIME", old), silent = TRUE))
    ok <- try(Sys.setlocale("LC_TIME", "C"), silent = TRUE)
    if (inherits(ok, "try-error") || is.na(ok)) try(Sys.setlocale("LC_TIME", "English"), silent = TRUE)
    format(d, fmt)
  }
}

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

# Debug: centres found
centri_map <- rac %>% count(punto_prelievo_std, sort = TRUE)
log_msg("Centres: ", paste(centri_map$punto_prelievo_std, collapse = ", "))

# -----------------------------------------------------------------------------
# PLOT BUILDERS (thesis style; date axis with quarterly ticks)
# -----------------------------------------------------------------------------
plot_timeseries_center <- function(df_center, center_name) {
  fams_present <- df_center %>% distinct(family) %>% pull(family)
  pal_use <- fam_pal[names(fam_pal) %in% fams_present]
  ggplot(df_center, aes(x = month, y = units, colour = family)) +
    geom_line(linewidth = 1.1, na.rm = TRUE) +
    scale_colour_manual(values = pal_use, name = NULL) +
    scale_y_continuous(labels = comma) +
    scale_x_date(date_breaks = "3 months", labels = label_date_en("%b %Y")) +
    labs(title = center_name, x = "Month", y = "Units collected") +
    theme_thesis + theme_legend_big + theme_axes_big
}

plot_weekday_share_center <- function(df_center_days, center_name) {
  ymax <- max(100, max(df_center_days$share_pct, na.rm = TRUE) + 5)
  ggplot(df_center_days, aes(x = weekday, y = share_pct)) +
    geom_col() +
    geom_text(aes(label = paste0(round(share_pct,1),"%")),
              vjust = -0.25, size = 5.2, fontface = "bold") +
    scale_y_continuous(labels = function(x) paste0(x, "%"),
                       limits = c(0, ymax),
                       expand = expansion(mult = c(0.01, 0.14))) +
    labs(title = center_name, x = "Weekday", y = "Share of donations (%)") +
    theme_thesis + theme_axes_big + theme(legend.position = "none")
}

# Utilities for weekday labels
weekday_order <- c("Mon","Tue","Wed","Thu","Fri","Sat","Sun")
weekday_label <- function(d) {
  i <- lubridate::wday(d, week_start = 1) # 1 = Monday
  weekday_order[i]
}

# -----------------------------------------------------------------------------
# ELABORATION BY CENTRE (no grouping)
# -----------------------------------------------------------------------------
centers_all <- rac %>% distinct(punto_prelievo_std) %>% arrange(punto_prelievo_std) %>% pull()

for (ct in centers_all) {
  log_msg("== Centre:", ct, "==")
  df_ct <- rac %>% filter(punto_prelievo_std == ct)
  if (nrow(df_ct) == 0) { log_msg("[", ct, "] No data."); next }
  
  # ---- Monthly units by family (CSV) ----
  monthly_units <- df_ct %>%
    mutate(family = factor(family, levels = fam_levels)) %>%
    group_by(punto_prelievo_std, family, month) %>%
    summarise(units = sum(replace_na(numero_emocomponenti, 0), na.rm = TRUE), .groups = "drop") %>%
    arrange(family, month)
  readr::write_csv(monthly_units, file.path(dir_csv, glue("{ct}_monthly_units_by_family_center.csv")))
  
  # ---- Time series PNG (single panel) ----
  p_ts <- plot_timeseries_center(monthly_units %>% select(-punto_prelievo_std), ct)
  ggsave(file.path(dir_plot, glue("{ct}_timeseries.png")), p_ts,
         width = content_width_in, height = ts_panel_height_in, units = "in", dpi = target_ppi)
  log_msg(glue("[{ct}] Saved time series PNG."))
  
  # ---- Weekday share (CSV) ----
  wkd <- df_ct %>%
    filter(!is.na(cdm), !is.na(data_prestazione_date)) %>%
    mutate(weekday = weekday_label(data_prestazione_date)) %>%
    filter(!is.na(weekday))
  weekday_share <- wkd %>%
    group_by(weekday) %>%
    summarise(don = n_distinct(cdm), .groups = "drop") %>%
    complete(weekday = weekday_order, fill = list(don = 0)) %>%
    mutate(total = sum(don), share_pct = ifelse(total > 0, 100*don/total, 0)) %>%
    ungroup() %>%
    mutate(weekday = factor(weekday, levels = weekday_order),
           punto_prelievo_std = ct) %>%
    select(punto_prelievo_std, everything())
  readr::write_csv(weekday_share, file.path(dir_csv, glue("{ct}_weekday_share_by_center.csv")))
  
  # ---- Weekday share PNG (single panel) ----
  p_wd <- plot_weekday_share_center(weekday_share %>% select(weekday, share_pct), ct)
  ggsave(file.path(dir_plot, glue("{ct}_weekday_share.png")), p_wd,
         width = content_width_in, height = wd_panel_height_in, units = "in", dpi = target_ppi)
  log_msg(glue("[{ct}] Saved weekday share PNG."))
}

log_msg("DONE. Outputs in: ", out_root)

