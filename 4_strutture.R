# =============================================================================
# Script name : Macrocomponents_by_Structure.R
# Purpose     : For each aggregated structure (from 'Reparto'), produce:
#               - weekday_distribution.png
#               - urgency_trend.png
#               - top5_by_component.png (facets: RBC / Plasma / Platelets)
#               - rows_used.csv (rows used in the figures)
#               - top5_by_component.csv (top-5 departments per macro)
# Horizon     : 2023-01-01 .. 2025-05-31
# Note        : Migliorie visive: etichette asse X in obliquo dove utili e
#               grafici sviluppati maggiormente in altezza per leggibilità.
#               Nessun cambiamento nella logica dei dati, tranne l'esclusione
#               delle "case di riposo" per la struttura MONTECCHIO.
# =============================================================================

suppressPackageStartupMessages({
  library(tidyverse)
  library(readr)
  library(lubridate)
  library(glue)
  library(forcats)
  library(scales)
  library(grid)
})

log_msg <- function(...) cat(format(Sys.time(), "[%Y-%m-%d %H:%M:%S]"), paste(..., collapse=" "), "\n")
`%||%` <- function(a,b) if (is.null(a)) b else a

# -----------------------------------------------------------------------------
# PATHS
# -----------------------------------------------------------------------------
in_dir     <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized"
file_ulss7 <- file.path(in_dir, "ULSS7.csv")
file_ulss8 <- file.path(in_dir, "ULSS8.csv")

out_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/analisi complessive/Macrocomponents"
dir.create(out_root, recursive = TRUE, showWarnings = FALSE)

# -----------------------------------------------------------------------------
# WORD/A4 SIZING — PNGs sized to Word text width (A4, side margins 3 pt)
# -----------------------------------------------------------------------------
page_width_in   <- 8.27   # A4 portrait width in inches
left_margin_pt  <- 3
right_margin_pt <- 3
target_ppi      <- 96

content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72

# Heights (inches): più alte per leggibilità degli assi X
h_weekday_in <- 5.0
h_trend_in   <- 5.8
top5_width_in  <- 13
top5_height_in <- 7.8

# -----------------------------------------------------------------------------
# READ
# -----------------------------------------------------------------------------
stopifnot(file.exists(file_ulss7), file.exists(file_ulss8))
log_msg("Reading:", file_ulss7, "and", file_ulss8)
ulss7 <- read_csv(file_ulss7, show_col_types = FALSE)
ulss8 <- read_csv(file_ulss8, show_col_types = FALSE)
data  <- bind_rows(ulss7, ulss8)

# -----------------------------------------------------------------------------
# PREP
# -----------------------------------------------------------------------------
if (!inherits(data[["Data Trasfusione"]], "Date")) {
  data <- data %>% mutate(`Data Trasfusione` = as.Date(`Data Trasfusione`))
}
data <- data %>%
  mutate(Date = `Data Trasfusione`) %>%
  filter(!is.na(Date),
         Date >= as.Date("2023-01-01"),
         Date <= as.Date("2025-05-31"))

# Macro classification from 'EMC Emolife'
is_rbc      <- function(x){ x0<-tolower(x); str_detect(x0,"emaz|eritro|\\brbc\\b|leucodeplet|\\b2504?\\b") & !str_detect(x0,"plasm|\\bplt\\b|piastr|buffy") }
is_plasma   <- function(x){ x0<-tolower(x); str_detect(x0,"plasm") & !str_detect(x0,"\\bplt\\b|piastr|buffy|4305") }
is_platelet <- function(x){ x0<-tolower(x); str_detect(x0,"\\b4305\\b|\\bplt\\b|piastr|platelet|buffy") }

data <- data %>%
  mutate(
    macro = case_when(
      is_rbc(`EMC Emolife`)      ~ "RBC",
      is_plasma(`EMC Emolife`)   ~ "Plasma",
      is_platelet(`EMC Emolife`) ~ "Platelets",
      TRUE ~ NA_character_
    ),
    macro = factor(macro, levels = c("RBC","Plasma","Platelets"))
  )

# Urgency to English labels
data <- data %>%
  mutate(
    Urgency = case_when(
      str_detect(tolower(Urgenza %||% ""), "programm")     ~ "Scheduled",
      str_detect(tolower(Urgenza %||% ""), "urgentissima") ~ "Very urgent",
      str_detect(tolower(Urgenza %||% ""), "urgent")       ~ "Urgent",
      TRUE ~ "Other"
    ),
    Urgency = factor(Urgency, levels = c("Scheduled","Urgent","Very urgent","Other"))
  )

# Weekday (Mon..Sun, EN) and month for trend
weekday_levels <- c("Mon","Tue","Wed","Thu","Fri","Sat","Sun")
data <- data %>%
  mutate(
    Weekday = factor(weekday_levels[wday(Date, week_start = 1)], levels = weekday_levels),
    Month   = floor_date(Date, "month")
  )

# -----------------------------------------------------------------------------
# STRUCTURE GROUPING (from 'Reparto' only)
# -----------------------------------------------------------------------------
final_structures <- c(
  "Vicenza","Santorso","Bassano","Asiago","Noventa Vicentina",
  "Arzignano","Valdagno","Montecchio","Lonigo","Altro"
)

classify_structure_by_reparto <- function(reparto){
  r <- tolower(as.character(reparto))
  case_when(
    str_detect(r, "vicenza|\\bvic\\.|\\bvic\\b|san\\s*bortolo|sanbortolo|s\\.\\s*bortolo|\\bbortolo\\b") ~ "Vicenza",
    str_detect(r, "santorso|alto\\s*vicentino|altovicentino|schio|thiene|\\bdsa\\b|\\bsant\\b|\\bsant\\.|\\bsant\\s|\\bsoap\\b|\\bcorsia\\b") ~ "Santorso",
    str_detect(r, "\\bbassano\\b") ~ "Bassano",
    str_detect(r, "\\basiago\\b")  ~ "Asiago",
    str_detect(r, "noventa(\\s|\\.|\\-|_)*vicent") | str_detect(r, "\\bnoventa\\b") ~ "Noventa Vicentina",
    str_detect(r, "arzignano|\\barz\\b|arz\\.") ~ "Arzignano",
    str_detect(r, "valdagno|san\\s*lorenzo|\\bvald\\.|\\bvald\\b") ~ "Valdagno",
    str_detect(r, "montecchio") ~ "Montecchio",
    str_detect(r, "lonigo") ~ "Lonigo",
    TRUE ~ "Altro"
  )
}

data <- data %>%
  mutate(Structure = classify_structure_by_reparto(Reparto)) %>%
  mutate(Structure = if_else(Structure %in% final_structures, Structure, "Altro"))

# -----------------------------------------------------------------------------
# *** PUNCTUAL CHANGE REQUESTED ***
# Exclude "casa di riposo" ONLY for Structure == "Montecchio"
# (matches 'casa di riposo' in either Reparto or Ospedale; case/spacing insensitive)
# -----------------------------------------------------------------------------
is_casa_di_riposo <- function(x) {
  str_detect(tolower(x %||% ""), "casa\\s*di\\s*riposo")
}

before_n <- nrow(data)
data <- data %>%
  filter(!(Structure == "Montecchio" & (is_casa_di_riposo(Reparto) | is_casa_di_riposo(Ospedale))))
after_n <- nrow(data)
log_msg("Filtered Casa di Riposo in Montecchio:", before_n - after_n, "rows removed")

present_structures <- data %>% distinct(Structure) %>% pull(Structure)
log_msg("Structures present:", paste(present_structures, collapse=", "))

# -----------------------------------------------------------------------------
# THEMES & PALETTES (coerenti con gli altri script)
# -----------------------------------------------------------------------------
theme_thesis <- theme_minimal(base_family = "serif", base_size = 13) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 18, margin = margin(b = 4)),
    axis.title = element_text(face = "bold", size = 13),
    axis.text  = element_text(size = 12),
    panel.grid.minor = element_blank(),
    legend.title = element_blank(),
    legend.position = "bottom",
    legend.text = element_text(size = 11),
    legend.key.width  = unit(0.9, "lines"),
    legend.key.height = unit(0.9, "lines"),
    legend.box.margin = margin(2, 2, 2, 2),
    plot.margin = margin(t = 6, r = 14, b = 6, l = 14)
  )

urg_cols <- c(
  "Scheduled"   = "#6BAED6",
  "Urgent"      = "#FD8D3C",
  "Very urgent" = "#EF3B2C",
  "Other"       = "#BDBDBD"
)

macro_cols <- c("RBC" = "#0072B2", "Plasma" = "#009E73", "Platelets" = "#D55E00")

all_months <- tibble(Month = seq.Date(as.Date("2023-01-01"), as.Date("2025-05-01"), by = "month"))

save_plot_or_placeholder <- function(pdata, plot_fn, path, width_in, height_in, dpi_in = target_ppi, title_when_empty="No data available"){
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  if (nrow(pdata) > 0) {
    ggsave(
      filename = path,
      plot     = plot_fn(),
      width    = width_in,
      height   = height_in,
      units    = "in",
      dpi      = dpi_in,
      bg       = "white"
    )
  } else {
    p_empty <- ggplot() + annotate("text", x=0, y=0, label=title_when_empty, size=5) + theme_void()
    ggsave(filename = path, plot = p_empty, width = width_in/2, height = height_in/2, units = "in", dpi = dpi_in, bg = "white")
  }
}

# -----------------------------------------------------------------------------
# PER-STRUCTURE LOOP: 3 PNG + CSV
# -----------------------------------------------------------------------------
for (hospital in present_structures) {
  hospital_dir <- file.path(out_root, hospital)
  dir.create(hospital_dir, recursive = TRUE, showWarnings = FALSE)
  
  dfh <- data %>% filter(Structure == hospital)
  
  # Rows used (for traceability)
  rows_csv <- dfh %>%
    select(Structure, Date, Month, Weekday, Urgency, macro, Ospedale, Reparto, `EMC Emolife`, everything())
  readr::write_csv(rows_csv, file.path(hospital_dir, "rows_used.csv"))
  
  # -------------------------
  # 1) Weekday distribution
  # -------------------------
  wk <- dfh %>%
    filter(!is.na(Urgency)) %>%
    count(Weekday, Urgency, name = "n") %>%
    mutate(Weekday = factor(Weekday, levels = weekday_levels)) %>%
    arrange(match(Weekday, weekday_levels))
  
  p_weekday_fn <- function(){
    ggplot(wk, aes(x = Weekday, y = n, fill = Urgency)) +
      geom_col(position = position_dodge(width = 0.8), width = 0.7, colour = "grey25", linewidth = 0.2) +
      scale_fill_manual(values = urg_cols, drop = FALSE) +
      scale_y_continuous(labels = comma, expand = expansion(mult = c(0.02, 0.06))) +
      labs(
        title = hospital,
        x = "Weekday", y = "Number of transfusions"
      ) +
      theme_thesis +
      theme(axis.text.x = element_text(angle = 35, hjust = 1, vjust = 1)) +
      guides(fill = guide_legend(nrow = 2, byrow = TRUE))
  }
  save_plot_or_placeholder(
    pdata   = wk,
    plot_fn = p_weekday_fn,
    path    = file.path(hospital_dir, "weekday_distribution.png"),
    width_in  = content_width_in,
    height_in = h_weekday_in
  )
  
  # -------------------------
  # 2) Urgency Trend (monthly)
  # -------------------------
  um <- dfh %>%
    filter(!is.na(Urgency)) %>%
    count(Month, Urgency, name = "n") %>%
    right_join(
      tidyr::expand(all_months, Month, Urgency = levels(data$Urgency)),
      by = c("Month","Urgency")
    ) %>%
    mutate(n = replace_na(n, 0L))
  
  p_trend_fn <- function(){
    ggplot(um, aes(x = Month, y = n, fill = Urgency)) +
      geom_col(colour = "grey25", linewidth = 0.2) +
      scale_fill_manual(values = urg_cols, drop = FALSE) +
      scale_y_continuous(labels = comma, expand = expansion(mult = c(0.02, 0.06))) +
      scale_x_date(date_breaks = "3 months", date_labels = "%b %Y",
                   limits = c(as.Date("2023-01-01"), as.Date("2025-05-31"))) +
      labs(
        title = hospital,
        x = "Month", y = "Number of transfusions"
      ) +
      theme_thesis +
      theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1)) +
      guides(fill = guide_legend(nrow = 2, byrow = TRUE))
  }
  save_plot_or_placeholder(
    pdata   = um,
    plot_fn = p_trend_fn,
    path    = file.path(hospital_dir, "urgency_trend.png"),
    width_in  = content_width_in,
    height_in = h_trend_in
  )
  
  # ------------------------------------------
  # 3) Top 5 departments by component (facet)
  # ------------------------------------------
  top_by_macro <- dfh %>%
    filter(!is.na(macro)) %>%
    group_by(macro, Reparto) %>%
    summarise(n = n(), .groups = "drop_last") %>%
    arrange(desc(n), .by_group = TRUE) %>%
    slice_head(n = 5) %>%
    ungroup() %>%
    mutate(
      Reparto_short = stringr::str_trunc(Reparto, width = 28, side = "right")
    ) %>%
    group_by(macro) %>%
    mutate(
      Reparto_short = forcats::fct_reorder(Reparto_short, n, .desc = FALSE)
    ) %>%
    ungroup() %>%
    mutate(macro = factor(macro, levels = c("RBC","Plasma","Platelets")))
  
  # CSV (top 5 per macro)
  readr::write_csv(
    top_by_macro %>% arrange(macro, desc(n)) %>% select(macro, Reparto, n),
    file.path(hospital_dir, "top5_by_component.csv")
  )
  
  # Padding per-facet per NON troncare le etichette numeriche a destra.
  pad_df <- top_by_macro %>%
    group_by(macro) %>%
    summarise(
      Reparto_short = dplyr::first(Reparto_short),
      pad_y = max(n, na.rm = TRUE) * if_else(macro == "RBC", 1.28, 1.18),
      .groups = "drop"
    )
  
  p_top_fn <- function(){
    ggplot(top_by_macro, aes(x = Reparto_short, y = n, fill = macro)) +
      geom_blank(data = pad_df, aes(x = Reparto_short, y = pad_y), inherit.aes = FALSE) +
      geom_col(width = 0.7, colour = "grey25", linewidth = 0.2) +
      geom_text(aes(label = scales::comma(n)),
                hjust = -0.10, size = 3.8, colour = "grey20") +
      coord_flip(clip = "off") +
      facet_wrap(~ macro, ncol = 3, scales = "free_y") +
      scale_fill_manual(values = macro_cols, guide = "none") +
      scale_y_continuous(labels = comma, expand = expansion(mult = c(0.02, 0.02))) +
      labs(
        title = hospital,
        x = "Department", y = "Number of transfusions"
      ) +
      theme_thesis +
      theme(
        axis.text.x      = element_text(angle = 35, hjust = 1, vjust = 1),
        strip.text       = element_text(face = "bold", size = 12),
        strip.background = element_rect(fill = "grey96", colour = NA),
        plot.margin      = margin(t = 6, r = 36, b = 10, l = 14)
      )
  }
  ggsave(
    filename = file.path(hospital_dir, "top5_by_component.png"),
    plot     = p_top_fn(),
    width    = top5_width_in, height = top5_height_in, units = "in", dpi = 150, bg = "white"
  )
}

log_msg("DONE. Outputs in:", out_root)
