# =============================================================================
# Script name : Macrocomponents_by_Structure.R
# Purpose     : For each aggregated structure (from 'Reparto'), produce:
#               - weekday_distribution.png
#               - urgency_trend.png
#               - top5_by_component.png (facets: RBC / Plasma / Platelets)
#               - rows_used.csv (rows used in the figures)
#               - top5_by_component.csv (top-5 departments per macro)
# Horizon     : 2023-01-01 .. 2025-05-31
# =============================================================================

suppressPackageStartupMessages({
  library(tidyverse)
  library(readr)
  library(lubridate)
  library(glue)
  library(forcats)
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
left_margin_pt  <- 3      # Word side margins in points
right_margin_pt <- 3
target_ppi      <- 96     # 96 ppi → Word inserts at exact size

content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72

# Recommended heights (inches) for good balance at text width
h_weekday_in <- 4.4
h_trend_in   <- 5.0
# top5 will be exported with legacy size (13x6.5 in at 150 dpi)

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

present_structures <- data %>% distinct(Structure) %>% pull(Structure)
log_msg("Structures present:", paste(present_structures, collapse=", "))

# -----------------------------------------------------------------------------
# PLOT THEME (thesis style) + helpers
# -----------------------------------------------------------------------------
theme_thesis <- theme_minimal(base_family = "serif", base_size = 12) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
    axis.title = element_text(face = "bold"),
    panel.grid.minor = element_blank(),
    legend.title = element_text(face = "bold"),
    plot.margin = margin(t = 6, r = 6, b = 6, l = 6)
  )

urg_cols <- c(
  "Scheduled"   = "#6baed6",
  "Urgent"      = "#fd8d3c",
  "Very urgent" = "#ef3b2c",
  "Other"       = "#bdbdbd"
)

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
      dpi      = dpi_in
    )
  } else {
    p_empty <- ggplot() + annotate("text", x=0, y=0, label=title_when_empty, size=5) + theme_void()
    ggsave(filename = path, plot = p_empty, width = width_in/2, height = height_in/2, units = "in", dpi = dpi_in)
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
  
  # 1) Weekday distribution (by Urgency) — keep A4 text-width
  wk <- dfh %>%
    filter(!is.na(Urgency)) %>%
    count(Weekday, Urgency, name = "n") %>%
    mutate(Weekday = factor(Weekday, levels = weekday_levels)) %>%
    arrange(match(Weekday, weekday_levels))
  
  p_weekday_fn <- function(){
    ggplot(wk, aes(x = Weekday, y = n, fill = Urgency)) +
      geom_col(position = "dodge") +
      scale_fill_manual(values = urg_cols, drop = FALSE) +
      labs(title = glue("Weekday Distribution — {hospital}"),
           x = "Weekday", y = "Number of transfusions") +
      theme_thesis
  }
  save_plot_or_placeholder(
    pdata   = wk,
    plot_fn = p_weekday_fn,
    path    = file.path(hospital_dir, "weekday_distribution.png"),
    width_in  = content_width_in,
    height_in = h_weekday_in
  )
  
  # 2) Urgency Trend (monthly) — keep A4 text-width
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
      geom_col() +
      scale_fill_manual(values = urg_cols, drop = FALSE) +
      scale_x_date(date_breaks = "3 months", date_labels = "%b %Y",
                   limits = c(as.Date("2023-01-01"), as.Date("2025-05-31"))) +
      labs(title = glue("Urgency Trend — {hospital}"),
           x = "Month", y = "Number of transfusions") +
      theme_thesis +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
  }
  save_plot_or_placeholder(
    pdata   = um,
    plot_fn = p_trend_fn,
    path    = file.path(hospital_dir, "urgency_trend.png"),
    width_in  = content_width_in,
    height_in = h_trend_in
  )
  
  # 3) Top 5 departments by component — REVERT to legacy export size (13x6.5 in, 150 dpi)
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
      # With coord_flip(): largest at the top within each facet
      Reparto_short = forcats::fct_reorder(Reparto_short, n, .desc = FALSE)
    ) %>%
    ungroup() %>%
    mutate(macro = factor(macro, levels = c("RBC","Plasma","Platelets")))
  
  # CSV (top 5 per macro)
  readr::write_csv(
    top_by_macro %>% arrange(macro, desc(n)) %>% select(macro, Reparto, n),
    file.path(hospital_dir, "top5_by_component.csv")
  )
  
  p_top_fn <- function(){
    ggplot(top_by_macro, aes(x = Reparto_short, y = n)) +
      geom_col(width = 0.7) +
      coord_flip() +
      facet_wrap(~ macro, ncol = 3, scales = "free_y") +
      labs(title = glue("Top 5 Departments by Component — {hospital}"),
           x = "Department", y = "Number of transfusions") +
      theme_thesis +
      theme(strip.text = element_text(face = "bold"),
            legend.position = "none")
  }
  # Legacy sizing to avoid overlap/squeezing
  save_plot_or_placeholder(
    pdata   = top_by_macro,
    plot_fn = p_top_fn,
    path    = file.path(hospital_dir, "top5_by_component.png"),
    width_in  = 13,
    height_in = 6.5,
    dpi_in    = 150
  )
}

log_msg("DONE. Outputs in:", out_root)
