# ============================================================
# compute_LT_by_macro.R — Lead Time (donazione → trasfusione)
# Per-CDM, poi accorpamento per macro (RBCs, Plasma, Platelets)
#
# NOTE: Grafici resi più leggibili e coerenti con gli altri script.
# - Titolo dei grafici: SOLO l’emocomponente (RBCs / Plasma / Platelets), centrato.
# - Sottotitoli rimossi.
# - Nessun cambiamento alla logica o agli output (stessi file CSV/PNG).
# ============================================================

suppressPackageStartupMessages({
  pkgs <- c("readr","dplyr","lubridate","stringr","janitor",
            "purrr","tidyr","ggplot2","gt","glue","scales")
  miss <- pkgs[!(pkgs %in% rownames(installed.packages()))]
  if (length(miss)) install.packages(miss, dependencies = TRUE)
  lapply(pkgs, library, character.only = TRUE)
})

# ---------------------------
# Percorsi
# ---------------------------
base_root <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi"
rac_file  <- file.path(base_root, "raccolta", "normalized", "raccolta_merged_clean_with_group.csv")
ulss7_csv <- file.path(base_root, "trasfuso", "normalized", "ULSS7.csv")
ulss8_csv <- file.path(base_root, "trasfuso", "normalized", "ULSS8.csv")

lead_root <- file.path(base_root, "analisi complessive", "LeadTime")
dir.create(lead_root, recursive = TRUE, showWarnings = FALSE)

mkdir <- function(...) { p <- file.path(...); dir.create(p, recursive = TRUE, showWarnings = FALSE); p }

macro_levels <- c("RBCs","Plasma","Platelets")
macro_dirs <- setNames(file.path(lead_root, macro_levels), macro_levels)
invisible(lapply(macro_levels, function(m) {
  mkdir(lead_root, m, "csv"); mkdir(lead_root, m, "tables_png"); mkdir(lead_root, m, "plots")
}))

# ---------------------------
# Utility
# ---------------------------
is_rbc <- function(x){ x0 <- tolower(as.character(x)); grepl("emaz|eritro|\\brbc\\b|leucodeplet|\\b25\\d{2}\\b", x0) & !grepl("plasm|\\bplt\\b|piastr|buffy", x0) }
is_plasma <- function(x){ x0 <- tolower(as.character(x)); grepl("plasm|\\bffp\\b|\\bpfc\\b", x0) & !grepl("\\bplt\\b|piastr|buffy|\\b4305\\b", x0) }
is_platelet <- function(x){ x0 <- tolower(as.character(x)); grepl("\\b4305\\b|\\bplt\\b|piastr|platelet|buffy|\\b2004\\b|\\b1904\\b", x0) }
macro_from_component <- function(x){ dplyr::case_when(is_rbc(x) ~ "RBCs", is_plasma(x) ~ "Plasma", is_platelet(x) ~ "Platelets", TRUE ~ NA_character_) }

# ---- Tema & colori (coerenti con altri script) -----------------------------
theme_thesis <- theme_minimal(base_family = "serif", base_size = 13) +
  theme(
    plot.title   = element_text(face="bold", size=18, hjust=0.5, margin=margin(b=6)),
    axis.title   = element_text(face="bold", size=13),
    axis.text    = element_text(size=12),
    panel.grid.minor = element_blank(),
    legend.position="bottom",
    legend.title = element_blank(),
    plot.margin  = margin(6, 18, 6, 18)   # poco spazio sopra/sotto, respiro laterale
  )

theme_axes_big <- theme(
  axis.text.x  = element_text(size=12),
  axis.text.y  = element_text(size=12),
  axis.title.x = element_text(size=13),
  axis.title.y = element_text(size=13)
)

# Accenti per macro (Okabe–Ito)
macro_accent <- c(RBCs="#0072B2", Plasma="#009E73", Platelets="#D55E00")

fd_binwidth <- function(x){
  x <- x[is.finite(x)]
  if (!length(x)) return(1)
  bw <- 2*IQR(x)/(length(x)^(1/3))
  bw <- ifelse(is.na(bw) || bw <= 0, 1, bw)
  max(1, round(bw))
}

save_gt_png <- function(gt_tbl, out_png, vwidth = 1900){
  gt::gtsave(gt_tbl, file = out_png, vwidth = vwidth, expand = 6)
}

# ---------------------------
# Caricamento (CSV normalizzati)
# ---------------------------
stopifnot(file.exists(rac_file), file.exists(ulss7_csv), file.exists(ulss8_csv))

rac <- readr::read_csv(rac_file, show_col_types = FALSE) %>% janitor::clean_names()
t7  <- readr::read_csv(ulss7_csv, show_col_types = FALSE) %>% janitor::clean_names()
t8  <- readr::read_csv(ulss8_csv, show_col_types = FALSE) %>% janitor::clean_names()
tra <- bind_rows(t7 %>% mutate(ulss="ULSS7"), t8 %>% mutate(ulss="ULSS8"))

# Controlli stretti sui nomi
need_rac <- c("cdm","data_prestazione_date")
need_tra <- c("cdm","emc_emolife","data_trasfusione")
if (!all(need_rac %in% names(rac))) stop("RACCOLTA: mancano colonne: ", paste(setdiff(need_rac, names(rac)), collapse=", "))
if (!all(need_tra %in% names(tra))) stop("TRASFUSO: mancano colonne: ", paste(setdiff(need_tra, names(tra)), collapse=", "))

# ---------------------------
# Prepara TRASFUSO e RACCOLTA
# ---------------------------
tra <- tra %>%
  mutate(
    component_raw    = as.character(emc_emolife),
    macro            = macro_from_component(component_raw),
    transfusion_date = as.Date(data_trasfusione),
    ospedale_std     = dplyr::coalesce(as.character(ospedale), as.character(reparto))
  ) %>%
  select(cdm, macro, transfusion_date, ulss, ospedale = ospedale_std, any_of("reparto"))

rac_min <- rac %>%
  mutate(donation_date = as.Date(data_prestazione_date)) %>%
  filter(!is.na(cdm), !is.na(donation_date)) %>%
  group_by(cdm) %>% summarise(donation_date = min(donation_date), .groups = "drop")

# ---------------------------
# Join + LT per record
# ---------------------------
lt_raw <- tra %>%
  filter(!is.na(cdm)) %>%
  left_join(rac_min, by = "cdm") %>%
  mutate(LT_days = as.integer(difftime(transfusion_date, donation_date, units = "days")))

excluded_na  <- lt_raw %>% filter(is.na(transfusion_date) | is.na(donation_date) | is.na(macro)) %>% nrow()
excluded_neg <- lt_raw %>% filter(!is.na(LT_days) & LT_days < 0) %>% nrow()

lt <- lt_raw %>%
  filter(!is.na(transfusion_date), !is.na(donation_date), !is.na(macro), LT_days >= 0) %>%
  mutate(macro = factor(macro, levels = macro_levels))

# Salvataggi complessivi (root)
readr::write_csv(lt, file.path(lead_root, "lt_per_record_all.csv"))
readr::write_csv(tibble(metric=c("excluded_missing_or_na","excluded_negative_lt"),
                        value=c(excluded_na, excluded_neg)),
                 file.path(lead_root, "lt_exclusions_summary.csv"))

# ---------------------------
# Vista per-CDM e accorpamento per macro (unweighted by CDM)
# ---------------------------
lt_by_cdm <- lt %>%
  group_by(cdm, macro) %>%
  summarise(
    n_tx      = n(),
    mean_lt   = mean(LT_days),
    median_lt = median(LT_days),
    min_lt    = min(LT_days),
    max_lt    = max(LT_days),
    .groups = "drop"
  )

lt_by_cdm_summary <- lt_by_cdm %>%
  group_by(macro) %>%
  summarise(
    n_cdms      = n(),
    min_days    = min(mean_lt),
    mean_days   = mean(mean_lt),
    median_days = median(mean_lt),
    max_days    = max(mean_lt),
    .groups = "drop"
  ) %>%
  arrange(macro)

lt_summary_record <- lt %>%
  group_by(macro) %>%
  summarise(
    n           = n(),
    min_days    = min(LT_days),
    mean_days   = mean(LT_days),
    median_days = median(LT_days),
    max_days    = max(LT_days),
    .groups = "drop"
  ) %>%
  arrange(macro)

# Salva in root
readr::write_csv(lt_by_cdm,                file.path(lead_root, "lt_by_cdm_all.csv"))
readr::write_csv(lt_by_cdm_summary,        file.path(lead_root, "lt_by_cdm_summary_by_macro.csv"))
readr::write_csv(lt_summary_record,        file.path(lead_root, "lt_summary_by_macro.csv"))

# Tabelle PNG complessive (INVARIATE)
tab_rec <- gt(lt_summary_record) |>
  tab_header(title = "Lead Time by macro — record-weighted (days)") |>
  fmt_number(columns = c(min_days, mean_days, median_days, max_days), decimals = 1) |>
  cols_label(n="N", min_days="Min", mean_days="Mean", median_days="Median", max_days="Max") |>
  tab_options(table.font.names="serif", table.font.size=px(16), heading.title.font.size=px(20), data_row.padding=px(6))
save_gt_png(tab_rec, file.path(lead_root, "LT_summary_all.png"), vwidth = 2000)

tab_cdm <- gt(lt_by_cdm_summary) |>
  tab_header(title = "Lead Time by macro — CDM-level (mean across CDM)") |>
  fmt_number(columns = c(min_days, mean_days, median_days, max_days), decimals = 1) |>
  cols_label(n_cdms="CDM count", min_days="Min mean", mean_days="Mean of means", median_days="Median of means", max_days="Max mean") |>
  tab_options(table.font.names="serif", table.font.size=px(16), heading.title.font.size=px(20), data_row.padding=px(6))
save_gt_png(tab_cdm, file.path(lead_root, "LT_byCDM_summary_all.png"), vwidth = 2000)

# ---------------------------
# Funzione per salvare artefatti per ciascun macro
# ---------------------------
process_macro <- function(macro_name){
  mdir     <- macro_dirs[[macro_name]]
  csv_dir  <- file.path(mdir, "csv")
  tab_dir  <- file.path(mdir, "tables_png")
  plot_dir <- file.path(mdir, "plots")
  
  df_rec   <- lt %>% filter(macro == macro_name)
  df_cdm   <- lt_by_cdm %>% filter(macro == macro_name)
  
  # --- CSV per-record e summary (record-weighted)
  per_rec <- df_rec %>% select(cdm, macro, donation_date, transfusion_date, LT_days, ulss, ospedale, any_of("reparto"))
  readr::write_csv(per_rec, file.path(csv_dir, glue("lt_per_record_{macro_name}.csv")))
  
  summ_rec <- df_rec %>%
    summarise(n=n(), min_days=min(LT_days), mean_days=mean(LT_days), median_days=median(LT_days), max_days=max(LT_days))
  readr::write_csv(summ_rec, file.path(csv_dir, glue("lt_summary_{macro_name}.csv")))
  
  # --- CSV per-CDM e summary (unweighted by CDM)
  readr::write_csv(df_cdm, file.path(csv_dir, glue("lt_by_cdm_{macro_name}.csv")))
  summ_cdm <- df_cdm %>%
    summarise(n_cdms=n(), min_days=min(mean_lt), mean_days=mean(mean_lt), median_days=median(mean_lt), max_days=max(mean_lt))
  readr::write_csv(summ_cdm, file.path(csv_dir, glue("lt_by_cdm_summary_{macro_name}.csv")))
  
  # --- Tabelle PNG (INVARIATE)
  tab1 <- gt(summ_rec) |>
    tab_header(title = glue("{macro_name} — Lead Time (record-weighted)")) |>
    fmt_number(columns = c(min_days, mean_days, median_days, max_days), decimals = 1) |>
    cols_label(n="N", min_days="Min", mean_days="Mean", median_days="Median", max_days="Max") |>
    tab_options(table.font.names="serif", table.font.size=px(16), heading.title.font.size=px(20), data_row.padding=px(6))
  save_gt_png(tab1, file.path(tab_dir, glue("LT_summary_{macro_name}.png")), vwidth = 1900)
  
  tab2 <- gt(summ_cdm) |>
    tab_header(title = glue("{macro_name} — Lead Time by CDM (mean across CDM)")) |>
    fmt_number(columns = c(min_days, mean_days, median_days, max_days), decimals = 1) |>
    cols_label(n_cdms="CDM count", min_days="Min mean", mean_days="Mean of means", median_days="Median of means", max_days="Max mean") |>
    tab_options(table.font.names="serif", table.font.size=px(16), heading.title.font.size=px(20), data_row.padding=px(6))
  save_gt_png(tab2, file.path(tab_dir, glue("LT_byCDM_summary_{macro_name}.png")), vwidth = 1900)
  
  # ---------------------------
  # GRAFICI (solo titolo = macro_name; NESSUN sottotitolo)
  # ---------------------------
  accent   <- macro_accent[[macro_name]]
  fill_col <- scales::alpha(accent, 0.35)
  dens_col <- accent
  line_col <- accent
  
  # --- 1) Distribuzione LT (per record)
  if (nrow(df_rec) > 0) {
    bw   <- fd_binwidth(df_rec$LT_days)
    mavg <- summ_rec$mean_days
    mmed <- summ_rec$median_days
    q    <- quantile(df_rec$LT_days, probs = c(0.25, 0.5, 0.75), na.rm = TRUE)
    
    g1 <- ggplot(df_rec, aes(x = LT_days)) +
      annotate("rect", xmin = q[1], xmax = q[3], ymin = -Inf, ymax = Inf,
               fill = scales::alpha(accent, 0.08), colour = NA) +
      geom_histogram(binwidth = bw, boundary = 0, fill = fill_col, colour = "grey25") +
      geom_density(aes(y = ..count..), linewidth = 1.2, colour = dens_col) +
      geom_vline(xintercept = mavg, linetype = "solid",  linewidth = 1.0, colour = line_col) +
      geom_vline(xintercept = mmed, linetype = "dashed", linewidth = 1.0, colour = line_col) +
      annotate("label", x = mavg, y = Inf, vjust = 1.2,
               label = glue("mean = {round(mavg,1)}"), size = 3.8,
               label.size = 0, fill = "white", colour = line_col) +
      annotate("label", x = mmed, y = Inf, vjust = 2.4,
               label = glue("median = {round(mmed,1)}"), size = 3.8,
               label.size = 0, fill = "white", colour = line_col) +
      scale_x_continuous(labels = scales::comma_format(accuracy = 1), breaks = scales::pretty_breaks(n = 10)) +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = macro_name,    # <<<<<< SOLO EMOCOMPONENTE
        x = "Lead Time (days)", y = "Count"
      ) +
      theme_thesis + theme_axes_big +
      coord_cartesian(clip = "off")
    
    ggsave(file.path(plot_dir, glue("LT_distribution_{macro_name}.png")),
           g1, width = 13, height = 6.8, dpi = 300, bg = "white")
  }
  
  # --- 2) Distribuzione delle medie per CDM
  if (nrow(df_cdm) > 0) {
    bw2 <- fd_binwidth(df_cdm$mean_lt)
    q2  <- quantile(df_cdm$mean_lt, probs = c(0.25, 0.5, 0.75), na.rm = TRUE)
    
    g2 <- ggplot(df_cdm, aes(x = mean_lt)) +
      annotate("rect", xmin = q2[1], xmax = q2[3], ymin = -Inf, ymax = Inf,
               fill = scales::alpha(accent, 0.08), colour = NA) +
      geom_histogram(binwidth = bw2, boundary = 0, fill = fill_col, colour = "grey25") +
      geom_density(aes(y = ..count..), linewidth = 1.2, colour = dens_col) +
      geom_vline(xintercept = q2[2], linetype = "dashed", linewidth = 1.0, colour = line_col) +
      annotate("label", x = q2[2], y = Inf, vjust = 1.2,
               label = glue("median = {round(q2[2],1)}"),
               size = 3.8, label.size = 0, fill = "white", colour = line_col) +
      scale_x_continuous(labels = scales::comma_format(accuracy = 1), breaks = scales::pretty_breaks(n = 10)) +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = macro_name,   # <<<<<< SOLO EMOCOMPONENTE
        x = "CDM mean LT (days)", y = "CDM count"
      ) +
      theme_thesis + theme_axes_big +
      coord_cartesian(clip = "off")
    
    ggsave(file.path(plot_dir, glue("LT_byCDM_distribution_{macro_name}.png")),
           g2, width = 13, height = 6.8, dpi = 300, bg = "white")
  }
  
  invisible(list(record_summary = summ_rec, bycdm_summary = summ_cdm))
}

# Esegui per ciascun macro
ret <- lapply(macro_levels, process_macro)

# ---------------------------
# Console finale
# ---------------------------
cat("\n================= LEAD TIME SUMMARY — record-weighted (days) =================\n")
print(lt_summary_record, n = nrow(lt_summary_record))
cat("-----------------------------------------------------------------------------\n")
cat("CDM-level (mean across CDM) summary:\n")
print(lt_by_cdm_summary, n = nrow(lt_by_cdm_summary))
cat("-----------------------------------------------------------------------------\n")
cat("Exclusions: missing/NA or macro =", excluded_na, " | negative LT =", excluded_neg, "\n", sep = "")
cat("Outputs at: ", lead_root, "\n", sep = "")
