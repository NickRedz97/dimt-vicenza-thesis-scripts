# =============================================================================
# Script name : normalizza_cessioni.R  (visual refresh per uniformità)
# Purpose     : Normalize "cessioni" Excel files into a unified CSV and produce
#               descriptive figures and tables for RBCs, Plasma, and Platelets.
# Change req. : RIMOSSI TUTTI I TITOLI DAI GRAFICI. Etichette % nei pie chart
#               più grandi/visibili, legende più visibili (dimensione maggiore).
#               Nessuna modifica alla logica dei dati.
# =============================================================================

suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr)
  library(lubridate); library(readr); library(purrr); library(tidyr)
  library(ggplot2); library(gt)
})

base_dir    <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/cessioni"
dataset_dir <- file.path(base_dir, "dataset")
out_dir     <- file.path(base_dir, "normalized")
out_file    <- file.path(out_dir, "cessioni_2023_2025.csv")

out_graphs  <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/cessioni/outputs"
dir.create(out_graphs, showWarnings = FALSE, recursive = TRUE)

# -----------------------------------------------------------------------------
# Word export sizing — A4 portrait + 3 pt side margins (exact text width)
# -----------------------------------------------------------------------------
page_width_in   <- 8.27
left_margin_pt  <- 3
right_margin_pt <- 3
target_ppi      <- 96

content_width_in <- page_width_in - (left_margin_pt + right_margin_pt)/72
content_width_px <- round(content_width_in * target_ppi)

# -----------------------------------------------------------------------------
# Robust date parser (Excel serials and common textual patterns)
# -----------------------------------------------------------------------------
parse_data_sicura <- function(x) {
  if (is.numeric(x)) {
    as.Date(x, origin = "1899-12-30")
  } else {
    x_chr <- as.character(x)
    dt <- suppressWarnings(parse_date_time(
      x_chr,
      orders = c("Y-m-d","d/m/Y","m/d/Y","d-m-Y","Ymd","d/m/Y HMS","Y-m-d HMS","d-m-Y HMS"),
      tz = "UTC"
    ))
    as.Date(dt)
  }
}

# -----------------------------------------------------------------------------
# Import and normalization (column standardization and types)
# -----------------------------------------------------------------------------
import_cessioni <- function(path_xlsx) {
  read_excel(path_xlsx) %>%
    rename(
      dipartimento          = "Dipartimento",
      ulss                  = "ULSS",
      mese_src              = "Mese",
      destinatario          = "Destinatario",
      tipo_movimento        = "Tipo Movimento",
      struttura_acquirente  = "Struttura Acquirente",
      emocomponente         = "Emocomponente",
      tipo_donazione        = "Tipo Donazione",
      settimana_src         = "Settimana",
      data_raw              = "Data",
      struttura_cedente     = "Struttura Cedente",
      numero_unita_cedute   = "Numero Unità Cedute"
    ) %>%
    mutate(
      data      = parse_data_sicura(data_raw),
      mese      = as.numeric(month(data)),
      anno      = as.numeric(year(data)),
      settimana = ifelse(is.na(data), NA_character_,
                         sprintf("%d/%02d", isoyear(data), isoweek(data))),
      across(
        c(dipartimento, ulss, destinatario, tipo_movimento,
          struttura_acquirente, emocomponente, tipo_donazione,
          struttura_cedente),
        ~ ifelse(is.na(.x), .x, str_squish(as.character(.x)))
      ),
      numero_unita_cedute = suppressWarnings(as.numeric(numero_unita_cedute))
    ) %>%
    select(
      dipartimento, ulss, mese, destinatario, tipo_movimento,
      struttura_acquirente, emocomponente, tipo_donazione,
      settimana, data, struttura_cedente, numero_unita_cedute, anno
    )
}

files <- file.path(dataset_dir, c("cessioni 2023.xlsx","cessioni 2024.xlsx","cessioni 2025.xlsx"))
stopifnot(all(file.exists(files)))

cessioni_all <- purrr::map_dfr(files, import_cessioni)

dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
write_csv(cessioni_all, out_file, na = "")
message("✅ CSV scritto: ", out_file)

# =============================================================================
# REPORT — macro names in EN (RBCs / Plasma / Platelets)
# =============================================================================

# Macro tagging based on component labels
tag_macro <- function(x) {
  dplyr::case_when(
    str_detect(x, regex("\\bEmaz|Eritro|\\bRBC\\b", ignore_case = TRUE)) ~ "RBCs",
    str_detect(x, regex("\\bPlasma\\b|Plasma\\s", ignore_case = TRUE))    ~ "Plasma",
    str_detect(x, regex("Piastrin|\\bPLT\\b|Platelet", ignore_case = TRUE)) ~ "Platelets",
    TRUE ~ "Other"
  )
}

df <- cessioni_all %>%
  mutate(macro = tag_macro(emocomponente)) %>%
  filter(!is.na(data), !is.na(numero_unita_cedute)) %>%
  filter(macro != "Other")

# -----------------------------------------------------------------------------
# Thesis style (uniform across figures) — NO TITLES, legende più grandi
# -----------------------------------------------------------------------------
# Okabe–Ito palette
pal_okabe <- c(
  "RBCs"      = "#0072B2",
  "Plasma"    = "#009E73",
  "Platelets" = "#D55E00"
)

theme_thesis <- theme_minimal(base_family = "serif", base_size = 12) +
  theme(
    plot.title   = element_blank(),              # <<— no titles
    plot.subtitle= element_blank(),              # <<— no subtitles
    axis.title   = element_text(face = "bold", size = 12),
    axis.text    = element_text(size = 11),
    panel.grid.minor = element_blank(),
    legend.position  = "bottom",
    legend.title     = element_blank(),
    legend.text      = element_text(size = 13),  # <<— più visibile
    legend.key.width  = grid::unit(1.0, "lines"),
    legend.key.height = grid::unit(1.0, "lines"),
    legend.box.margin = margin(8, 24, 8, 24),
    plot.margin       = margin(4, 18, 4, 18)     # margini compatti, niente bianco sopra/sotto
  )

theme_axes_big <- theme(
  axis.text.x  = element_text(size = 11, angle = 45, hjust = 1),
  axis.text.y  = element_text(size = 11),
  axis.title.x = element_text(size = 12),
  axis.title.y = element_text(size = 12)
)

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
# Figure 1 — Monthly trend (quarterly ticks, monthly resolution) — NO TITLE
# -----------------------------------------------------------------------------
df_mens <- df %>%
  mutate(mese_ref = floor_date(data, "month")) %>%
  group_by(macro, mese_ref) %>%
  summarise(units = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  arrange(macro, mese_ref)

min_m <- min(df_mens$mese_ref, na.rm = TRUE)
max_m <- max(df_mens$mese_ref, na.rm = TRUE)
quarter_breaks <- seq(floor_date(min_m, "quarter"), ceiling_date(max_m, "quarter"), by = "3 months")

g1 <- ggplot(df_mens, aes(mese_ref, units, colour = macro, group = macro)) +
  geom_line(linewidth = 1.15) +
  geom_point(size = 2) +
  scale_color_manual(values = pal_okabe) +
  scale_y_continuous(labels = scales::comma) +
  scale_x_date(
    breaks = quarter_breaks,
    labels = function(x) paste0(year(x), "-Q", quarter(x)),
    expand = expansion(mult = c(0.01, 0.02))
  ) +
  labs(x = "Month (quarterly ticks)", y = "Units") +
  theme_thesis + theme_axes_big
ggsave(file.path(out_graphs, "trend_mensile.png"),
       g1, width = content_width_in, height = 5.8, dpi = target_ppi, units = "in")

# -----------------------------------------------------------------------------
# Figure 2 — Distribution of ceded units (top ceding facilities) — NO TITLE
# -----------------------------------------------------------------------------
top5_tbl <- df %>%
  group_by(struttura_cedente) %>%
  summarise(tot = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(tot))
top5_fac <- head(top5_tbl$struttura_cedente, min(5, nrow(top5_tbl)))
df_top <- df %>% filter(struttura_cedente %in% top5_fac)

g2 <- ggplot(
  df_top,
  aes(x = reorder(struttura_cedente, numero_unita_cedute, FUN = sum, na.rm = TRUE),
      y = numero_unita_cedute, fill = macro)
) +
  geom_boxplot(outlier.size = 0.9) +
  scale_fill_manual(values = pal_okabe) +
  coord_flip() +
  labs(x = "Facility", y = "Units") +
  theme_thesis
ggsave(file.path(out_graphs, "boxplot_strutture.png"),
       g2, width = content_width_in, height = 6.2, dpi = target_ppi, units = "in")

# -----------------------------------------------------------------------------
# Figure 3 — Overall share by macro (totals) — NO TITLE
# -----------------------------------------------------------------------------
df_tot <- df %>%
  group_by(macro) %>%
  summarise(units = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(units))

g3 <- ggplot(df_tot, aes(x = reorder(macro, units), y = units, fill = macro)) +
  geom_col(width = 0.75, show.legend = FALSE) +
  scale_fill_manual(values = pal_okabe) +
  coord_flip() +
  labs(x = NULL, y = "Total units") +
  theme_thesis
ggsave(file.path(out_graphs, "share_macro.png"),
       g3, width = content_width_in, height = 4.8, dpi = target_ppi, units = "in")

# -----------------------------------------------------------------------------
# Figure 4 — Top 5 acquiring facilities (totals) — NO TITLE
# -----------------------------------------------------------------------------
df_acq <- df %>%
  filter(!is.na(struttura_acquirente), struttura_acquirente != "") %>%
  group_by(struttura_acquirente) %>%
  summarise(units = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(units)) %>%
  slice_head(n = 5)

g4 <- ggplot(df_acq, aes(x = reorder(struttura_acquirente, units), y = units)) +
  geom_col(fill = "#0072B2", width = 0.75) +
  coord_flip() +
  labs(x = "Acquiring facility", y = "Total units") +
  theme_thesis
ggsave(file.path(out_graphs, "top5_acquirenti.png"),
       g4, width = content_width_in, height = 5.8, dpi = target_ppi, units = "in")

# -----------------------------------------------------------------------------
# Figure 5 — Pie chart: share by movement type — NO TITLE, % più grandi
# -----------------------------------------------------------------------------
df_pie_mov <- df %>%
  filter(!is.na(tipo_movimento), tipo_movimento != "") %>%
  group_by(tipo_movimento) %>%
  summarise(units = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  mutate(pct = units / sum(units),
         lbl = ifelse(pct >= 0.06, paste0(round(pct * 100, 1), "%"), ""))

pal_disc <- scales::hue_pal()(nrow(df_pie_mov))

g5 <- ggplot(df_pie_mov, aes(x = "", y = units, fill = tipo_movimento)) +
  geom_col(width = 1, color = "white") +
  coord_polar(theta = "y") +
  geom_text(aes(label = lbl),
            position = position_stack(vjust = 0.5),
            size = 5.2, fontface = "bold") +   # <<— % più grandi e in grassetto
  scale_fill_manual(values = pal_disc) +
  labs(x = NULL, y = NULL, fill = NULL) +
  theme_thesis +
  theme(
    axis.text = element_blank(), axis.ticks = element_blank(), panel.grid = element_blank(),
    legend.position = "bottom",
    legend.text = element_text(size = 13),                  # <<— legenda più visibile
    legend.box.margin = margin(8, 24, 8, 24)
  ) +
  guides(fill = guide_legend(ncol = 2, byrow = TRUE))
ggsave(file.path(out_graphs, "pie_tipo_movimento.png"),
       g5, width = content_width_in, height = 7.2, dpi = target_ppi, units = "in")

# -----------------------------------------------------------------------------
# Figure 6 — Pie chart: share by recipient — NO TITLE, % più grandi
# -----------------------------------------------------------------------------
df_pie_dest <- df %>%
  filter(!is.na(destinatario), destinatario != "") %>%
  group_by(destinatario) %>%
  summarise(units = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  mutate(pct = units / sum(units),
         lbl = ifelse(pct >= 0.06, paste0(round(pct * 100, 1), "%"), ""))

pal_disc2 <- scales::hue_pal()(nrow(df_pie_dest))

g6 <- ggplot(df_pie_dest, aes(x = "", y = units, fill = destinatario)) +
  geom_col(width = 1, color = "white") +
  coord_polar(theta = "y") +
  geom_text(aes(label = lbl),
            position = position_stack(vjust = 0.5),
            size = 5.2, fontface = "bold") +   # <<— % più grandi e in grassetto
  scale_fill_manual(values = pal_disc2) +
  labs(x = NULL, y = NULL, fill = NULL) +
  theme_thesis +
  theme(
    axis.text = element_blank(), axis.ticks = element_blank(), panel.grid = element_blank(),
    legend.position = "bottom",
    legend.text = element_text(size = 13),                  # <<— legenda più visibile
    legend.box.margin = margin(8, 24, 8, 24)
  ) +
  guides(fill = guide_legend(ncol = 2, byrow = TRUE))
ggsave(file.path(out_graphs, "pie_destinatario.png"),
       g6, width = content_width_in, height = 7.6, dpi = target_ppi, units = "in")

# =============================================================================
# Tables (gt -> PNG) — invariate (titoli mantenuti: non sono grafici)
# =============================================================================

table_pct <- 80

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

save_gt_png <- function(gt_tbl, filename) {
  gt::gtsave(
    gt_tbl,
    file   = file.path(out_graphs, paste0(filename, ".png")),
    vwidth = content_width_px,
    expand = 0
  )
}

# --------------------
# Table 1 — Annual totals by macro
# --------------------
macros <- c("RBCs", "Plasma", "Platelets")

tab1_data <- df %>%
  group_by(anno, macro) %>%
  summarise(Total = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop")

tab1_wide <- tab1_data %>%
  mutate(macro = factor(macro, levels = macros)) %>%
  tidyr::pivot_wider(names_from = macro, values_from = Total, values_fill = 0) %>%
  arrange(anno) %>%
  mutate(Total = rowSums(dplyr::across(all_of(macros)), na.rm = TRUE)) %>%
  dplyr::select(anno, dplyr::all_of(macros), Total)

tab1 <- tab1_wide %>%
  gt() %>%
  tab_header(title = "Annual totals by macro") %>%
  cols_label(
    anno      = "Year",
    RBCs      = "RBCs",
    Plasma    = "Plasma",
    Platelets = "Platelets",
    Total     = "Total"
  ) %>%
  fmt_number(columns = c(RBCs, Plasma, Platelets, Total), decimals = 0, use_seps = TRUE) %>%
  gt_style_thesis()

save_gt_png(tab1, "tabella_totali_annui")

# --------------------
# Table 2 — Top facilities by macro
# --------------------
tab2_data <- df %>%
  group_by(macro, struttura_cedente) %>%
  summarise(Total = sum(numero_unita_cedute, na.rm = TRUE), .groups = "drop") %>%
  group_by(macro) %>%
  slice_max(Total, n = 5, with_ties = FALSE) %>%
  arrange(macro, desc(Total))

tab2 <- tab2_data %>%
  gt(groupname_col = "macro") %>%
  tab_header(title = "Top facilities by macro") %>%
  fmt_number(columns = "Total", decimals = 0, use_seps = TRUE) %>%
  gt_style_thesis()

save_gt_png(tab2, "tabella_top5_strutture")

# --------------------
# Table 3 — Descriptive statistics by macro
# --------------------
tab3_data <- df %>%
  group_by(macro) %>%
  summarise(Mean = mean(numero_unita_cedute, na.rm = TRUE),
            SD   = sd(numero_unita_cedute,   na.rm = TRUE),
            .groups = "drop")

tab3 <- tab3_data %>%
  gt() %>%
  tab_header(title = "Mean and standard deviation of ceded units") %>%
  fmt_number(columns = c(Mean, SD), decimals = 1, use_seps = TRUE) %>%
  gt_style_thesis()

save_gt_png(tab3, "tabella_statistiche")

message("📊 Report esportato in: ", out_graphs)
