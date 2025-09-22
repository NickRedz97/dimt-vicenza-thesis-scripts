# =============================================================================
# Script name : raccolta_excel_clean_dedup_group_fill.R
# Purpose     : Clean and standardize "RACCOLTA" Excel files, perform controlled
#               de-duplication, and attach ABO/Rh groups via ULSS normalized CSVs,
#               with SCARTATO as fallback. Outputs are written to /raccolta/normalized/.
# Inputs      : - RACCOLTA Excel files: "<raccolta/dataset>/raccolta 2023.xlsx" ... "raccolta 2025.xlsx"
#               - Normalized TRASFUSO CSV: ULSS7.csv, ULSS8.csv
#               - Normalized SCARTATO CSV: scartato_merged_cleaned.csv
# Outputs     : - raccolta_merged_clean.csv
#               - raccolta_merged_clean_with_group.csv
#               - raccolta_duplicati_examples.csv
#               - ulss_cdm_conflicts.csv
#               - raccolta_merged_with_group_keys.csv
#               - raccolta_merged_keys.csv
# Scope       : Donation-level records with keys {id_donatore, data_prestazione_date, cdm, emocomponente}.
# Notes       : - Date parsing supports multiple formats with robust fallbacks.
#               - ABO/Rh mapping: ULSS is primary source; SCARTATO provides fallback.
# =============================================================================

# --- Packages -----------------------------------------------------------------
pkg_needed <- c("tidyverse","janitor","lubridate","stringi","readxl")
for (p in pkg_needed) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
library(tidyverse); library(janitor); library(lubridate); library(stringi); library(readxl)

# =============================================================================
# RACCOLTA — Excel-first workflow (RACCOLTA only), CSV conversion (RACCOLTA only)
# Group mapping sources: TRASFUSO (normalized/ULSS7.csv, ULSS8.csv) + SCARTATO
# Outputs remain identical in: .../raccolta/normalized/
# =============================================================================

# --- Packages (duplication preserved intentionally; no logic changes) ----------
pkg_needed <- c("tidyverse","janitor","lubridate","stringi","readxl")
for (p in pkg_needed) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
library(tidyverse); library(janitor); library(lubridate); library(stringi); library(readxl)

# --- Paths --------------------------------------------------------------------
rac_base <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/raccolta"
rac_in   <- file.path(rac_base, "dataset")
rac_out  <- file.path(rac_base, "normalized"); if (!dir.exists(rac_out)) dir.create(rac_out, TRUE)

# Normalized TRASFUSO (pre-produced by dedicated scripts)
ulss7_csv <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized/ULSS7.csv"
ulss8_csv <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso/normalized/ULSS8.csv"

# Normalized SCARTATO (pre-produced by dedicated scripts)
scartato_csv <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/scartato/normalized/scartato_merged_cleaned.csv"

# RACCOLTA input: Excel-only files (one per year)
rac_xlsx <- file.path(rac_in, sprintf("raccolta %d.xlsx", 2023:2025))
if (!all(file.exists(rac_xlsx))) {
  stop("Mancano uno o più Excel RACCOLTA in: ", rac_in, 
       ". Attesi: ", paste(basename(rac_xlsx), collapse=", "))
}

# --- Output file definitions (unchanged) --------------------------------------
out_merged_file        <- file.path(rac_out, "raccolta_merged_clean.csv")
out_merged_with_group  <- file.path(rac_out, "raccolta_merged_clean_with_group.csv")
out_dup_examples       <- file.path(rac_out, "raccolta_duplicati_examples.csv")
out_ulss_conflicts     <- file.path(rac_out, "ulss_cdm_conflicts.csv")
out_control_keys       <- file.path(rac_out, "raccolta_merged_with_group_keys.csv")
out_keys_file          <- file.path(rac_out, "raccolta_merged_keys.csv")

# --- Utilities ----------------------------------------------------------------
# Standardize column names (ASCII, snake_case, trimmed)
clean_names_safe <- function(df) {
  names(df) <- names(df) %>%
    stri_trans_general("Latin-ASCII") %>% tolower() %>%
    str_replace_all("\\s+", "_") %>% str_replace_all("[^0-9a-z_]", "_") %>%
    str_replace_all("_+", "_") %>% str_replace_all("^_|_$", "")
  df
}

# Normalize character columns (UTF-8, trimmed, NA preserved)
clean_char_cols <- function(df) {
  df %>% mutate(across(where(is.character),
                       ~ ifelse(is.na(.x), NA_character_,
                                str_squish(iconv(.x, from="latin1", to="UTF-8")) )))
}

# Supported date-time parsing patterns (multiple locales and granularities)
fallback_orders <- c("dmy HMS","dmy HM","dmy H:M:S","dmy H:M","dmy",
                     "mdy HMS","mdy HM","mdy H:M:S","mdy H:M","mdy",
                     "Ymd HMS","Ymd HM","Ymd")

# Robust date-time parser: handles Excel numerics and varied string formats
robust_parse_datetime <- function(vec) {
  if (is.numeric(vec)) return(as_datetime(as.Date(vec, origin="1899-12-30")))
  parsed <- parse_date_time(vec, orders=fallback_orders, exact=FALSE, quiet=TRUE)
  if (any(is.na(parsed) & !is.na(vec))) {
    try_vec <- str_replace_all(vec, "-", "/")
    parsed2 <- parse_date_time(try_vec, orders=fallback_orders, exact=FALSE, quiet=TRUE)
    parsed[is.na(parsed) & !is.na(parsed2)] <- parsed2[is.na(parsed) & !is.na(parsed2)]
  }
  parsed
}

# Non-NA statistical mode, returning a single representative value
mode_non_na <- function(x) {
  x <- x[!is.na(x)]; if (!length(x)) return(NA_character_)
  ux <- unique(x); ux[which.max(tabulate(match(x, ux)))]
}

# ABO/Rh string normalization (uppercase, sign unification, whitespace removal)
normalize_group_str <- function(x) {
  x2 <- as.character(x); x2[is.na(x2)] <- NA_character_
  x2 <- str_squish(toupper(x2))
  x2 <- str_replace_all(x2, "POSITIVO", "+")
  x2 <- str_replace_all(x2, "NEGATIVO", "-")
  x2 <- str_replace_all(x2, "\\s+POS$", "+")
  x2 <- str_replace_all(x2, "\\s+NEG$", "-")
  x2 <- str_replace_all(x2, " ", "")
  x2 <- str_replace_all(x2, "APOS$", "A+")
  x2 <- str_replace_all(x2, "ANEG$", "A-")
  x2
}

# --- Read RACCOLTA from Excel + CSV conversion (RACCOLTA only) ----------------
read_raccolta_excel <- function(path_xlsx) {
  message("Leggo RACCOLTA: ", path_xlsx)
  df <- read_excel(path_xlsx, sheet = 1) %>% clean_names_safe() %>% clean_char_cols()
  # CSV side-output in the same folder (as required)
  out_csv <- sub("\\.xlsx?$", ".csv", path_xlsx, ignore.case = TRUE)
  suppressWarnings(readr::write_csv(df, out_csv))
  # Standardizations aligned with the original workflow
  if ("stato_donazione" %in% names(df)) {
    df <- df %>% mutate(stato_donazione = ifelse(is.na(stato_donazione), NA_character_,
                                                 str_to_title(str_squish(str_to_lower(stato_donazione)))))
  }
  for (col in intersect(c("id_donatore","numero_presentazioni","numero_emocomponenti"), names(df))) {
    df[[col]] <- df[[col]] %>% as.character() %>%
      str_replace_all("\\.", "") %>% str_replace_all(",", ".") %>% na_if("") %>% as.numeric()
  }
  if ("data_prestazione" %in% names(df)) {
    parsed <- robust_parse_datetime(df$data_prestazione)
    df$data_prestazione_dt   <- parsed
    df$data_prestazione_date <- as_date(parsed)
  }
  if ("data_nascita_donatore" %in% names(df)) {
    parsed2 <- robust_parse_datetime(df$data_nascita_donatore)
    df$data_nascita_donatore_dt   <- parsed2
    df$data_nascita_donatore_date <- as_date(parsed2)
  }
  df %>% mutate(source_file = basename(path_xlsx),
                anno_file   = str_extract(basename(path_xlsx), "\\d{4}") %>% as.integer())
}

processed <- lapply(rac_xlsx, read_raccolta_excel)
merged    <- bind_rows(processed)

# --- Deduplication and intermediate outputs -----------------------------------
# Key uniqueness defined on donor/date/cdm/component
key_cols <- c("id_donatore","data_prestazione_date","cdm","emocomponente")
missing_keys <- setdiff(key_cols, names(merged))
if (length(missing_keys) > 0) stop("Mancano colonne chiave in RACCOLTA: ", paste(missing_keys, collapse=", "))

# Duplicate groups summary (examples exported for inspection)
dup_groups <- merged %>%
  group_by(id_donatore, data_prestazione_date, cdm, emocomponente) %>%
  tally(name = "n_rows") %>% filter(n_rows > 1) %>% arrange(desc(n_rows))
if (nrow(dup_groups) > 0) {
  merged %>%
    semi_join(dup_groups, by = c("id_donatore","data_prestazione_date","cdm","emocomponente")) %>%
    arrange(id_donatore, data_prestazione_date, cdm, emocomponente, desc(data_prestazione_dt)) %>%
    readr::write_csv(out_dup_examples)
}

# Dedup strategy (preserved): choose preferred record within each key-group
dedup_strategy <- "prefer_conclusa"; prefer_stato <- "Conclusa"
if (dedup_strategy == "keep_all") {
  merged_dedup <- merged
} else if (dedup_strategy == "prefer_conclusa") {
  merged_dedup <- merged %>%
    arrange(id_donatore, data_prestazione_date, cdm, emocomponente,
            desc(if_else(stato_donazione == prefer_stato, 1L, 0L)),
            desc(replace_na(numero_emocomponenti, 0)),
            desc(data_prestazione_dt)) %>%
    group_by(id_donatore, data_prestazione_date, cdm, emocomponente) %>%
    slice_head(n = 1) %>% ungroup()
} else if (dedup_strategy == "max_num_emocomponenti") {
  merged_dedup <- merged %>%
    arrange(id_donatore, data_prestazione_date, cdm, emocomponente,
            desc(replace_na(numero_emocomponenti, 0)),
            desc(if_else(stato_donazione == prefer_stato, 1L, 0L)),
            desc(data_prestazione_dt)) %>%
    group_by(id_donatore, data_prestazione_date, cdm, emocomponente) %>%
    slice_head(n = 1) %>% ungroup()
} else stop("dedup_strategy non riconosciuta.")
readr::write_csv(merged_dedup, out_merged_file)

# --- ABO/Rh mapping from ULSS normalized CSVs ---------------------------------
if (!file.exists(ulss7_csv) || !file.exists(ulss8_csv)) {
  stop("Non trovo i CSV normalizzati ULSS in trasfuso/normalized/. Attesi: ULSS7.csv e ULSS8.csv")
}
ulss7 <- readr::read_csv(ulss7_csv, guess_max = 20000, show_col_types = FALSE) %>% clean_names_safe() %>% clean_char_cols()
ulss8 <- readr::read_csv(ulss8_csv, guess_max = 20000, show_col_types = FALSE) %>% clean_names_safe() %>% clean_char_cols()

# Detect column names for CDM and group (robust to naming variants)
detect_cdm_and_group_cols <- function(df) {
  nm <- names(df)
  cdm_idx   <- grep("^cdm$|\\bcdm\\b", nm, ignore.case = TRUE)
  group_idx <- grep("grupp|ab0|ab0_rh|gruppo_ab0|gruppo", nm, ignore.case = TRUE)
  list(cdm   = if (length(cdm_idx)) nm[cdm_idx[1]] else NULL,
       group = if (length(group_idx)) nm[group_idx[1]] else NULL)
}
d7 <- detect_cdm_and_group_cols(ulss7)
d8 <- detect_cdm_and_group_cols(ulss8)
if (is.null(d7$cdm) | is.null(d7$group)) stop("ULSS7: non trovo colonne CDM/Gruppo nei CSV normalizzati.")
if (is.null(d8$cdm) | is.null(d8$group)) stop("ULSS8: non trovo colonne CDM/Gruppo nei CSV normalizzati.")

ulss7_sel <- ulss7 %>%
  select(cdm = all_of(d7$cdm), gruppo_ab0_rh = all_of(d7$group)) %>%
  mutate(cdm = str_squish(as.character(cdm)), gruppo_ab0_rh = str_squish(as.character(gruppo_ab0_rh)))
ulss8_sel <- ulss8 %>%
  select(cdm = all_of(d8$cdm), gruppo_ab0_rh = all_of(d8$group)) %>%
  mutate(cdm = str_squish(as.character(cdm)), gruppo_ab0_rh = str_squish(as.character(gruppo_ab0_rh)))

ulss_all <- bind_rows(ulss7_sel %>% mutate(source = "ULSS7"),
                      ulss8_sel %>% mutate(source = "ULSS8")) %>%
  filter(!is.na(cdm) & cdm != "") %>%
  distinct(cdm, gruppo_ab0_rh, source)

# Resolve cross-source inconsistencies at CDM level; export conflicts if any
ulss_grouped <- ulss_all %>%
  group_by(cdm) %>%
  summarise(groups = list(unique(na.omit(gruppo_ab0_rh))),
            n_groups = length(unique(na.omit(gruppo_ab0_rh))),
            .groups = "drop")
conflicts <- ulss_grouped %>% filter(n_groups > 1)
if (nrow(conflicts) > 0) {
  conflicts_details <- ulss_all %>%
    semi_join(conflicts, by = "cdm") %>%
    arrange(cdm, gruppo_ab0_rh, source)
  readr::write_csv(conflicts_details, out_ulss_conflicts)
  ulss_resolved <- ulss_all %>% group_by(cdm) %>% summarise(gruppo_ab0_rh = mode_non_na(gruppo_ab0_rh), .groups="drop")
} else {
  ulss_resolved <- ulss_all %>% group_by(cdm) %>% summarise(gruppo_ab0_rh = first(na.omit(gruppo_ab0_rh)), .groups="drop")
}
ulss_resolved <- ulss_resolved %>% distinct(cdm, gruppo_ab0_rh)

# --- SCARTATO mapping (fallback) ----------------------------------------------
if (!file.exists(scartato_csv)) stop("Non trovo SCARTATO normalizzato: scartato_merged_cleaned.csv")
scartato <- readr::read_csv(scartato_csv, guess_max = 20000, show_col_types = FALSE) %>%
  clean_names_safe() %>% clean_char_cols()

d_s <- detect_cdm_and_group_cols(scartato)
if (is.null(d_s$cdm) | is.null(d_s$group)) {
  stop("SCARTATO: non trovo colonne CDM/Gruppo nel CSV normalizzato.")
}
scartato_map <- scartato %>%
  rename(cdm = all_of(d_s$cdm), scartato_group_raw = all_of(d_s$group)) %>%
  mutate(cdm = str_squish(as.character(cdm)),
         scartato_group = normalize_group_str(scartato_group_raw)) %>%
  filter(!is.na(cdm) & cdm != "") %>%
  group_by(cdm) %>% summarise(scartato_group = mode_non_na(scartato_group), .groups = "drop") %>%
  distinct(cdm, scartato_group)

# --- Final join: ULSS primary, SCARTATO fallback ------------------------------
merged_join_ready <- merged_dedup %>% mutate(cdm = str_squish(as.character(cdm)))
merged_with_ulss  <- merged_join_ready %>%
  left_join(ulss_resolved %>% mutate(cdm = as.character(cdm)) %>% rename(ulss_group = gruppo_ab0_rh),
            by = "cdm") %>%
  mutate(ulss_group_norm = normalize_group_str(ulss_group))

merged_with_group <- merged_with_ulss %>%
  left_join(scartato_map, by = "cdm") %>%
  mutate(
    gruppo_ab0_rh_final = coalesce(ulss_group_norm, scartato_group),
    gruppo_source       = case_when(
      !is.na(ulss_group_norm) ~ "ULSS",
      is.na(ulss_group_norm) & !is.na(scartato_group) ~ "SCARTATO",
      TRUE ~ NA_character_
    ),
    gruppo_ab0_rh = gruppo_ab0_rh_final
  ) %>%
  select(-ulss_group, -ulss_group_norm, -scartato_group, -gruppo_ab0_rh_final)

# --- Final outputs ------------------------------------------------
readr::write_csv(merged_with_group, out_merged_with_group)

control_cols <- intersect(
  c("id_donatore","data_prestazione_date","cdm","emocomponente",
    "gruppo_ab0_rh","gruppo_source","source_file","anno_file"),
  names(merged_with_group)
)
if (length(control_cols) > 0) {
  readr::write_csv(merged_with_group %>% select(all_of(control_cols)), out_control_keys)
}
# Full "keys" file (compatibility)
readr::write_csv(merged_with_group, out_keys_file)

cat("\n--- COMPLETATO: RACCOLTA da Excel; TRASFUSO/SCARTATO da normalized CSV; output in raccolta/normalized ---\n")
