# ------------------------
# Pacchetti
# ------------------------
library(readxl)
library(dplyr)
library(lubridate)

# ------------------------
# Configurazione percorsi (AGGIORNATA)
# ------------------------
base_path   <- "C:/Users/nicolo.rossi/Desktop/Elaborazioni Tesi/R tesi/trasfuso"
input_path  <- file.path(base_path, "dataset")
output_path <- file.path(base_path, "normalized")

if (!dir.exists(output_path)) dir.create(output_path, recursive = TRUE)

# File di input presenti in 'trasfuso/dataset'
files <- c("ULSS7.xlsx", "ULSS8.xlsx")
names(files) <- c("ULSS7", "ULSS8")

# Utility: stampa check esistenza file
for (nm in names(files)) {
  cat("Check input", nm, "->", file.path(input_path, files[nm]), 
      "exists:", file.exists(file.path(input_path, files[nm])), "\n")
}

# ------------------------
# Funzione di normalizzazione (logica identica agli script originali)
# ------------------------
process_file <- function(file_xls, nome) {
  message("\n--- Inizio normalizzazione: ", nome, " ---")
  
  # 1) Leggi il primo foglio
  df <- read_excel(file_xls, sheet = 1)
  
  # 2) Mostra i nomi delle colonne per verifica
  cat("Colonne importate:\n")
  print(names(df))
  
  # 3) Identifica i nomi esatti delle colonne data
  cols_date <- names(df)[ 
    tolower(names(df)) %in% 
      c("data richiesta", "data rich", "data_richiesta",
        "data consegna", "dataconsegna",
        "data trasfusione", "datatrasfusione")
  ]
  
  # Controllo di sicurezza
  if (!all(c("data richiesta", "data consegna", "data trasfusione") %in% tolower(cols_date))) {
    stop("Non ho trovato tutte e tre le colonne data (Richiesta, Consegna, Trasfusione). Rivedi i nomi stampati sopra in ", nome)
  }
  
  # 4) Converte tutte le colonne individuate in Date (ISO)
  df <- df %>% 
    mutate(across(
      all_of(cols_date),
      ~ as.Date(.x, format = "%Y-%m-%d"),
      .names = "{.col}"
    ))
  
  # 5) Controlla la conversione
  cat("Classi delle colonne data dopo mutate:\n")
  print(sapply(df[cols_date], class))
  
  # 6) Esporta il CSV normalizzato in 'trasfuso/normalized'
  output_csv <- file.path(output_path, paste0(nome, ".csv"))
  write.csv(df, output_csv, row.names = FALSE, fileEncoding = "UTF-8")
  
  message("Conversione completata: ", output_csv)
}

# ------------------------
# Esecuzione per ULSS7 e ULSS8
# ------------------------
for (nm in names(files)) {
  process_file(file.path(input_path, files[nm]), nm)
}

cat("\n--- Script completato: ULSS7 e ULSS8 normalizzati ---\n")
