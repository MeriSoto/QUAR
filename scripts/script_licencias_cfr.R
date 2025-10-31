# -----------------------------------------------------------------------------
# Carga de datos de Excel a PostgreSQL (tabla: licencias_cfr)
# -----------------------------------------------------------------------------

library(readxl)
library(dplyr)
library(tools)
library(DBI)
library(RPostgres)
library(lubridate)

# 1. Seleccionar archivo
archivo_excel <- file.choose()
extension <- tolower(file_ext(archivo_excel))

# 2. Convertir .xlsb a .xlsx si es necesario
if (extension == "xlsb") {
  cat("Convirtiendo .xlsb a .xlsx...\n")
  archivo_xlsx <- sub("\\.xlsb$", ".xlsx", archivo_excel, ignore.case = TRUE)
  comando <- sprintf('soffice --headless --convert-to xlsx "%s" --outdir "%s"', archivo_excel, dirname(archivo_excel))
  system(comando, intern = TRUE)
  archivo_excel <- archivo_xlsx
  extension <- "xlsx"
  cat("Conversión completa →", archivo_excel, "\n")
}

# 3. Leer la primera hoja
hoja <- excel_sheets(archivo_excel)[1]
raw <- read_excel(archivo_excel, sheet = hoja)
names(raw) <- make.names(names(raw), unique = TRUE)

# 4. Columnas esperadas en PostgreSQL
columnas_esperadas_pg <- c(
  "licencia", "categoriaacuerdo", "anio", "codigocfr", "nombrebuque",
  "1trim", "2trim", "3trim", "4trim", "anual",
  "parobiologico", "observaciones", "otronombre"
)

# 5. Relación para renombrar desde Excel a PostgreSQL
nombres_equivalentes <- c(
  "LICENCIA" = "licencia",
  "CATEGORIA.ACUERDO" = "categoriaacuerdo",
  "AÑO" = "anio",
  "CodigoCFR" = "codigocfr",
  "BUQUE" = "nombrebuque",
  "X1TRIM" = "1TRIM",
  "X2TRIM" = "2TRIM",
  "X3TRIM" = "3TRIM",
  "X4TRIM" = "4TRIM",
  "ANUAL" = "anual",
  "PARO.BIOLOGICO" = "parobiologico",
  "Observaciones" = "observaciones",
  "Otro.nombre" = "otronombre"
)

# 6. Renombrar columnas usando nombres de PostgreSQL (en mayúsculas)
df <- raw
for (col_excel in names(nombres_equivalentes)) {
  if (col_excel %in% colnames(df)) {
    names(df)[which(colnames(df) == col_excel)] <- nombres_equivalentes[[col_excel]]
  }
}

# 7. Filtrar solo columnas válidas
columnas_esperadas_pg <- unname(nombres_equivalentes)
df <- df[, intersect(names(df), columnas_esperadas_pg)]

# 8. Convertir columnas numéricas si existen
cols_numericas <- c("AÑO", "1TRIM", "2TRIM", "3TRIM", "4TRIM", "ANUAL")
cols_presentes <- intersect(cols_numericas, names(df))
df[cols_presentes] <- lapply(df[cols_presentes], function(x) suppressWarnings(as.integer(x)))
# 6. Renombrar columnas si están en el archivo
colnames(raw) <- make.names(colnames(raw))
colnames(raw) <- tolower(colnames(raw))
df <- raw
for (col_excel in names(nombres_equivalentes)) {
  col_r <- tolower(make.names(col_excel))
  if (col_r %in% colnames(df)) {
    names(df)[which(colnames(df) == col_r)] <- nombres_equivalentes[[col_excel]]
  }
}

# 7. Filtrar columnas válidas
df <- df[, intersect(names(df), columnas_esperadas_pg)]

# 8. Convertir algunas columnas a integer
cols_numericas <- c("anio", "1TRIM", "2TRIM", "3TRIM", "4TRIM", "anual")
df[cols_numericas] <- lapply(df[cols_numericas], function(x) suppressWarnings(as.integer(x)))

# 9. Conectar a PostgreSQL (NEON)
con <- dbConnect(RPostgres::Postgres(),
                 dbname = "neondb",
                 host = "ep-plain-fire-aba8ctis.eu-west-2.aws.neon.tech",
                 port = 5432,
                 user = "editor",
                 password = "CECAF_editor",
                 sslmode = "require")

# 10. Subir a tabla temporal
tabla_temp <- "temp_licencias_cfr"
dbWriteTable(con, name = tabla_temp, value = df, overwrite = TRUE, row.names = FALSE)

# 11. Insertar ignorando duplicados, con comillas en columnas
columnas_insert <- paste0('"', colnames(df), '"', collapse = ", ")
sql_insert <- sprintf("
  INSERT INTO licencias_cfr (%s)
  SELECT %s FROM %s
  ON CONFLICT DO NOTHING;
", columnas_insert, columnas_insert, tabla_temp)

dbExecute(con, sql_insert)

# 12. Eliminar tabla temporal
dbExecute(con, sprintf("DROP TABLE IF EXISTS %s;", tabla_temp))

# 13. Cerrar conexión
dbDisconnect(con)

cat("✅ Datos cargados en PostgreSQL → tabla licencias_cfr.\n")
