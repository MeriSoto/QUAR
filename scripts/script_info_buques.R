# Autor: Carlos García Moya
# -------------------------------------------------------------------
# Script para cargar datos en la tabla info_buques desde Excel a PostgreSQL
# -------------------------------------------------------------------

library(readxl)
library(dplyr)
library(tools)
library(DBI)
library(RPostgres)
library(lubridate)

# 1. Seleccionar archivo Excel
archivo_excel <- file.choose()
extension <- tolower(file_ext(archivo_excel))

# 2. Convertir .xlsb a .xlsx si es necesario
if (extension == "xlsb") {
  cat("Convirtiendo .xlsb a .xlsx con LibreOffice...\n")
  archivo_xlsx <- sub("\\.xlsb$", ".xlsx", archivo_excel, ignore.case = TRUE)
  comando <- sprintf('soffice --headless --convert-to xlsx "%s" --outdir "%s"', archivo_excel, dirname(archivo_excel))
  system(comando, intern = TRUE)
  archivo_excel <- archivo_xlsx
  extension <- "xlsx"
  cat("Conversión terminada →", archivo_excel, "\n")
}

# 3. Definir columnas esperadas (según estructura del Excel)
columnas_esperadas <- c(
  "IdInfoBase", "InfoBase", "idcierrexea", "IdBuque", "BUQUE", "CodigoCFR",
  "EsloraTotal", "EsloraPP", "ArqueoGT", "FcBaja", "FechaCambioDatosTécnicos",
  "FechaCambioDatosMotor", "PotenciaPropulsionTotalKW", "FcCambioIdentificacion",
  "PuertoBase", "AL5_PuertoBase", "ProvinciaPuertoBase", "CensoPorModalidad",
  "FcCambioCensoxModalidad"
)

# 4. Leer datos desde Excel
if (extension %in% c("xls", "xlsx")) {
  hoja <- excel_sheets(archivo_excel)[1]  # Asume primera hoja
  raw <- read_excel(archivo_excel, sheet = hoja)
  names(raw) <- make.names(names(raw), unique = TRUE)
  
  df <- as.data.frame(matrix(NA, nrow = nrow(raw), ncol = length(columnas_esperadas)))
  colnames(df) <- columnas_esperadas
  
  for (col in columnas_esperadas) {
    idx <- which(tolower(make.names(names(raw))) == tolower(make.names(col)))
    if (length(idx) == 1) df[[col]] <- raw[[idx]]
  }
} else {
  stop("Archivo no soportado. Usa .xls o .xlsx")
}

# 5. Convertir fechas correctamente
fechas <- c("FcCambioCensoxModalidad", "FcBaja", "FechaCambioDatosTécnicos", 
            "FechaCambioDatosMotor", "FcCambioIdentificacion")
for (fc in fechas) {
  if (fc %in% colnames(df)) {
    df[[fc]] <- as.Date(df[[fc]])
  }
}

# 6. Agregar columna ID artificial
df$Id <- seq_len(nrow(df))
df <- df %>% select(Id, everything())

# 7. Renombrar columnas a los nombres reales de PostgreSQL
nombres_postgres <- c(
  "Id" = "id",
  "BUQUE" = "nombre_buque",
  "FcCambioCensoxModalidad" = "fccambiocensoxmodalidad",
  "IdInfoBase" = "idinfobase",
  "InfoBase" = "infobase",
  "idcierrexea" = "idcierrexea",
  "IdBuque" = "idbuque",
  "EsloraTotal" = "esloratotal",
  "EsloraPP" = "eslorapp",
  "ArqueoGT" = "arqueogt",
  "FcBaja" = "fcbaja",
  "FechaCambioDatosTécnicos" = "fechacambiodatostecnicos",
  "FechaCambioDatosMotor" = "fechacambiodatosmotor",
  "FcCambioIdentificacion" = "fechacambioidentifacion",
  "PotenciaPropulsionTotalKW" = "potenciapropulsiontotalkw",
  "PuertoBase" = "puertobase",
  "AL5_PuertoBase" = "al5_puertobase",
  "ProvinciaPuertoBase" = "provinciapuertobase",
  "CensoPorModalidad" = "censopormodalidad",
  "CodigoCFR" = "codigocfr_info_buques"
)

nombres_validos <- intersect(names(nombres_postgres), colnames(df))
colnames(df)[match(nombres_validos, colnames(df))] <- nombres_postgres[nombres_validos]

# 8. Conexión a PostgreSQL
con <- dbConnect(RPostgres::Postgres(),
                 dbname = "neondb",
                 host = "ep-plain-fire-aba8ctis.eu-west-2.aws.neon.tech",
                 port = 5432,
                 user = "editor",
                 password = "CECAF_editor",
                 sslmode = "require")

# 9. Verificar columnas reales de la tabla info_buques
columnas_db <- dbGetQuery(con, "
  SELECT column_name FROM information_schema.columns
  WHERE table_name = 'info_buques' AND table_schema = 'public'
")$column_name

df <- df[, intersect(colnames(df), columnas_db)]

# 10. Crear tabla temporal y subir los datos
dbWriteTable(con, name = "temp_info_buques", value = df, overwrite = TRUE, row.names = FALSE)

# 11. Insertar en tabla principal (usando columnas explícitas)
columnas_insert <- paste(colnames(df), collapse = ", ")
sql_insert <- sprintf("
  INSERT INTO info_buques (%s)
  SELECT %s FROM temp_info_buques 
  ON CONFLICT (codigocfr_info_buques) DO NOTHING;", 
                      columnas_insert, columnas_insert)
dbExecute(con, sql_insert)

# 12. Eliminar tabla temporal y cerrar conexión
dbExecute(con, "DROP TABLE temp_info_buques;")
dbDisconnect(con)

cat("✅ Datos insertados correctamente en info_buques.\n")
