# install.packages(c("odbc", "DBI", "openxlsx")) # Ejecutar si no los tienes
library(DBI)
library(odbc)
library(openxlsx)

# Seleccionar archivo Access
access_file <- file.choose()

# Conectar a Access
con <- dbConnect(odbc::odbc(),
                 .connection_string = paste0(
                   "Driver={Microsoft Access Driver (*.mdb, *.accdb)};",
                   "Dbq=", access_file, ";"
                 ))

# --- Consulta 1: Desembarcados ---
query1 <- "
SELECT d.ID, db.CODIGO_BUQUE, b.NOMBRE AS Buque, a.nombre AS Arte,
       Format(d.FECHASALIDA, 'dd-mm-yyyy') AS FechaSalida,
       Format(d.FECHAREGRESO, 'dd-mm-yyyy') AS FechaRegreso,
       Format(d.FECHADESEMBARCO, 'dd-mm-yyyy') AS FechaDesembarco,
       d.PUERTO_DES, e.AL3, e.CIENTIFICO,
       Sum(de.PESO_VIVO) AS KgDesembarcados
FROM (((((INFODIARIOS AS d
  INNER JOIN artes AS a ON d.IDARTE = a.ID)
  INNER JOIN INFODIARIOSxBARCO AS db ON d.ID = db.IDDIARIO)
  INNER JOIN BARCOS AS b ON db.CODIGO_BUQUE = b.CODIGO_BUQUE)
  INNER JOIN NUMDIARIOS AS n ON d.ID = n.IDDIARIO)
  INNER JOIN DESEMBARQUES AS de ON n.ID = de.IDNUMDIARIO)
  INNER JOIN ESPECIES AS e ON de.IDESPECIE = e.ID
WHERE d.CECAF = 1 AND db.ORDENACION = 1
GROUP BY d.ID, db.CODIGO_BUQUE, b.NOMBRE, a.nombre,
         Format(d.FECHASALIDA, 'dd-mm-yyyy'),
         Format(d.FECHAREGRESO, 'dd-mm-yyyy'),
         Format(d.FECHADESEMBARCO, 'dd-mm-yyyy'),
         d.PUERTO_DES, e.AL3, e.CIENTIFICO
ORDER BY d.ID;
"

df1 <- dbGetQuery(con, query1)

# --- Consulta 2: Capturados ---
query2 <- "
SELECT d.ID, db.CODIGO_BUQUE, b.NOMBRE AS Buque, a.nombre AS Arte,
       Format(d.FECHASALIDA, 'dd-mm-yyyy') AS FechaSalida,
       Format(d.FECHAREGRESO, 'dd-mm-yyyy') AS FechaRegreso,
       Format(d.FECHADESEMBARCO, 'dd-mm-yyyy') AS FechaDesembarco,
       d.PUERTO_DES, Format(l.FECHA, 'dd-mm-yyyy') AS FechaCaptura,
       di.DIVISION, e.AL3, e.CIENTIFICO, c.PESO AS KgCapturados
FROM (((((((INFODIARIOS AS d
  INNER JOIN artes AS a ON d.IDARTE = a.ID)
  INNER JOIN DIARIOSxBARCO AS db ON d.ID = db.IDDIARIO)
  INNER JOIN BARCOS AS b ON db.CODIGO_BUQUE = b.CODIGO_BUQUE)
  INNER JOIN NUMDIARIOS AS n ON d.ID = n.IDDIARIO)
  INNER JOIN LINEAS AS l ON n.ID = l.IDNUMDIARIO)
  INNER JOIN CAPTURAS AS c ON l.ID = c.IDLINEA)
  INNER JOIN ESPECIES AS e ON c.IDESPECIE = e.ID)
  INNER JOIN GEN_DIVISIONES AS di ON l.IDDIVISION = di.ID
WHERE d.CECAF = 1 AND db.ORDENACION = 1
ORDER BY d.ID;
"

df2 <- dbGetQuery(con, query2)

# --- Consulta 3: Esfuerzo ---
query3 <- "
SELECT d.ID, db.CODIGO_BUQUE, b.NOMBRE AS Buque, a.nombre AS Arte,
       Format(d.FECHASALIDA, 'dd-mm-yyyy') AS FechaSalida,
       Format(d.FECHAREGRESO, 'dd-mm-yyyy') AS FechaRegreso,
       Format(d.FECHADESEMBARCO, 'dd-mm-yyyy') AS FechaDesembarco,
       d.PUERTO_DES,
       DateDiff('d', d.FECHASALIDA, Nz(d.FECHADESEMBARCO, d.FECHAREGRESO)) + 1 AS DiasMar,
       b.POTENCIA_KW
FROM (((INFODIARIOS AS d
  INNER JOIN artes AS a ON d.IDARTE = a.ID)
  INNER JOIN DIARIOSxBARCO AS db ON d.ID = db.IDDIARIO)
  INNER JOIN BARCOS AS b ON db.CODIGO_BUQUE = b.CODIGO_BUQUE)
WHERE d.CECAF = 1 AND db.ORDENACION = 1
ORDER BY d.ID;
"

df3 <- dbGetQuery(con, query3)

# --- Exportar a Excel ---
write.xlsx(df1, "salida_desembarcados.xlsx")
write.xlsx(df2, "salida_capturados.xlsx")
write.xlsx(df3, "salida_esfuerzo.xlsx")

# Cerrar conexión
dbDisconnect(con)
