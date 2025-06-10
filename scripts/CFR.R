# Cargar la librería necesaria para manejar JSON
library(jsonlite)

#-------------------------#
# FUNCION: obtener_cfr
#-------------------------#
# Dada una cadena con el nombre de un buque, esta función:
# 1. Llama a un script Python externo que busca CFRs por nombre.
# 2. Procesa el JSON devuelto por Python.
# 3. Devuelve el CFR más adecuado o un mensaje explicativo.
obtener_cfr <- function(nombre_buque) {
  # Formatear nombre: convertir a mayúsculas y reemplazar guiones bajos por espacios
  nombre_buque <- toupper(gsub("_", " ", nombre_buque))
  
  # Llamar al script Python pasando el nombre del buque
  resultado <- system2("python", args = c("buscar_cfr.py", shQuote(nombre_buque)), stdout = TRUE, stderr = TRUE)

# Verifica si hubo un error (el atributo 'status' es clave)
estado_salida <- attr(resultado, "status")

# Imprime todo lo que devolvió el proceso para depurar
print("--- DEBUGGING PYTHON CALL ---")
print(paste("Nombre del buque pasado a Python:", nombre_buque))
print("Salida STDOUT (lo que debería ser JSON):")
print(resultado)
print("Salida STDERR (errores):")
# Stderr podría estar mezclado con stdout, lo filtramos si es necesario
# pero por ahora lo mostramos todo.
print(paste("Código de estado de salida:", estado_salida))
print("--- END DEBUGGING ---")


# Ahora, si el estado no es 0 (éxito), sabemos que falló.
if (!is.null(estado_salida) && estado_salida != 0) {
    return("error en python, revisar consola")
}
  
  # Unir líneas del output en una sola cadena y parsear JSON
  json_raw <- paste(resultado, collapse = "")
  datos <- tryCatch(fromJSON(json_raw), error = function(e) NULL)
  
  # Si no se obtuvo resultado válido, devolver "no encontrado"
  if (is.null(datos) || length(datos) == 0) {
    return("no encontrado")
  }
  
  # Convertir a data.frame si viene como lista
  if (!is.data.frame(datos)) {
    df <- as.data.frame(do.call(rbind, datos), stringsAsFactors = FALSE)
  } else {
    df <- datos
  }
  
  # Nombrar columnas para claridad
  colnames(df) <- c("cfr", "estado")
  
  # Si todos los CFR son "error", tratar como no encontrado
  if (all(df$cfr == "error")) {
    return("no encontrado")
  }
  
  # ⚠️ Lógica especial: si hay solo una entrada, devolver su CFR (salvo que sea "-")
  if (nrow(df) == 1) {
    return(if (df$cfr != "-") df$cfr else "CFR no encontrado")
  }
  
  # Filtrar resultados activos: que no contengan "baja" y que el CFR sea válido
  activos <- df[!grepl("baja", tolower(df$estado)) & df$cfr != "-", ]
  
  # Decidir qué devolver según resultados activos
  if (nrow(activos) == 1) {
    return(activos$cfr)  # Solo uno válido
  } else if (nrow(activos) > 1) {
    return("consultar manualmente")  # Ambigüedad
  } else {
    return("CFR no encontrado")  # Ninguno válido
  }
}


#-----------------------------#
# FUNCION: procesar_buques
#-----------------------------#
# Lee un archivo .txt o .csv con nombres de buques y devuelve un dataframe con su CFR.
procesar_buques <- function(input_path) {
  # Leer según el tipo de archivo
  if (grepl("\\.csv$", input_path, ignore.case = TRUE)) {
    df <- read.csv(input_path, stringsAsFactors = FALSE)
    colnames(df)[1] <- "Buque"  # Renombrar primera columna como 'Buque'
  } else if (grepl("\\.txt$", input_path, ignore.case = TRUE)) {
    buques <- readLines(input_path)
    df <- data.frame(Buque = buques, stringsAsFactors = FALSE)
  } else {
    stop("Formato de archivo no soportado (debe ser .csv o .txt)")
  }
  
  # Normalizar nombres de buques
  df$Buque <- toupper(gsub("_", " ", df$Buque))
  
  # Aplicar la función obtener_cfr a cada buque
  df$CFR <- sapply(df$Buque, obtener_cfr)
  
  return(df)
}

 # pip install selenium webdriver-manager
verificar_entorno_python <- function() {
  python_ok <- tryCatch(
    system2("python", args = "--version", stdout = TRUE, stderr = TRUE),
    error = function(e) return(NULL)
  )
  
  if (is.null(python_ok)) {
    stop(
      "❌ Python no está instalado o no está en el PATH.\n",
      "🔧 Descárgalo desde: https://www.python.org/downloads/\n",
      "🔁 Una vez instalado, reinicia R y vuelve a ejecutar esta función."
    )
  } else {
    cat("✔️ Python detectado:", python_ok, "\n")
  }
  
  librerias <- list(
    list(pip = "selenium", import = 'from selenium import webdriver'),
    list(pip = "webdriver-manager", import = 'from webdriver_manager.chrome import ChromeDriverManager')
  )
  
  for (lib in librerias) {
    # Envolver el código de importación en comillas dobles
    import_code <- sprintf('"%s"', lib$import)
    comprobacion <- system2("python", c("-c", import_code), stderr = TRUE, stdout = TRUE)
    
    if (length(comprobacion) > 0) {
      cat("⚠️ Instalando librería:", lib$pip, "...\n")
      install_result <- system2("pip", c("install", lib$pip), stderr = TRUE, stdout = TRUE)
      
      if (length(grep("ERROR", install_result, ignore.case = TRUE)) > 0) {
        stop(paste("❌ No se pudo instalar la librería:", lib$pip))
      }
    } else {
      cat("✔️ Librería", lib$pip, "ya instalada.\n")
    }
  }
}
