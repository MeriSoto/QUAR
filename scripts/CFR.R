#-----------------------------------------------------------------------------#
#            SISTEMA DE AUTOCOMPLETADO DE CFR PARA BUQUES                     #
#-----------------------------------------------------------------------------#
# Autor: Carlos García Moya
# Fecha: 11/06/2025
# Descripción: Este script contiene un conjunto de funciones para buscar
#              códigos CFR de buques y actualizar archivos de datos.
#-----------------------------------------------------------------------------#

# Cargar librerías necesarias
library(jsonlite)
library(readxl)
library(writexl)
library(openxlsx) # Usado en una de las versiones de la función de Excel


#--------------------------------------#
# FUNCION: verificar_entorno_python
#--------------------------------------#
#' @title Comprueba y configura el entorno de Python.
#' @description Verifica si Python está en el PATH y si las librerías necesarias
#'              (selenium, webdriver-manager) están instaladas. Si no lo están,
#'              intenta instalarlas automáticamente usando pip.
#' @return Imprime mensajes de estado en la consola. Se detiene con un error si
#'         algo crítico falla (Python no encontrado, fallo de instalación).
verificar_entorno_python <- function() {
  cat("--- Verificando entorno de Python ---\n")
  
  # 1. Verificar si Python está disponible
  python_ok <- suppressWarnings(try(
    system2("python", args = "--version", stdout = TRUE, stderr = TRUE),
    silent = TRUE
  ))
  
  if (inherits(python_ok, "try-error")) {
    stop(
      "❌ Python no está instalado o no se encuentra en el PATH del sistema.\n",
      "   Solución:\n",
      "   1. Descárgalo desde: https://www.python.org/downloads/\n",
      "   2. Asegúrate de marcar la casilla 'Add Python to PATH' durante la instalación.\n",
      "   3. Reinicia R/RStudio y vuelve a intentarlo."
    )
  } else {
    cat("✔️ Python detectado:", python_ok[1], "\n")
  }
  
  # 2. Verificar e instalar librerías necesarias
  librerias <- list(
    list(pip = "selenium", import = "from selenium import webdriver"),
    list(pip = "webdriver-manager", import = "from webdriver_manager.chrome import ChromeDriverManager")
  )
  
  for (lib in librerias) {
    # Comando para intentar importar la librería en Python
    import_code <- sprintf('"%s"', lib$import)
    comprobacion <- suppressWarnings(system2("python", c("-c", import_code), stderr = TRUE, stdout = TRUE))
    
    # Si `comprobacion` tiene longitud > 0, es que hubo un error (ModuleNotFoundError)
    if (length(comprobacion) > 0 && any(grepl("ModuleNotFoundError", comprobacion))) {
      cat(paste0("⚠️ Librería '", lib$pip, "' no encontrada. Intentando instalar...\n"))
      
      # Intentar instalar con pip
      install_result <- system2("pip", args = c("install", lib$pip), stdout = TRUE, stderr = TRUE)
      
      # Comprobar si la instalación falló
      if (!is.null(attr(install_result, "status")) && attr(install_result, "status") != 0) {
        stop(
          paste0("❌ ERROR: No se pudo instalar la librería '", lib$pip, "'.\n"),
          "   Causa probable: pip no está instalado o no está en el PATH.\n",
          "   Solución: Intenta instalarla manualmente desde la terminal con 'pip install ", lib$pip, "'"
        )
      } else {
        cat(paste0("✔️ Librería '", lib$pip, "' instalada correctamente.\n"))
      }
    } else {
      cat(paste0("✔️ Librería '", lib$pip, "' ya está instalada.\n"))
    }
  }
  cat("--- Verificación completada con éxito ---\n")
}


#-------------------------#
# FUNCION: obtener_cfr
#-------------------------#
#' @title Obtiene el CFR para un nombre de buque.
#' @description Llama a un script Python externo (`buscar_cfr.py`) que realiza la búsqueda,
#'              procesa el JSON devuelto y aplica una lógica para decidir el CFR más adecuado.
#' @param nombre_buque El nombre del buque a buscar.
#' @return Un string con el CFR, o un mensaje explicativo.
obtener_cfr <- function(nombre_buque) {
  nombre_formateado <- toupper(gsub("_", " ", nombre_buque))
  
  # Llamada al script de Python
  resultado_raw <- system2("python", args = c("buscar_cfr.py", shQuote(nombre_formateado)), stdout = TRUE, stderr = TRUE)
  
  # Comprobar si el script Python falló
  estado_salida <- attr(resultado_raw, "status")
  if (!is.null(estado_salida) && estado_salida != 0) {
    warning("El script de Python devolvió un error para el buque: ", nombre_buque, ". Salida: ", paste(resultado_raw, collapse="\n"))
    return("error en python, revisar consola")
  }
  
  # Parsear JSON
  json_str <- paste(resultado_raw, collapse = "")
  datos <- tryCatch(fromJSON(json_str), error = function(e) NULL)
  
  if (is.null(datos) || length(datos) == 0) {
    return("Barco no encontrado")
  }
  
  df <- if (!is.data.frame(datos)) as.data.frame(do.call(rbind, datos), stringsAsFactors = FALSE) else datos
  if (nrow(df) == 0) return("Barco no encontrado")
  
  colnames(df) <- c("cfr", "estado")
  
  if (all(df$cfr == "error")) return("Barco no encontrado")
  
  if (nrow(df) == 1) {
    return(if (df$cfr[1] != "-") df$cfr[1] else "CFR no encontrado")
  }
  
  # Filtrar por buques activos
  activos <- df[!grepl("baja", tolower(df$estado)) & df$cfr != "-", ]
  
  if (nrow(activos) == 1) {
    return(activos$cfr)
  } else if (nrow(activos) > 1) {
    return("consultar manualmente")
  } else {
    return("CFR no encontrado")
  }
}


#-----------------------------#
# FUNCION: procesar_buques
#-----------------------------#
#' @title Procesa una lista de buques desde un archivo .txt o .csv.
#' @param input_path Ruta al archivo de entrada.
#' @return Un data.frame con las columnas "Buque" y "CFR".
procesar_buques <- function(input_path) {
  if (grepl("\\.csv$", input_path, ignore.case = TRUE)) {
    df <- read.csv(input_path, stringsAsFactors = FALSE, header = TRUE)
    colnames(df)[1] <- "Buque"
  } else if (grepl("\\.txt$", input_path, ignore.case = TRUE)) {
    buques <- readLines(input_path)
    df <- data.frame(Buque = buques, stringsAsFactors = FALSE)
  } else {
    stop("Formato de archivo no soportado. Debe ser .csv o .txt")
  }
  
  df$Buque <- toupper(gsub("_", " ", df$Buque))
  df$CFR <- sapply(df$Buque, obtener_cfr)
  
  return(df)
}

#--------------------------------------#
# FUNCION: completar_cfr_y_modificar_excel
#--------------------------------------#
#' @title Completa Códigos CFR en un archivo Excel y lo modifica directamente.
#' @description Lee un archivo Excel, busca buques sin Código CFR, los completa
#'              y guarda los cambios en el archivo original, preservando otras hojas.
#' @param ruta_excel La ruta al archivo Excel a modificar.
#' @param hoja El nombre o número de la hoja a procesar.
#' @param guardar Lógico. Si es TRUE, el archivo se sobrescribe.
#' @return Una lista con todos los data frames del libro de Excel (modificados).
#' @warning ¡Esta función sobrescribe el archivo original!
completar_cfr_y_modificar_excel <- function(ruta_excel, hoja = 1, guardar = TRUE) {
  if (!file.exists(ruta_excel)) {
    stop("❌ Error: El archivo especificado no existe en la ruta: ", ruta_excel)
  }
  
  nombres_hojas <- excel_sheets(ruta_excel)
  
  nombre_hoja_a_modificar <- if (is.numeric(hoja)) {
    if (hoja > length(nombres_hojas) || hoja < 1) stop("❌ Error: El número de hoja es inválido.")
    nombres_hojas[hoja]
  } else {
    if (!hoja %in% nombres_hojas) stop("❌ Error: La hoja '", hoja, "' no se encuentra en el archivo.")
    hoja
  }
  
  todas_las_hojas <- lapply(setNames(nombres_hojas, nombres_hojas), function(nombre) {
    read_excel(ruta_excel, sheet = nombre)
  })
  
  df <- todas_las_hojas[[nombre_hoja_a_modificar]]
  
  colnames_norm <- tolower(gsub("[ _]", "", colnames(df)))
  nombre_buque_idx <- which(colnames_norm %in% c("buque", "nombrebuque"))
  codigo_cfr_idx  <- which(colnames_norm %in% c("codigocfr", "codigo", "cfr", "codigocfrbuque"))
  
  if (length(nombre_buque_idx) == 0 || length(codigo_cfr_idx) == 0) {
    stop("❌ Error: La hoja '", nombre_hoja_a_modificar, "' debe contener columnas para 'Buque' y 'CodigoCfr'.")
  }
  
  nombre_original_buque <- colnames(df)[nombre_buque_idx[1]]
  nombre_original_cfr <- colnames(df)[codigo_cfr_idx[1]]
  
  # Usar nombres temporales estandarizados
  colnames(df)[nombre_buque_idx[1]] <- "TempBuque"
  colnames(df)[codigo_cfr_idx[1]]   <- "TempCFR"
  
  # Convertir la columna de CFR a caracter para evitar problemas con tipos de datos
  df$TempCFR <- as.character(df$TempCFR)
  
  indices_incompletos <- which(is.na(df$TempCFR) | trimws(df$TempCFR) == "")
  
  if (length(indices_incompletos) == 0) {
    cat("✅ Todos los buques ya tienen un CódigoCfr en la hoja '", nombre_hoja_a_modificar, "'. No hay nada que completar.\n")
    return(todas_las_hojas)
  }
  
  cat(paste("ℹ️ Se encontraron", length(indices_incompletos), "buques sin CFR para procesar.\n"))
  
  cache_cfr <- list()
  
  for (i in indices_incompletos) {
    nombre_buque_original <- df$TempBuque[i]
    if (is.na(nombre_buque_original) || trimws(nombre_buque_original) == "") next # Saltar si el nombre del buque está vacío
    
    nombre_buque_proc <- toupper(gsub("_", " ", nombre_buque_original))
    
    if (!is.null(cache_cfr[[nombre_buque_proc]])) {
      df$TempCFR[i] <- cache_cfr[[nombre_buque_proc]]
      cat(sprintf("♻️ Reutilizado CFR: %s --> %s\n", nombre_buque_proc, df$TempCFR[i]))
    } else {
      resultado_cfr <- obtener_cfr(nombre_buque_proc)
      df$TempCFR[i] <- resultado_cfr
      cache_cfr[[nombre_buque_proc]] <- resultado_cfr
      cat(sprintf("🔍 Procesado buque: %s --> CFR: %s\n", nombre_buque_proc, resultado_cfr))
    }
  }
  
  # Restaurar nombres de columna originales
  colnames(df)[colnames(df) == "TempBuque"] <- nombre_original_buque
  colnames(df)[colnames(df) == "TempCFR"]   <- nombre_original_cfr
  
  todas_las_hojas[[nombre_hoja_a_modificar]] <- df
  
  if (guardar) {
    write_xlsx(todas_las_hojas, path = ruta_excel)
    cat("\n📁 ¡Éxito! El archivo original ha sido modificado y guardado en:", ruta_excel, "\n")
  } else {
    cat("\nℹ️ Proceso completado. 'guardar' es FALSE, el archivo no fue modificado en disco.\n")
  }
  
  return(todas_las_hojas)
}