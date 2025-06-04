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
