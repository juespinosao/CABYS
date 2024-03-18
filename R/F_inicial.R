#' Función inicial
#'
#' Funcion que instala las librerias necesarias para todo el proceso y crea las carpetas en donde se guardaran los archivos.
#'
#' @param directorio definir el directorio donde se crearan las carpetas.
#'
#' @details Se revisa si las funciones están instaladas en el entorno, en caso de
#' que no estén instaladas se procederán a instalar, luego se procede a revisar
#' si están creadas las carpetas donde se guardarán los archivos creados con la librería
#' en total se crearán siete carpetas, para cada una de las funciones que genera algún archivo de salida.
#'
#' 1 S1_integracion:\code{\link{f1_integracion}}
#'
#' 2 S2_estandarizacion:\code{\link{f2_estandarizacion}}
#'
#' 3 S3_identificacion_alertas:\code{\link{f3_identificacion_alertas}}
#'
#' 4 S4_imputacion:\code{\link{f4_imputacion}}
#'
#' 5 S5_tematica:\code{\link{f5_tematica}}
#'
#' 6 S6_anexos:\code{\link{f6_anacional}} y \code{\link{f7_aterritorial}}
#'
#' 7 S7_boletin:\code{\link{f8_boletin}}
#'
#' @examples f0_inicial(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET")
#'
#' @export


f0_inicial<-function(directorio,mes,anio){

  #instalar todas las librerias necesarias para el proceso

  # Lista de librerías que deseas instalar o cargar
  librerias <- c("tidyverse", "ggplot2", "dplyr","readr","readxl","openxlsx","rmarkdown","roxygen2","gt",
                 "gtExtras")

  # Verificar si las librerías están instaladas
  librerias_faltantes <- librerias[!sapply(librerias, requireNamespace, quietly = TRUE)]

  # Instalar librerías faltantes
  if (length(librerias_faltantes) > 0) {
    install.packages(librerias_faltantes)
  }

  # Cargar todas las librerías
  lapply(librerias, require, character.only = TRUE)



  #crear la función que revisa si la carpeta existe, de lo contrario la crea
  crearCarpeta <- function(ruta) {

    if (!dir.exists(ruta)) {
      dir.create(ruta)
      mensaje <- paste("Se ha creado la carpeta:", ruta)
      print(mensaje)
    } else {
      mensaje <- paste("La carpeta", ruta, "ya existe.")
      print(mensaje)
    }
  }
  carpeta=nombre_carpeta(mes,anio)

  #crear la carpeta results
  ruta=paste0(directorio,"/ISE/",anio,"/",carpeta)
  crearCarpeta(ruta)

  #crear la carpeta results
  ruta=paste0(directorio,"/ISE/",anio,"/",carpeta,"/Results")
  crearCarpeta(ruta)

  #crear la carpeta de Doc
  ruta=paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc")
  crearCarpeta(ruta)

  #crear la carpeta de Data
  ruta=paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data")
  crearCarpeta(ruta)

  #crear la carpeta coyuntura
  ruta=paste0(directorio,"/ISE/",anio,"/",carpeta,"/Results/Coyuntura")
  crearCarpeta(ruta)

  carpeta_actual=nombre_carpeta(mes,anio)
  entrada=paste0(directorio,"/Doc/Nombres_archivos_general.xlsx")
  salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx")

  # Cargar el archivo de entrada
  wb <- loadWorkbook(entrada)
  writeData(wb, sheet = "Nombres", x = mes,colNames = FALSE,startCol = "E", startRow = 2)
  writeData(wb, sheet = "Nombres", x = anio,colNames = FALSE,startCol = "F", startRow = 2)

  # Guardar el libro --------------------------------------------------------


  if (!file.exists(salida)) {
    saveWorkbook(wb, file = salida)
  } else {
    saveWorkbook(wb, file = salida,overwrite= TRUE)
  }

  print(paste0("Se creo el archivo Nombres_archivos_",nombres_meses[mes]," en ",directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc"))
  }
