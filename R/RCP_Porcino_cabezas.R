#' @export
# Ganado_Porcino_cabezas
# Cargar la biblioteca readxl

f_Porcino_cabezas<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="ESAG1","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Ganado_Porcino <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/ESAG/",archivo),
                              sheet = "Cuadro_3")


  #Identificar la fila donde esta la palabra Periodo
  n_fila=which(Ganado_Porcino == "Periodo",arr.ind = TRUE)[, "row"]

  #En esa fila, reemplazar NA por el valor de la columna anterior
  Ganado_Porcino[n_fila, ] <- na.locf0(Ganado_Porcino[n_fila,])


  ###crear alerta de que cambia formato

  #si which es 0 entonces generar error o algo

  #identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Total general",Ganado_Porcino),arr.ind = TRUE)
  columna2=which(grepl("Cabezas",Ganado_Porcino),arr.ind = TRUE)
  columnaf1 <- intersect(columna1, columna2)
  columna1=which(grepl("Machos",Ganado_Porcino),arr.ind = TRUE)
  columnaf2 <- intersect(columna1, columna2)
  columna1=which(grepl("Hembras",Ganado_Porcino),arr.ind = TRUE)
  columnaf3 <- intersect(columna1, columna2)
  columna1=which(grepl("Terneros",Ganado_Porcino),arr.ind = TRUE)
  columnaf4 <- intersect(columna1, columna2)
  columna1=which(grepl("Exportación",Ganado_Porcino),arr.ind = TRUE)
  columnaf5 <- intersect(columna1, columna2)

  fila=which(Ganado_Porcino=="Enero",arr.ind = TRUE)[,"row"]
  #Filtrar la fila del mes de interes



  #Tomar el valor que nos interesa
  Valor_Porcino=as.data.frame(Ganado_Porcino[fila:(fila+mes-1),c(columnaf1,columnaf2,columnaf3,columnaf4,columnaf5)])


  Valor_Porcino=as.data.frame(lapply(Valor_Porcino, as.numeric))

  return(Valor_Porcino)
}
