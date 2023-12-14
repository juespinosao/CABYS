# Ganado_Bovino_cabezas
# Cargar la biblioteca readxl

f_Bovino_cabezas<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)

  # Especifica la ruta del archivo de Excel
  Ganado_Bovino <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/ESAG/censo-Sacrificio-total-nacional-",nombres_meses[mes],"-",anio,".xls"),
                              sheet = "Cuadro_1")


  #Identificar la fila donde esta la palabra Periodo
  n_fila=which(Ganado_Bovino == "Periodo",arr.ind = TRUE)[, "row"]

  #En esa fila, reemplazar NA por el valor de la columna anterior
  Ganado_Bovino[n_fila, ] <- na.locf0(Ganado_Bovino[n_fila,])


  ###crear alerta de que cambia formato

  #si which es 0 entonces generar error o algo

  #identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Total general",Ganado_Bovino),arr.ind = TRUE)
  columna2=which(grepl("Cabezas",Ganado_Bovino),arr.ind = TRUE)
  columnaf1 <- intersect(columna1, columna2)
  columna1=which(grepl("Machos",Ganado_Bovino),arr.ind = TRUE)
  columnaf2 <- intersect(columna1, columna2)
  columna1=which(grepl("Hembras",Ganado_Bovino),arr.ind = TRUE)
  columnaf3 <- intersect(columna1, columna2)
  columna1=which(grepl("Terneros",Ganado_Bovino),arr.ind = TRUE)
  columnaf4 <- intersect(columna1, columna2)
  columna1=which(grepl("Exportación",Ganado_Bovino),arr.ind = TRUE)
  columnaf5 <- intersect(columna1, columna2)

  fila=which(Ganado_Bovino=="Enero",arr.ind = TRUE)[,"row"]
  #Filtrar la fila del mes de interes



  #Tomar el valor que nos interesa
  Valor_Bovino=as.data.frame(Ganado_Bovino[fila:(fila+mes-1),c(columnaf1,columnaf2,columnaf3,columnaf4,columnaf5)])


  Valor_Bovino=as.data.frame(lapply(Valor_Bovino, as.numeric))

  return(Valor_Bovino)
}
