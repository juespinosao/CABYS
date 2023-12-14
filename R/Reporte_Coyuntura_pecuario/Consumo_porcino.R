# Consumo_porcino
# Cargar la biblioteca readxl

f_Consumo_porcino<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)

  # Especifica la ruta del archivo de Excel
  Consumo_porcino <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/ESAG/Destino CTAS2016_",anio,".xlsx"),
                              sheet = "PORCI",startRow = 2)



  ###crear alerta de que cambia formato

  #si which es 0 entonces generar error o algo

  #identificar las columna donde dice total general y peso en pie
  fila1=which(grepl(anio,Consumo_porcino$Año),arr.ind = TRUE)
  fila2=which(grepl(1,Consumo_porcino$Mes),arr.ind = TRUE)
  filaf <- intersect(fila1, fila2)





  #Tomar el valor que nos interesa
  Valor_porcino=as.data.frame(Consumo_porcino[filaf:(filaf+mes-1),c("Plazas.y.famas","Supermercados.de.cadena","Mercado.institucional")])


  Valor_porcino=as.data.frame(lapply(Valor_porcino, as.numeric))

  return(Valor_porcino)
}
