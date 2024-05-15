#' @export
# Cafe_verde
# Cargar la biblioteca readxl

f_produccion_cafetos<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)




  #precio interno
  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafe","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Produccion <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                               sheet = "Produccion")

  columna1=which(grepl("Producción",Produccion),arr.ind = TRUE)


  tamaño=36+mes

  Produccion=as.data.frame(tail(na.omit(Produccion[,columna1[1]]),tamaño))
  Produccion=as.numeric(Produccion[,1])
  var_anual=Produccion/lag(Produccion,12)*100-100
  Produccion_ant=lag(Produccion,12)
  tamaño=length(Produccion)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Produccion[(i-2):i]) / sum(Produccion_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Produccion[(i-11):i]) / sum(Produccion_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
Observaciones=as.numeric(Observaciones)

  cuadro_Produccion=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                           Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])
  return(cuadro_Produccion)
}
