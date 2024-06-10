#' @export
# Pesca
# Cargar la biblioteca readxl

f_Resultados_trim<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)




  # Actualizacion mensual ---------------------------------------------------


  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Participaciones2","NOMBRE"]
  participaciones<- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/",archivo))



}
