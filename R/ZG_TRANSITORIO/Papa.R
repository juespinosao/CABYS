# Papa

directorio="C:/Users/Asus/Desktop/Dane/proyecto2/Automatizacion CABYS/Automatizacion/Formato_carpetas"
mes=7
anio=2023
f_Papa<-function(directorio,mes,anio){
  
  
  library(readxl)
  library(dplyr)
  library(zoo)
  
  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Papa <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Mensualización_papa/SeriesTd_",tolower(nombres_siglas[mes]),"_",anio,".xlsX"))
  
  return(as.numeric(Papa$V1))
}
