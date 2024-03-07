#' @export
# Papa


f_Papa<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Papa","NOMBRE"]

  Papa <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Mensualización_papa/",archivo))

  return(as.numeric(Papa$V1))
}
