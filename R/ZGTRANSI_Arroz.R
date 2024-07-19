#' @export
# Arroz


f_Arroz<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Arroz","NOMBRE"]

  Arroz <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Arroz/",archivo))


  columna=which(grepl("No. Tn",Arroz),arr.ind = TRUE)
  filas=which(Arroz==(anio-1) | Arroz== anio,arr.ind = TRUE)[,"row"]

  Valor_Arroz=as.data.frame(Arroz[filas,columna])

  return(as.numeric(Valor_Arroz[,1]))
}
