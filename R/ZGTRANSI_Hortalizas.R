#' @export
# Hortalizas


f_Hortalizas<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Hortalizas <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/Datos_SIPSA/Base_EB_SIPSA.xlsx"))
  fila1=which(Hortalizas==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Hortalizas==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_Hortalizas=as.data.frame(na.omit(Hortalizas[1:fila_f,"Hortalizas"]))

  return(as.numeric(Valor_Hortalizas$Hortalizas))
}
