# Yuca


f_Yuca<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Yuca <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/Datos_SIPSA/Base_EB_SIPSA.xlsx"))

  fila1=which(Yuca==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Yuca==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_Yuca=as.data.frame(na.omit(Yuca[1:fila_f,"Yuca"]))


  return(as.numeric(Valor_Yuca$Yuca))
}
