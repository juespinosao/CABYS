# Arroz


f_Arroz<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Arroz <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Arroz/",toupper(nombres_siglas[mes+1]),anio," COMRAS ENE2022 A ",toupper(nombres_siglas[mes]),anio,".xlsx"))


  columna=which(grepl("No.Tn",Arroz),arr.ind = TRUE)
  filas=which(Arroz== (anio-1) | Arroz== anio,arr.ind = TRUE)[,"row"]

  Valor_Arroz=as.data.frame(Arroz[filas,columna])

  return(as.numeric(Valor_Arroz$...3))
}
