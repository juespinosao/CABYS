# Huevos

directorio="C:/Users/Asus/Desktop/Dane/proyecto2/Automatizacion CABYS/Automatizacion/Formato_carpetas"
mes=7
anio=2023
f_Huevos<-function(directorio,mes,anio){
  
  
  library(readxl)
  library(dplyr)
  
  
  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Huevos <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENAVI/Produccion_mensual_",nombres_meses[mes],"_",anio,".xlsx"))
  dos_digitos <- anio %% 100
  n_fila=which(Huevos == "Producto",arr.ind = TRUE)[, "row"]
  fila_tabla=as.numeric(which(Huevos == "HUEVOS (millones de Unidades)",arr.ind = TRUE)[, "row"])
  Tabla=Huevos[fila_tabla:(fila_tabla+11),]
  
  
  
  columna=which(grepl(dos_digitos,Huevos[n_fila,]),arr.ind = TRUE)
  Valor_Huevos=as.data.frame(Tabla[,(columna-1):columna])
  Valor_Huevos=as.numeric(c(Valor_Huevos[,1],Valor_Huevos[,2]))*1000000
  return(Valor_Huevos[1:(12+mes)])
  
}
