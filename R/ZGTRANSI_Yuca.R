#' @export
# Yuca


f_Yuca<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]

  Yuca <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/",archivo))

  fila1=min(which(Yuca[,78]==2013,arr.ind = TRUE)[,"row"])
  fila2=min(which(Yuca[,78]==anio,arr.ind = TRUE)[,"row"])


  Valor_Yuca=as.data.frame(na.omit(Yuca[fila1:(fila2+mes-1),"Yuca retropolado...83"]))


  return(as.numeric(Valor_Yuca$Yuca))
}
