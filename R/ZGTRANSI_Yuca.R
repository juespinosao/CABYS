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

  fila1=which(Yuca==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Yuca==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_Yuca=as.data.frame(na.omit(Yuca[1:fila_f,"Yuca"]))


  return(as.numeric(Valor_Yuca$Yuca))
}
