#' @export
# Hortalizas


f_Hortalizas<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]

  Hortalizas <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/Microdatos desde 2013/",archivo))

  fila1=min(which(Hortalizas[,"year"]==(anio-2),arr.ind = TRUE)[,"row"])
  fila2=min(which(Hortalizas[,"year"]==anio,arr.ind = TRUE)[,"row"])


  Valor_Hortalizas=as.data.frame(na.omit(Hortalizas[fila1:(fila2+mes-1),"Hortalizas"]))

  return(as.numeric(Valor_Hortalizas[,1]))
}
