#' @export
# Hortalizas


f_Hortalizas<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]

  Hortalizas <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/",archivo))
  columna_fila=which(grepl("Hortalizas con exclusion plazas",Hortalizas),arr.ind = TRUE)
  columna=which(grepl("Hortalizas retropolado",Hortalizas),arr.ind = TRUE)
  fila1=min(which(Hortalizas[,columna_fila[1]-3]==2013,arr.ind = TRUE)[,"row"])
  fila2=min(which(Hortalizas[,columna_fila[1]-3]==anio,arr.ind = TRUE)[,"row"])


  Valor_Hortalizas=as.data.frame(na.omit(Hortalizas[fila1:(fila2+mes-1),columna[1]]))

  return(as.numeric(Valor_Hortalizas[,1]))
}
