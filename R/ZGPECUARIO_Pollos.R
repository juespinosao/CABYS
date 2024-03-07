#' @export
# Pollos


f_Pollos<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)


  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="FENAVI1","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Pollos <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENAVI/",archivo))
  dos_digitos <- anio %% 100
  n_fila=which(Pollos == "Producto",arr.ind = TRUE)[, "row"]
  fila_tabla=as.numeric(which(Pollos == "POLLO (Toneladas)",arr.ind = TRUE)[, "row"])
  Tabla=Pollos[fila_tabla:(fila_tabla+11),]

  columna=which(grepl(dos_digitos,Pollos[n_fila,]),arr.ind = TRUE)
  Valor_Pollos=as.data.frame(Tabla[,(columna-1):columna])
  Valor_Pollos=as.numeric(c(Valor_Pollos[,1],Valor_Pollos[,2]))
  return(Valor_Pollos[1:(12+mes)])
}
