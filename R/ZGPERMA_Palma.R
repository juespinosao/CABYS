#' @export
# Palma
# Cargar la biblioteca readxl

f_Palma<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # Fruto de palma ------------------------------------------------------------------

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Palma","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Palma <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Palma/",archivo),colNames = FALSE)


  n_fila1=which(grepl(paste0("FRUTO.*",anio-1),as.data.frame(t(Palma))))
  n_fila2=which(grepl(paste0("FRUTO.*",anio),as.data.frame(t(Palma))))
  n_col=which(Palma== "ENERO",arr.ind = TRUE)[,"col"][[1]]



  valor_fruto=c(as.numeric(Palma[n_fila1+6,n_col:(n_col+11)]),as.numeric(Palma[n_fila2+6,n_col:(n_col+11)]))/1000




  # Aceite de palma ------------------------------------------------------------------


  n_fila1=which(grepl(paste0("ACEITE.*",anio-1),as.data.frame(t(Palma))))
  n_fila2=which(grepl(paste0("ACEITE.*",anio),as.data.frame(t(Palma))))

  valor_aceite=c(as.numeric(Palma[n_fila1+6,n_col:(n_col+11)]),as.numeric(Palma[n_fila2+6,n_col:(n_col+11)]))


  # Agrupar datos -----------------------------------------------------------


  return(list(fruto=valor_fruto,aceite=valor_aceite))
}
