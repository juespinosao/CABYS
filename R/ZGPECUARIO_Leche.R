#' @export
# Leche


f_Leche<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)


# Leche_sipsa -------------------------------------------------------------

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Leche_SIPSA","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Leche <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Leche/SIPSA/",archivo),
                       sheet = "LecheDANE")

  Leche$Año=na.locf(Leche$Año)
  Valor_Leche=as.data.frame(Leche[Leche[,"Año"] == anio,"PRODUCCION LECHE CRUDA DANE"])




  return(as.numeric(Valor_Leche[,1]))
}
