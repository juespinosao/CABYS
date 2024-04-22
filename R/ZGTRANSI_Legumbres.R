#' @export
# Legumbres

f_Legumbres<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Legumbres","NOMBRE"]

  Legumbres <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Mensualización_lejumbres/",archivo))
if(mes==1){
  fila=which(Legumbres$Año==anio)+(mes-1)
}else{
  fila=which(Legumbres$Año==anio)+(mes-1)
}



  vector=as.data.frame(Legumbres[1:fila,"Serie retropolada y mensualizada con r"])
  vector=as.numeric(vector$`Serie retropolada y mensualizada con r`)
  variacion=vector[fila]/tail(lag(vector,12),1)*100-100


  return(list(variacion = variacion, vector = vector))
}
