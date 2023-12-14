# Legumbres

f_Legumbres<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel
  Legumbres <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Mensualización_lejumbres/1.mensualizacion_legumbres a ",anio,".xlsX"))

  fila=which(Legumbres$Año==anio)+(mes-1)


  vector=as.data.frame(Legumbres[1:fila,"Serie retropolada y mensualizada con r"])
  vector=as.numeric(vector$`Serie retropolada y mensualizada con r`)
  variacion=vector[fila]/tail(lag(vector,12),1)*100-100


  return(list(variacion = variacion, vector = vector))
}
