#' @export
# Panela


f_Panela<-function(directorio,mes,anio){


library(readxl)
library(dplyr)
library(zoo)



#Crear el nombre de las carpetas del mes anterior y el actual
if(mes==1){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
}else{
  carpeta_anterior=nombre_carpeta(mes-1,anio)
}
carpeta=nombre_carpeta(mes,anio)


nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Panela","NOMBRE"]

Panela<-read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Caña de azucar y panela/Panela/",archivo),colNames = FALSE)

n_fila=which(grepl("AREAS",as.data.frame(t(Panela))))
Panela[n_fila, ] <- na.locf0(Panela[n_fila,])


columna1=which(grepl((anio),Panela),arr.ind = TRUE)
columna2=which(grepl("Produccion",Panela),arr.ind = TRUE)
columnaf_act <- intersect(columna1, columna2)

n_fila=which(Panela=="Total general",arr.ind = TRUE)[,"row"]

Valor_actual=as.numeric(Panela[n_fila[[1]],columnaf_act])


Panela_Historico<- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Data/consolidado_ISE/Caña de azucar y panela/Panela/Historico_panela_",nombres_meses[mes-1], "_",anio,".xlsx"))
fila_ant=which(Panela_Historico==(anio-1),arr.ind = TRUE)[,"row"]
preliminar_actual=Panela_Historico[fila_ant,3]*(1+(Valor_actual/Panela_Historico[fila_ant,2]*100-100)/100)
fila_año=which(Panela_Historico==anio,arr.ind = TRUE)[,"row"]
if(length(fila_año)==0){
  nuevos_datos=c(anio,Valor_actual,preliminar_actual)
  Panela_Historico=rbind(Panela_Historico,nuevos_datos)
}else{
  Panela_Historico[fila_año,]=nuevos_datos
}
write.xlsx(Panela_Historico,paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Caña de azucar y panela/Panela/Historico_panela_",nombres_meses[mes], "_",anio,".xlsx"))
return(preliminar_actual)
}
