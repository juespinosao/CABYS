#' @export
# Algodon


f_Algodon<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)


  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Algodon","NOMBRE"]

  Algodon<-read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Algodón/",archivo),colNames = FALSE)

 if (semestre==1){
 n_fila1=which(grepl("COSTA",as.data.frame(t(Algodon))))
 n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
 n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
 Algodon_actual=Algodon[n_filaf,3]
 Algodon_pasado<-read.xlsx(paste0(directorio,"/ISE/",anio-1,"/","06Junio","/Data/consolidado_ISE/Algodón/INFORMACION_DANE_Julio_",anio-1,".xlsx"),colNames = FALSE)
 n_fila1=which(grepl("COSTA",as.data.frame(t(Algodon))))
 n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
 n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
 Algodon_anterior=Algodon_pasado[n_filaf,3]
 variacion=Algodon_actual/Algodon_anterior*100-100
 }else{
   n_fila1=which(grepl("INTERIOR",as.data.frame(t(Algodon))))
   n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
   n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
   Algodon_actual=Algodon[n_filaf,3]
   Algodon_pasado<-read.xlsx(paste0(directorio,"/ISE/",anio-1,"/","12Diciembre","/Data/consolidado_ISE/Algodón/INFORMACION_DANE_Enero_",anio,".xlsx"),colNames = FALSE)
   n_fila1=which(grepl("INTERIOR",as.data.frame(t(Algodon_pasado))))
   n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
   n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
   Algodon_anterior=Algodon_pasado[n_filaf,3]
   variacion=Algodon_actual/Algodon_anterior*100-100
 }


return(variacion)
}
