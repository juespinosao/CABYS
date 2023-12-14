# Algodon


f_Algodon<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  library(zoo)

  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)
  Algodon<-read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Algodón/INFORMACION DANE ",toupper(nombres_meses[mes+1])," ",anio,".xlsx"),colNames = FALSE)

 if (semestre==1){
 n_fila1=which(grepl("COSTA",as.data.frame(t(Algodon))))
 n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
 n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
 Algodon_actual=Algodon[n_filaf,3]
 Algodon_pasado<-read.xlsx(paste0(directorio,"/",anio-1,"/","06 Junio_ISE_",anio-1,"/consolidado_ISE/Algodón/INFORMACION DANE JULIO ",anio-1,".xlsx"),colNames = FALSE)
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
   Algodon_pasado<-read.xlsx(paste0(directorio,"/",anio-1,"/","12 Diciembre_ISE_",anio-1,"/consolidado_ISE/Algodón/INFORMACION DANE ENERO"," ",anio,".xlsx"),colNames = FALSE)
   n_fila1=which(grepl("INTERIOR",as.data.frame(t(Algodon_pasado))))
   n_fila2=which(Algodon=="PRODUCCION ALGODÓN CAMPO: TONELADAS",arr.ind = TRUE)[,"row"]
   n_filaf=n_fila2[min(which(n_fila2>n_fila1))]
   Algodon_anterior=Algodon_pasado[n_filaf,3]
   variacion=Algodon_actual/Algodon_anterior*100-100
 }


return(variacion)
}
