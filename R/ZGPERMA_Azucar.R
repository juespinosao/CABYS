#' @export
# Caña_azucar


f_Caña_azucar<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]
  # Especifica la ruta del archivo de Excel
  Caña_azucar <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/EMMET/",archivo),
                       sheet = "COMPLETO")
  # Seleccionar solo las columnas que necesitas
  Caña_azucar_tabla <- Caña_azucar[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]
  Caña_azucar_tabla=Caña_azucar_tabla %>%
    group_by(anio,mes)%>%
    filter(Clase_CIIU4==1071)%>%
    summarise(suma=sum(ProduccionRealPond))

  Valor_Caña_azucar=as.numeric((tail(Caña_azucar_tabla$suma,1))/
                            (sum(tail(lag(Caña_azucar_tabla$suma,12),1)))*100-100 )



# Valor_anterior ----------------------------------------------------------
if(mes==1){
  anio_mes=(anio-1)

}else{
  anio_mes=anio
}

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Azucar","NOMBRE"]

  Caña_anterior <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Caña de azucar y panela/Azucar/",archivo),
                        sheet = paste0("Azúcar - Alcohol ",anio),startRow = 4)
  n_fila1=which(Caña_anterior=="Caña Molida (toneladas)",arr.ind = TRUE)[,"row"]
  n_fila2=which(Caña_anterior=="Acumulado",arr.ind = TRUE)[,"row"]

  n_col=which(grepl("Caña Molida",Caña_anterior))

  Valor_anterior=as.numeric(Caña_anterior[(n_fila1+1):(n_fila2-1),n_col])


  return(list(variacion = Valor_Caña_azucar, vector = Valor_anterior))
}
