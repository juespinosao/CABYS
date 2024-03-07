#' @export
# Tabaco


f_Tabaco<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]

  Tabaco <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/EMMET/",archivo),
                       sheet = "COMPLETO")
  # Seleccionar solo las columnas que necesitas
  Tabaco_tabla <- Tabaco[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]
  Tabaco_tabla=Tabaco_tabla %>%
         filter(Clase_CIIU4==1200 )%>%
         group_by(anio,mes)%>%
         summarise(suma=sum(ProduccionRealPond))%>%
         as.data.frame()

  fila=which(Tabaco_tabla==(anio-1),arr.ind = TRUE)[,"row"]
  Tabaco_tabla$anterior=lag(Tabaco_tabla$suma,12)
  Tabaco_tabla$Estado <- ""

  for (i in seq(fila[1]+2, nrow(Tabaco_tabla), by = 3)) {
    if(sum(Tabaco_tabla$anterior[(i-2):i])==0){
      Tabaco_tabla$Estado[i]=0
    }else{
    Tabaco_tabla$Estado[i] <- (sum(Tabaco_tabla$suma[(i-2):i]) / sum(Tabaco_tabla$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  }
Valor_Tabaco=as.numeric(Tabaco_tabla[fila[1]:nrow(Tabaco_tabla),"Estado"])
Valor_Tabaco=Valor_Tabaco[!is.na(Valor_Tabaco)]
  return(Valor_Tabaco)
}
