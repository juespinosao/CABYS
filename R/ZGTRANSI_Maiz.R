#' @export
# Maiz


f_Maiz<-function(directorio,mes,anio){

  #archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENALCE"))
  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)
  letra=ifelse(semestre==1,"A","B")
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="FENALCE","NOMBRE"]

  Maiz <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENALCE/",archivo),
                     sheet = "Historico APR",startRow = 2)

  Maiz_tabla <- Maiz %>%
                filter((AÑO == anio | AÑO == (anio - 1)| AÑO == (anio - 2)), grepl("Maíz", PRODUCTO)) %>%
                group_by(AÑO,SEMESTRE) %>%
                summarize(suma_produccion = sum(PRODUCCIÓN))%>%
                arrange(AÑO)

    variacion <- (Maiz_tabla$suma_produccion / lag(Maiz_tabla$suma_produccion,2) * 100) - 100

    variacion=variacion[!is.na(variacion)]
  return(as.numeric(variacion))
}
