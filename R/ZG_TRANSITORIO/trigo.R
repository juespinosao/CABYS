# Trigo

directorio="C:/Users/Asus/Desktop/Dane/proyecto2/Automatizacion CABYS/Automatizacion/Formato_carpetas"
mes=7
anio=2023
f_Trigo<-function(directorio,mes,anio){

  #archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENALCE"))
  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)
  letra=ifelse(semestre==1,"A","B")
  # Especifica la ruta del archivo de Excel
  Trigo <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENALCE/APR- 2005 A ",anio,letra," - DANE.xlsx"),
                    sheet = "Historico APR",startRow = 5)

  Trigo_tabla <- Trigo %>%
    filter((AÑO == anio | AÑO == (anio - 1))| AÑO == (anio - 2), grepl("Trigo", PRODUCTO)) %>%
    group_by(AÑO,SEMESTRE) %>%
    summarize(suma_produccion = sum(PRODUCCIÓN))%>%
    arrange(AÑO)

    variacion <- na.omit((Trigo_tabla$suma_produccion / lag(Trigo_tabla$suma_produccion,2) * 100) - 100)

  return(as.numeric(variacion))
}
