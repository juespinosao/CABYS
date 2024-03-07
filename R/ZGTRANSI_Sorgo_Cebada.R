#' @export
# Sorgo


f_Sorgo_Cebada<-function(directorio,mes,anio){

  #archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENALCE"))
  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)
  letra=ifelse(semestre==1,"A","B")
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]

  Sorgo <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENALCE/",archivo),
                    sheet = "Historico APR",startRow = 5)

  Sorgo_valor <- Sorgo %>%
    filter(AÑO == anio | AÑO == (anio-1), grepl("Sorgo", PRODUCTO)) %>%
    group_by(AÑO,SEMESTRE) %>%
    summarize(suma_produccion = sum(PRODUCCIÓN))%>%
    arrange(AÑO,SEMESTRE)

  Cebada_valor <- Sorgo %>%
    filter(AÑO == anio | AÑO == (anio-1), grepl("Cebada", PRODUCTO)) %>%
    group_by(AÑO,SEMESTRE) %>%
    summarize(suma_produccion = sum(PRODUCCIÓN))%>%
    arrange(AÑO)
  Sorgo_2015=5945
  Cebada_2015=3190.85
  Sorgo_indice2015=Sorgo_valor[3]/Sorgo_2015*100
  Cebada_indice2015=Cebada_valor[3]/Cebada_2015*100
  Sorgo_Participacion=23
  Cebada_Participacion=16
  Sorgo_final=Sorgo_indice2015/2*Sorgo_Participacion/100
  Cebada_final=Cebada_indice2015/2*Cebada_Participacion/100
  indicador_final=Sorgo_final+Cebada_final


  return(indicador_final$suma_produccion/2)
}
