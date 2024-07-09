#' @export
# Cafe_verde
# Cargar la biblioteca readxl

f_producto_maiz<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)


  carpeta=nombre_carpeta(mes,anio)
  semestre=f_semestre(mes)
  letra=ifelse(semestre==1,"A","B")
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="FENALCE","NOMBRE"]

  Maiz <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENALCE/",archivo),
                    sheet = "Historico APR",startRow = 2)

  Maiz_tabla <- Maiz %>%
    mutate(PRODUCTO = tolower(PRODUCTO)) %>%
    mutate(PRODUCTO = gsub("tecnificado", "", PRODUCTO)) %>%
    mutate(PRODUCTO = gsub("tradicional", "", PRODUCTO)) %>%
    filter((AÑO == anio | AÑO == (anio - 1)| AÑO == (anio - 2)| AÑO == (anio - 3)), grepl("maíz", PRODUCTO)) %>%
    group_by(AÑO,SEMESTRE,PRODUCTO) %>%
    summarize(suma_produccion = sum(PRODUCCIÓN))%>%
    arrange(AÑO) %>%
    as.data.frame()



  var_anual=Maiz_tabla$suma_produccion/lag(Maiz_tabla$suma_produccion,4)*100-100
   blanco_ant=(tail(lag(Maiz_tabla$suma_produccion,6),1)+tail(lag(Maiz_tabla$suma_produccion,4),1))/(tail(lag(Maiz_tabla$suma_produccion,8),1)+tail(lag(Maiz_tabla$suma_produccion,10),1))*100-100
  blanco_act=(tail(Maiz_tabla$suma_produccion,1)+tail(lag(Maiz_tabla$suma_produccion,2),1))/(tail(lag(Maiz_tabla$suma_produccion,6),1)+tail(lag(Maiz_tabla$suma_produccion,4),1))*100-100
  amarillo_ant=(tail(lag(Maiz_tabla$suma_produccion,7),1)+tail(lag(Maiz_tabla$suma_produccion,5),1))/(tail(lag(Maiz_tabla$suma_produccion,9),1)+tail(lag(Maiz_tabla$suma_produccion,11),1))*100-100
  amarillo_act=(tail(lag(Maiz_tabla$suma_produccion),1)+tail(lag(Maiz_tabla$suma_produccion,3),1))/(tail(lag(Maiz_tabla$suma_produccion,7),1)+tail(lag(Maiz_tabla$suma_produccion,5),1))*100-100
  matriz=matrix(c(tail(lag(var_anual,4),1),tail(var_anual,1),tail(lag(var_anual,4),1),tail(var_anual,1),blanco_ant,
                  blanco_act,tail(lag(var_anual,5),1),tail(lag(var_anual),1),tail(lag(var_anual,5),1),
                  tail(lag(var_anual),1),amarillo_ant,amarillo_act),nrow = 2,ncol = 6,byrow = TRUE)
matriz=as.data.frame(matriz)
  return(matriz)
}
