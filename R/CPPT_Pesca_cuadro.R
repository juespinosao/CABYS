#' @export
# Pesca
# Cargar la biblioteca readxl

f_Pesca_complemento<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)




  #precio interno
  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")

  #EMMET
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]

  Pesca <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/EMMET/",archivo),
                     sheet = "COMPLETO")
  # Seleccionar solo las columnas que necesitas
  Pesca_tabla <- Pesca[, c("anio", "mes", "EMMET_Clase", "ProduccionRealPond")]
  Pesca_tabla=Pesca_tabla %>%
    filter(EMMET_Clase==1012 )%>%
    group_by(anio,mes)%>%
    filter(anio>(anio-3)) %>%
    summarise(suma=sum(ProduccionRealPond))%>%
    as.data.frame()
  tamaño=36+mes
  Pesca_tabla=tail(Pesca_tabla$suma,tamaño)
  var_anual=Pesca_tabla/lag(Pesca_tabla,12)*100-100
  Pesca_tabla_ant=lag(Pesca_tabla,12)
  tamaño=length(Pesca_tabla)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Pesca_tabla[(i-2):i]) / sum(Pesca_tabla_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Pesca_tabla[(i-11):i]) / sum(Pesca_tabla_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_EMMET=data.frame(Estado[c(24+mes)],
                          Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])

  #Exportaciones
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Pesca <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                     sheet = "PNK")
  n_fila=which(Pesca == "040101",arr.ind = TRUE)[,"row"]
  n_col1=which(Pesca== paste0(anio-2," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Pesca== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

  #Tomar el valor que nos interesa
  exportaciones=as.numeric(Pesca[n_fila[1],(n_col1[1]:n_col2[1])])
  var_anual=exportaciones/lag(exportaciones,12)*100-100
  exportaciones_ant=lag(exportaciones,12)
  tamaño=length(exportaciones)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(exportaciones[(i-2):i]) / sum(exportaciones_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(exportaciones[(i-11):i]) / sum(exportaciones_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_exportaciones=data.frame(Estado[c(12+mes)],
                                  Estado[c(24+mes)],Observaciones[c(12+mes)],Observaciones[c(24+mes)])


  colnames(cuadro_exportaciones)=colnames(cuadro_EMMET)
  nuevos_datos=bind_rows(cuadro_EMMET,cuadro_exportaciones)
  return(nuevos_datos)
}




