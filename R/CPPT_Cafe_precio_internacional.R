#' @export
# Cafe_verde
# Cargar la biblioteca readxl

f_Cafe_precio_internacional_ppt<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)




  #precio interno
  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Precio_Cafe","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Precio_internacional <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                               sheet = "Monthly Prices")

  columna1=which(grepl("Arabica",Precio_internacional),arr.ind = TRUE)


  tamaño=36+mes

  Precio=as.data.frame(tail(Precio_internacional[,columna1[1]],tamaño))
  Precio=as.numeric(Precio[,1])
  var_anual=Precio/lag(Precio,12)*100-100
  Precio_ant=lag(Precio,12)
  tamaño=length(Precio)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Precio[(i-2):i]) / sum(Precio_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Precio[(i-11):i]) / sum(Precio_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
Observaciones=as.numeric(Observaciones)

  cuadro_precio=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                           Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])
  return(cuadro_precio)
}
