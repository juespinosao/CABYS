#' @export

f_maiz_complemento<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")

archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]

archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Importaciones","NOMBRE"]

# Especifica la ruta del archivo de Excel
Impor <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                   sheet = "TOTAL IMPO_KTES")

n_col1=which(Impor== paste0((anio-3)," 01"),arr.ind = TRUE)[,"col"]
n_col2=which(Impor== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]
n_fila1=which(Impor== "010102",arr.ind = TRUE)[,"row"]


#Tomar el valor que nos interesa
importaciones=as.data.frame(t(Impor[c(n_fila1[1]),(n_col1[1]:n_col2[1])]))
importaciones=as.numeric(importaciones[,1])

var_anual=importaciones/lag(importaciones,12)*100-100
importaciones_ant=lag(importaciones,12)
tamaño=length(importaciones)
Estado <- rep("",tamaño)

for (i in seq(3, tamaño, by = 3)) {
  Estado[i] <- (sum(importaciones[(i-2):i]) / sum(importaciones_ant[(i-2):i]))*100-100  # Realiza la suma y división
}
Estado=as.numeric(Estado)
Observaciones <- rep("",tamaño)

for (i in seq(12, tamaño, by = 12)) {
  Observaciones[i] <- (sum(importaciones[(i-11):i]) / sum(importaciones_ant[(i-11):i]))*100-100  # Realiza la suma y división
}
Observaciones=as.numeric(Observaciones)

cuadro_importaciones=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                         Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])





nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPP","NOMBRE"]

IPP <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                     sheet = "1.1")

#Identificar la fila donde esta la palabra Periodo
n_col=which(IPP == "CODIGO",arr.ind = TRUE)[, "col"]

dos_digitos=anio %% 100
#identificar las columna donde dice total general y peso en pie
#identificar las columna donde dice total general y peso en pie
  columna1=max(which(grepl(paste0(nombres_siglas[1],"-",(dos_digitos-2)),IPP),arr.ind = TRUE))
  columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),IPP),arr.ind = TRUE)

fila1=which(grepl("01120",IPP[,n_col]),arr.ind = TRUE)


#Tomar el valor que nos interesa
Valor_IPP=as.data.frame(t(IPP[c(fila1[1]),c(columna1:columna2[1])]))
tamaño=nrow(Valor_IPP)
filas=c(seq(1, tamaño, by = 2),tamaño)
Valor_IPP=as.numeric(Valor_IPP[filas,1])

var_anual=Valor_IPP/lag(Valor_IPP,12)*100-100
Valor_IPP_ant=lag(Valor_IPP,12)
tamaño=length(Valor_IPP)
Estado <- rep("",tamaño)

for (i in seq(3, tamaño, by = 3)) {
  Estado[i] <- (sum(Valor_IPP[(i-2):i]) / sum(Valor_IPP_ant[(i-2):i]))*100-100  # Realiza la suma y división
}
Estado=as.numeric(Estado)
Observaciones <- rep("",tamaño)

for (i in seq(12, tamaño, by = 12)) {
  Observaciones[i] <- (sum(Valor_IPP[(i-11):i]) / sum(Valor_IPP_ant[(i-11):i]))*100-100  # Realiza la suma y división
}
Observaciones=as.numeric(Observaciones)

cuadro_Valor_IPP=data.frame(var_anual[c(12+mes)],var_anual[c(24+mes)],Estado[c(12+mes)],
                                Estado[c(24+mes)],Observaciones[c(12+mes)],Observaciones[c(24+mes)])


#precio internacional

carpeta=nombre_carpeta(mes,anio)
nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Precio_Cafe","NOMBRE"]

# Especifica la ruta del archivo de Excel
Precio_internacional <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                                  sheet = "Monthly Prices")

columna1=which(grepl("Maize",Precio_internacional),arr.ind = TRUE)


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

return(list(cuadro_importaciones,cuadro_Valor_IPP,cuadro_precio))
}
