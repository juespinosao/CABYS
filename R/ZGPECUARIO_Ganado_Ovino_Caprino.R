#' @export
# Ganado_Ovino y caprino
# Cargar la biblioteca readxl

f_Ovino_Caprino<-function(directorio,mes,anio){

#Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)

#Identificar la carpeta del mes actual
  carpeta=nombre_carpeta(mes,anio)

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="ESAG1","NOMBRE"]

# Especifica la ruta del archivo de Excel
  Ganado_Ovino<- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/ESAG/",archivo),
                               sheet = "Cuadro_7")
  Ganado_Caprino<- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/ESAG/",archivo),
                            sheet = "Cuadro_6")

#Identificar la fila donde esta la palabra Periodo
  n_fila=which(Ganado_Ovino == "Periodo",arr.ind = TRUE)[, "row"]

#En esa fila, reemplazar NA por el valor de la columna anterior
  Ganado_Ovino[n_fila, ] <- na.locf0(Ganado_Ovino[n_fila,])
  ###crear alerta de que cambia formato

#identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Total general",Ganado_Ovino),arr.ind = TRUE)
  columna2=which(grepl("Peso en pie",Ganado_Ovino),arr.ind = TRUE)
  columnaf <- intersect(columna1, columna2)

#Obtener el valor de los ultimos 3 meses
  fila_tabla=as.numeric(which(Ganado_Ovino == "Enero",arr.ind = TRUE)[, "row"])
  Tabla_Ovino=Ganado_Ovino[fila_tabla:(fila_tabla+(mes-1)),columnaf]


##Mismo proceso para el producto ganado caprino


#Identificar la fila donde esta la palabra Periodo
  n_fila=which(Ganado_Caprino == "Periodo",arr.ind = TRUE)[, "row"]

#En esa fila, reemplazar NA por el valor de la columna anterior
  Ganado_Caprino[n_fila, ] <- na.locf0(Ganado_Caprino[n_fila,])


#identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Total general",Ganado_Caprino),arr.ind = TRUE)
  columna2=which(grepl("Peso en pie",Ganado_Caprino),arr.ind = TRUE)
  columnaf <- intersect(columna1, columna2)

#Obtener el valor de los ultimos 3 meses
  fila_tabla=as.numeric(which(Ganado_Caprino == "Enero",arr.ind = TRUE)[, "row"])
  Tabla_Caprino=Ganado_Caprino[fila_tabla:(fila_tabla+(mes-1)),columnaf]


#Pegar los datos en una sola tabla y cambiarle los nombres de las columnas
  Tabla_datos=cbind(Tabla_Caprino,Tabla_Ovino)
  colnames(Tabla_datos)=c("Caprino","Ovino")

#Convertir las columnas en numericas
  Tabla_datos$Caprino=as.numeric(Tabla_datos$Caprino)
  Tabla_datos$Ovino=as.numeric(Tabla_datos$Ovino)

#Crear dato del promedio de 2015, necesario para el calculo
  promedio_2015=175020.58333

#Calcular la suma,indice e indice anual
  Tabla_datos=Tabla_datos%>%
    mutate(suma=Tabla_datos$Caprino+Tabla_datos$Ovino,indice=suma/promedio_2015*100,
           indice_anual=(103/12)*indice/100)
indice_trimestral=NULL
#Calcular el valor del indice trimestral, usado en el ZG

for (i in seq(3, nrow(Tabla_datos), by = 3)) {
  indice_trimestral$indice_trimestral[i]=as.numeric(sum(Tabla_datos$indice_anual[(i-2):i]))  # Realiza la suma y división
}

indice_final=na.omit(indice_trimestral$indice_trimestral)
return(indice_final)

}
