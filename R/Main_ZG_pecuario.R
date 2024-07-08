#' @export

### pecuario

ZG_pecuario=function(directorio,mes,anio){

#Cargar librerias
library(openxlsx)
#utils

#Crear el nombre de las carpetas del mes anterior y el actual
if(mes==1){
carpeta_anterior=nombre_carpeta(12,(anio-1))
entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[12],"_",(anio-1),".xlsx")

}else{
carpeta_anterior=nombre_carpeta(mes-1,anio)
entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

}

carpeta_actual=nombre_carpeta(mes,anio)

#Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb <- loadWorkbook(entrada)


# Bovino ------------------------------------------------------------------

#Leer solo la hoja de Bovino
data <- read.xlsx(entrada, sheet = "Ganado_Bovino", colNames = TRUE,startRow = 10)


ultima_fila=nrow(data)

if(mes==1){
  fila=ultima_fila
}else{
  fila=which(data$Año== anio)
}


#Correr la funcion Bovino
valor_Bovino=f_Bovino(directorio,mes,anio)
valor_Bovino=valor_Bovino[,1]
valor_Bovino=as.data.frame(valor_Bovino)
valor_Bovino$anterior=tail(lag(data$Ganado.bovino.Kilos,11),mes)
valor_Bovino$Estado <- ""
if(nrow(valor_Bovino)>2){
for (i in seq(3, nrow(valor_Bovino), by = 3)) {
  valor_Bovino$Estado[i] <- (sum(valor_Bovino$valor_Bovino[(i-2):i]) / sum(valor_Bovino$anterior[(i-2):i]))*100-100  # Realiza la suma y división
}
}else{
valor_Bovino$Estado <- ""
}

if(mes==1){
  nuevos_datos <- data.frame(
    Consecutivo = c((data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion=rep("ESAG Sacrificio de ganado vacuno, peso kilo en pie",mes),
    Ganado.bovino.Kilos=valor_Bovino$valor_Bovino,
    Variacion.Anual=valor_Bovino$valor_Bovino/tail(lag(data$Ganado.bovino.Kilos,11),mes)*100-100,
    Estado=as.numeric(valor_Bovino$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),sum(valor_Bovino$valor_Bovino)/sum(valor_Bovino$anterior)*100-100)
    } else {
      c("")
    },
    Tipo=c("")
  )
}else{
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion=rep("ESAG Sacrificio de ganado vacuno, peso kilo en pie",mes),
    Ganado.bovino.Kilos=valor_Bovino$valor_Bovino,
    Variacion.Anual=valor_Bovino$valor_Bovino/tail(lag(data$Ganado.bovino.Kilos,11),mes)*100-100,
    Estado=as.numeric(valor_Bovino$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),sum(valor_Bovino$valor_Bovino)/sum(valor_Bovino$anterior)*100-100)
    } else {
      c(data[fila[1]:ultima_fila,"observaciones"],"")
    },
    Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
  )
}
#Crear la nueva fila

nuevos_datos$observaciones=as.numeric(nuevos_datos$observaciones)

# Escribe los datos en la hoja "Ganado_Bovino"
if(mes==1){
writeData(wb, sheet = "Ganado_Bovino", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+11))

}else{
writeData(wb, sheet = "Ganado_Bovino", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

}


addStyle(wb, sheet = "Ganado_Bovino",style=col1,rows = (ultima_fila+11),cols = 1:4,gridExpand = TRUE)
addStyle(wb, sheet = "Ganado_Bovino",style=col2,rows = (ultima_fila+11),cols = 5,gridExpand = TRUE)
addStyle(wb, sheet = "Ganado_Bovino",style=col3,rows = (ultima_fila+11),cols = 6,gridExpand = TRUE)
addStyle(wb, sheet = "Ganado_Bovino",style=col4,rows = (ultima_fila+11),cols = 7,gridExpand = TRUE)
addStyle(wb, sheet = "Ganado_Bovino",style=col10,rows = (ultima_fila+11),cols = 8,gridExpand = TRUE)

# Pollos ------------------------------------------------------------------

#Leer solo la hoja de Pollos
data <- read.xlsx(wb, sheet = "Pollos", colNames = TRUE,startRow = 10)

fila=which(data$Año==(anio-1))
ultima_fila=nrow(data)


#Correr la funcion Pollos


valor_Pollos=f_Pollos(directorio,mes,anio)
valor_Pollos=as.data.frame(valor_Pollos)
valor_Pollos$anterior=c(data[data$Año==(anio-2),"Pollos.Toneladas"],valor_Pollos[1:mes,"valor_Pollos"])
valor_Pollos$Estado <- ""

for (i in seq(3, nrow(valor_Pollos), by = 3)) {
  valor_Pollos$Estado[i] <- (sum(valor_Pollos$valor_Pollos[(i-2):i]) / sum(valor_Pollos$anterior[(i-2):i]))*100-100  # Realiza la suma y división
}
valor_Pollos$Observaciones <- ""

for (i in seq(12, nrow(valor_Pollos), by = 12)) {
  valor_Pollos$Observaciones[i] <- (sum(valor_Pollos$valor_Pollos[(i-11):i]) / sum(valor_Pollos$anterior[(i-11):i]))*100-100  # Realiza la suma y división
}
#Crear la nueva fila
nuevos_datos <- data.frame(
  Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
  Año = c(data[fila[1]:ultima_fila,"Año"],anio),
  Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
  Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Producción Fenavi"),
  Pollos.Toneladas=valor_Pollos$valor_Pollos,
  Variacion.Anual=valor_Pollos$valor_Pollos/tail(lag(data$Pollos.Toneladas,11),(12+mes))*100-100,
  Estado=as.numeric(valor_Pollos$Estado),
  observaciones=as.numeric(valor_Pollos$Observaciones),
  Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
)




# Escribe los datos en la hoja "Pollos"
writeData(wb, sheet = "Pollos", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


addStyle(wb, sheet = "Pollos",style=col1,rows = (ultima_fila+11),cols = 1:4)
addStyle(wb, sheet = "Pollos",style=col2,rows = (ultima_fila+11),cols = 5)
addStyle(wb, sheet = "Pollos",style=col3,rows = (ultima_fila+11),cols = 6)
addStyle(wb, sheet = "Pollos",style=col4,rows = (ultima_fila+11),cols = 7:8)
# Porcino ------------------------------------------------------------------

#Leer solo la hoja de Porcinos
data <- read.xlsx(wb, sheet = "Porcino", colNames = TRUE,startRow = 10)

ultima_fila=nrow(data)

if(mes==1){
  fila=ultima_fila
}else{
  fila=which(data$Año== anio)
}



#Correr la funcion Pollos
valor_Porcino=f_Porcino(directorio,mes,anio)
valor_Porcino=valor_Porcino[,1]
valor_Porcino=as.data.frame(valor_Porcino)
valor_Porcino$anterior=tail(lag(data$Porcino.Kilos,11),mes)
valor_Porcino$Estado <- ""

if(nrow(valor_Porcino)>2){
for (i in seq(3, nrow(valor_Porcino), by = 3)) {
  valor_Porcino$Estado[i] <- (sum(valor_Porcino$valor_Porcino[(i-2):i]) / sum(valor_Porcino$anterior[(i-2):i]))*100-100  # Realiza la suma y división
}
}else{
  valor_Porcino$Estado <- ""
}
if(mes==1){
  nuevos_datos <- data.frame(
    Consecutivo = c((data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion=rep("ESAG-DANE",mes),
    Porcino.Kilos=valor_Porcino$valor_Porcino,
    Variacion.Anual=valor_Porcino$valor_Porcino/tail(lag(data$Porcino.Kilos,11),mes)*100-100,
    Estado=as.numeric(valor_Porcino$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),as.numeric(sum(valor_Porcino$valor_Porcino)/sum(valor_Porcino$anterior)*100-100))
    } else {
      c("")
    },
    Tipo=c("")
  )
}else{
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion=rep("ESAG-DANE",mes),
    Porcino.Kilos=valor_Porcino$valor_Porcino,
    Variacion.Anual=valor_Porcino$valor_Porcino/tail(lag(data$Porcino.Kilos,11),mes)*100-100,
    Estado=as.numeric(valor_Porcino$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),as.numeric(sum(valor_Porcino$valor_Porcino)/sum(valor_Porcino$anterior)*100-100))
    } else {
      c(data[fila[1]:ultima_fila,"observaciones"],"")
    },
    Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
  )
}
#Crear la nueva fila




# Escribe los datos en la hoja "Porcino"
if(mes==1){
  writeData(wb, sheet = "Porcino", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+11))
}else{
writeData(wb, sheet = "Porcino", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))
}

addStyle(wb, sheet = "Porcino",style=col1,rows = (ultima_fila+11),cols = 1:4)
addStyle(wb, sheet = "Porcino",style=col2,rows = (ultima_fila+11),cols = 5)
addStyle(wb, sheet = "Porcino",style=col3,rows = (ultima_fila+11),cols = 6)
addStyle(wb, sheet = "Porcino",style=col4,rows = (ultima_fila+11),cols = 7:8)

# Leche ------------------------------------------------------------------

#Leer solo la hoja de Leche
data <- read.xlsx(wb, sheet = "Leche", colNames = TRUE,startRow = 10)

fila=which(data$Año== anio)
ultima_fila=nrow(data)



#Correr la funcion Pollos
valor_Leche=f_Leche(directorio,mes,anio)
valor_Leche=as.data.frame(valor_Leche)
valor_Leche$anterior=tail(lag(data$Leche.sin.elaborar.Volumen,11),mes)
valor_Leche$Estado <- ""

if(nrow(valor_Leche)>2){
for (i in seq(3, nrow(valor_Leche), by = 3)) {
  valor_Leche$Estado[i] <- (sum(valor_Leche$valor_Leche[(i-2):i]) / sum(valor_Leche$anterior[(i-2):i]))*100-100  # Realiza la suma y división
}
}else{
  valor_Leche$Estado <- ""
}

if(mes==1){
  nuevos_datos <- data.frame(
    Consecutivo = c((data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion="Litros",
    Leche.Toneladas=valor_Leche$valor_Leche,
    Variacion.Anual=valor_Leche$valor_Leche/tail(lag(data$Leche.sin.elaborar.Volumen,11),mes)*100-100,
    Estado=as.numeric(valor_Leche$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),as.numeric(sum(valor_Leche$valor_Leche)/sum(valor_Leche$anterior)*100-100))
    } else {
      c("")
    },
    Tipo=c("")
  )
}else{
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion="Litros",
    Leche.Toneladas=valor_Leche$valor_Leche,
    Variacion.Anual=valor_Leche$valor_Leche/tail(lag(data$Leche.sin.elaborar.Volumen,11),mes)*100-100,
    Estado=as.numeric(valor_Leche$Estado),
    observaciones=if (mes==12) {
      c(rep("",11),as.numeric(sum(valor_Leche$valor_Leche)/sum(valor_Leche$anterior)*100-100))
    } else {
      c(data[fila[1]:ultima_fila,"observaciones"],"")
    },
    Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
  )
}
#Crear la nueva fila





# Escribe los datos en la hoja "Leche"
if(mes==1){
  writeData(wb, sheet = "Leche", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+11))

}else{
  writeData(wb, sheet = "Leche", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

}


addStyle(wb, sheet = "Leche",style=col1,rows = (ultima_fila+11),cols = 1:4)
addStyle(wb, sheet = "Leche",style=col2,rows = (ultima_fila+11),cols = 5)
addStyle(wb, sheet = "Leche",style=col3,rows = (ultima_fila+11),cols = 6)
addStyle(wb, sheet = "Leche",style=col4,rows = (ultima_fila+11),cols = 7:8)

# Huevos ------------------------------------------------------------------

#Leer solo la hoja de Huevos
data <- read.xlsx(entrada, sheet = "Huevos", colNames = TRUE,startRow = 10)


#Correr la funcion Huevos
fila=which(data$Año==(anio-1))
ultima_fila=nrow(data)


#Correr la funcion Pollos


valor_Huevos=f_Huevos(directorio,mes,anio)
valor_Huevos=as.data.frame(valor_Huevos)
valor_Huevos$anterior=c(data[data$Año==(anio-2),"Huevos.Unidades"],valor_Huevos[1:mes,"valor_Huevos"])
valor_Huevos$Estado <- ""

for (i in seq(3, nrow(valor_Huevos), by = 3)) {
  valor_Huevos$Estado[i] <- (sum(valor_Huevos$valor_Huevos[(i-2):i]) / sum(valor_Huevos$anterior[(i-2):i]))*100-100  # Realiza la suma y división
}

valor_Huevos$Observaciones <- ""

for (i in seq(12, nrow(valor_Huevos), by = 12)) {
  valor_Huevos$Observaciones[i] <- (sum(valor_Huevos$valor_Huevos[(i-11):i]) / sum(valor_Huevos$anterior[(i-11):i]))*100-100  # Realiza la suma y división
}
#Crear la nueva fila
nuevos_datos <- data.frame(
  Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
  Año = c(data[fila[1]:ultima_fila,"Año"],anio),
  Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
  Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"producción Fenavi"),
  Huevos.Unidades=valor_Huevos$valor_Huevos,
  Variacion.Anual=valor_Huevos$valor_Huevos/tail(lag(data$Huevos.Unidades,11),(12+mes))*100-100,
  Estado=as.numeric(valor_Huevos$Estado),
  observaciones=as.numeric(valor_Huevos$Observaciones),
  Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
)



# Escribe los datos en la hoja "Huevos"
writeData(wb, sheet = "Huevos", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

addStyle(wb, sheet = "Huevos",style=col1,rows = (ultima_fila+11),cols = 1:4)
addStyle(wb, sheet = "Huevos",style=col2,rows = (ultima_fila+11),cols = 5)
addStyle(wb, sheet = "Huevos",style=col3,rows = (ultima_fila+11),cols = 6)
addStyle(wb, sheet = "Huevos",style=col4,rows = (ultima_fila+11),cols = 7:8)
# Ovino y Caprino ------------------------------------------------------------------


#Realizar solo el proceso para los trimestres
if (mes %in% c(3, 6, 9, 12)){

#Leer solo la hoja de Huevos
data <- read.xlsx(entrada, sheet = "Ovino y Caprino trimestral", colNames = TRUE,startRow = 10)

ultima_fila=nrow(data)

if(mes==3){
  fila=ultima_fila
}else{
  fila=which(data$Año== anio)
}

#Correr la funcion Ovino_Caprino
valor_trimestre=f_Ovino_Caprino(directorio,mes,anio)
valor_trimestre=as.data.frame(valor_trimestre)
tamaño=nrow(valor_trimestre)
valor_trimestre$anterior=tail(lag(data$Ovino.y.Caprino,3),tamaño)
#Identificar el numero de trimestre
trimestre=f_trimestre(mes)
if(mes==3){
  nuevos_datos <- data.frame(
    Consecutivo = c((data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,tamaño),
    Periodicidad=c(trimestre),
    Descripcion=rep("Toneladas",tamaño),
    Ovino_caprino=valor_trimestre$valor_trimestre,
    Variacion.Anual=valor_trimestre$valor_trimestre/tail(lag(data$Ovino.y.Caprino,3),tamaño)*100-100,
    Estado=if (trimestre==4) {
      c(rep("",3),(sum(valor_trimestre$valor_trimestre))/
          (sum(tail(lag(data$Ovino.y.Caprino,3),4)))*100-100 )
    } else {
      rep("",tamaño)
    },
    observaciones=c(""),
    Tipo=c("")
  )
}else{
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,tamaño),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],trimestre),
    Descripcion=rep("Toneladas",tamaño),
    Ovino_caprino=valor_trimestre$valor_trimestre,
    Variacion.Anual=valor_trimestre$valor_trimestre/tail(lag(data$Ovino.y.Caprino,3),tamaño)*100-100,
    Estado=if (trimestre==4) {
      c(rep("",3),(sum(valor_trimestre$valor_trimestre))/
          (sum(tail(lag(data$Ovino.y.Caprino,3),4)))*100-100 )
    } else {
      rep("",tamaño)
    },
    observaciones=c(data[fila[1]:ultima_fila,"observaciones"],""),
    Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
  )
}
#Crear la nueva fila


nuevos_datos$Estado=as.numeric(nuevos_datos$Estado)

# Escribe los datos en la hoja "Ovino y Caprino trimestral"
writeData(wb, sheet = "Ovino y Caprino trimestral", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


}else{
  print("Este mes no se actualiza la hoja de Ovino y Caprino")
}



addStyle(wb, sheet = "Ovino y Caprino trimestral",style=col1,rows = (ultima_fila+11),cols = 1:4)
addStyle(wb, sheet = "Ovino y Caprino trimestral",style=col3,rows = (ultima_fila+11),cols = 5)
addStyle(wb, sheet = "Ovino y Caprino trimestral",style=col5,rows = (ultima_fila+11),cols = 6:7)

# Guardar el libro --------------------------------------------------------


if (!file.exists(salida)) {
  saveWorkbook(wb, file = salida)
} else {
  saveWorkbook(wb, file = salida,overwrite= TRUE)
}

}
