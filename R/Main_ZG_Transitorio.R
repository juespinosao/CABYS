#' @export


ZG_Transitorio=function(directorio,mes,anio){

  #Cargar librerias
  library(openxlsx)
  #utils

  #Crear el nombre de las carpetas del mes anterior y el actual
  if(mes==1){
    carpeta_anterior=nombre_carpeta(12,(anio-1))
    entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[12],"_",(anio-1),".xlsx")

  }else{
    carpeta_anterior=nombre_carpeta(mes-1,anio)
    entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

  }

  carpeta_actual=nombre_carpeta(mes,anio)

  #Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
    salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb <- loadWorkbook(entrada)


  # Maiz ------------------------------------------------------------------

  #Leer solo la hoja de Maiz
  data <- read.xlsx(entrada, sheet = "Maíz", colNames = TRUE,startRow = 10)
  #quedo igual, no cambia la función solo agregar el valor a la ultima columna

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-1))


  #Correr la funcion Maiz
  valor_Maiz=f_Maiz(directorio,mes,anio)
  Maiz_tabla=data.frame(
  Maiz=c(data[data$Año==(anio-2) |data$Año==(anio-1) | data$Año==(anio),"Maiz"])
  )
  Maiz_tabla$Anterior=lag(Maiz_tabla$Maiz,6)
  for(i in 1:4){
    Maiz_tabla[((i-1)*6 + 7):(i*6+6),"Valor_final"] <- Maiz_tabla$Anterior[((i-1)*6 + 7):(i*6+6)] * (1+valor_Maiz[i]/100)
  }

  Maiz_valor=Maiz_tabla[!is.na(Maiz_tabla$Valor_final),"Valor_final"]
  tamaño=12+mes
  variacion=lag(c(data[data$Año==(anio-2),"Maiz"],Maiz_valor),12)
  variacion=variacion[!is.na(variacion)]
  Estado <- rep("",tamaño)
  for (i in seq(3, length(Maiz_valor), by = 3)) {
    Estado[i] <- head(rep(valor_Maiz, each = 6),tamaño)[i]
  }

  Observaciones <-rep("",tamaño)

  for (i in seq(12, length(Maiz_valor), by = 12)) {
    Observaciones[i] <- sum(Maiz_valor[(i-11):i])/sum(variacion[(i-11):i])*100-100
  }
  adicional<-rep("",tamaño)
  contador=0
  for (i in seq(1, length(Maiz_valor), by = 6)) {
    contador=contador+1
    adicional[i] <- valor_Maiz[contador]
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Producción - toneladas"),
    Maiz_valor=Maiz_valor[1:tamaño],
    Variacion.Anual=head(rep(valor_Maiz, each = 6),tamaño),
    Estado=as.numeric(Estado[1:tamaño]),
    observaciones=as.numeric(Observaciones[1:tamaño]),
    Tipo=rep("",tamaño)#,
    #adicional=as.numeric(adicional[1:tamaño])
  )



  # Escribe los datos en la hoja "Ganado_Maiz"
  writeData(wb, sheet = "Maíz", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


  addStyle(wb, sheet = "Maíz", style = col1, rows = ultima_fila+11, cols = 1:4)
  addStyle(wb, sheet = "Maíz",style= col4 ,rows = ultima_fila+11,cols = 7:8)
  semestre=f_semestre(mes)
  if(semestre==1){
    addStyle(wb, sheet = "Maíz",style=colma,rows = ultima_fila+11,cols = 5)
  }else{
    addStyle(wb, sheet = "Maíz",style=colmb,rows = ultima_fila+11,cols = 5)
  }

  addStyle(wb, sheet = "Maíz",style=col3,rows = ultima_fila+11,cols = 6)

  # Arroz ------------------------------------------------------------------

  #Leer solo la hoja de Arroz
  data <- read.xlsx(wb, sheet = "Arroz", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-1))


  #Correr la funcion Arroz
  valor_Arroz=f_Arroz(directorio,mes,anio)
  valor_Arroz=as.data.frame(valor_Arroz)
  valor_Arroz$anterior=c(data[data$Año==(anio-2),"Arroz"],valor_Arroz[1:mes,"valor_Arroz"])
  valor_Arroz$Estado <- ""

  for (i in seq(3, nrow(valor_Arroz), by = 3)) {
    valor_Arroz$Estado[i] <- (sum(valor_Arroz$valor_Arroz[(i-2):i]) / sum(valor_Arroz$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  valor_Arroz$Observaciones <- ""

  for (i in seq(12, nrow(valor_Arroz), by = 12)) {
    valor_Arroz$Observaciones[i] <- (sum(valor_Arroz$valor_Arroz[(i-11):i]) / sum(valor_Arroz$anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Producción"),
    Arroz.Toneladas=valor_Arroz$valor_Arroz,
    Variacion.Anual=valor_Arroz$valor_Arroz/tail(lag(data$Arroz,11),(12+mes))*100-100,
    Estado=as.numeric(valor_Arroz$Estado),
    observaciones=as.numeric(valor_Arroz$Observaciones),
    Tipo=c(data[fila[1]:ultima_fila,"Tipo"],"")
  )



  # Escribe los datos en la hoja "Arroz"
  writeData(wb, sheet = "Arroz", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

  addStyle(wb, sheet = "Arroz",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Arroz",style=col6,rows = (ultima_fila+11),cols = 5)
  addStyle(wb, sheet = "Arroz",style=col3,rows = (ultima_fila+11),cols = 6)
  addStyle(wb, sheet = "Arroz",style=col4,rows = (ultima_fila+11),cols = 7:8)
  # Hortalizas ------------------------------------------------------------------

  #Leer solo la hoja de Hortalizass
  data <- read.xlsx(wb, sheet = "Hortalizas", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))


  #Correr la funcion Pollos
  valor_Hortalizas=f_Hortalizas(directorio,mes,anio)
  valor_Hortalizas=as.data.frame(valor_Hortalizas)
  valor_Hortalizas$anterior=c(data[data$Año==(anio-3),"Hortalizas"],valor_Hortalizas[1:(nrow(valor_Hortalizas)-12),1])
  valor_Hortalizas$variacion_anual=valor_Hortalizas$valor_Hortalizas/valor_Hortalizas$anterior*100-100
  valor_Hortalizas$Estado <- ""

  for (i in seq(3, nrow(valor_Hortalizas), by = 3)) {
    valor_Hortalizas$Estado[i] <- (sum(valor_Hortalizas$valor_Hortalizas[(i-2):i]) / sum(valor_Hortalizas$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  valor_Hortalizas$Observaciones <- ""
  for (i in seq(6, nrow(valor_Hortalizas), by = 6)) {
    valor_Hortalizas$Observaciones[i] <- (sum(valor_Hortalizas$valor_Hortalizas[(i-5):i]) / sum(valor_Hortalizas$anterior[(i-5):i]))*100-100  # Realiza la suma y división
  }

  valor_Hortalizas$Tipo <- ""

  for (i in seq(12, nrow(valor_Hortalizas), by = 12)) {
    valor_Hortalizas$Tipo[i] <- (sum(valor_Hortalizas$valor_Hortalizas[(i-11):i]) / sum(valor_Hortalizas$anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripcion"],"Toneladas"),
    Hortalizas.Kilos=valor_Hortalizas$valor_Hortalizas,
    Variacion.Anual=valor_Hortalizas$variacion_anual,
    Estado=as.numeric(valor_Hortalizas$Estado),
    Observaciones=as.numeric(valor_Hortalizas$Observaciones),
    tipo=as.numeric(valor_Hortalizas$Tipo)
  )




  # Escribe los datos en la hoja "Hortalizas"
  writeData(wb, sheet = "Hortalizas", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


  addStyle(wb, sheet = "Hortalizas",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Hortalizas",style=col7,rows = (ultima_fila+11),cols = 5)
  addStyle(wb, sheet = "Hortalizas",style=col3,rows = (ultima_fila+11),cols = 6)
  addStyle(wb, sheet = "Hortalizas",style=col4,rows = (ultima_fila+11),cols = 7:9)
  # Yuca ------------------------------------------------------------------

  #Leer solo la hoja de Yuca
  data <- read.xlsx(wb, sheet = "Yuca", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))


  #Correr la funcion Yuca
  valor_Yuca=f_Yuca(directorio,mes,anio)
  valor_Yuca=as.data.frame(valor_Yuca)
  valor_Yuca$anterior=c(data[data$Año==(anio-3),"Yuca"],valor_Yuca[1:(nrow(valor_Yuca)-12),"valor_Yuca"])
  valor_Yuca$variacion_anual=valor_Yuca$valor_Yuca/valor_Yuca$anterior*100-100
  valor_Yuca$Estado <- ""

  for (i in seq(3, nrow(valor_Yuca), by = 3)) {
    valor_Yuca$Estado[i] <- (sum(valor_Yuca$valor_Yuca[(i-2):i]) / sum(valor_Yuca$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  valor_Yuca$Observaciones <- ""
  for (i in seq(6, nrow(valor_Yuca), by = 6)) {
    valor_Yuca$Observaciones[i] <- (sum(valor_Yuca$valor_Yuca[(i-5):i]) / sum(valor_Yuca$anterior[(i-5):i]))*100-100  # Realiza la suma y división
  }

  valor_Yuca$Tipo <- ""

  for (i in seq(12, nrow(valor_Yuca), by = 12)) {
    valor_Yuca$Tipo[i] <- (sum(valor_Yuca$valor_Yuca[(i-11):i]) / sum(valor_Yuca$anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripcion"],"Toneladas"),
    Hortalizas.Kilos=valor_Yuca$valor_Yuca,
    Variacion.Anual=valor_Yuca$variacion_anual,
    Estado=as.numeric(valor_Yuca$Estado),
    Observaciones=as.numeric(valor_Yuca$Observaciones),
    tipo=as.numeric(valor_Yuca$Tipo)
  )



  # Escribe los datos en la hoja "Yuca"
  writeData(wb, sheet = "Yuca", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


  addStyle(wb, sheet = "Yuca",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Yuca",style=col6,rows = (ultima_fila+11),cols = 5)
  addStyle(wb, sheet = "Yuca",style=col3,rows = (ultima_fila+11),cols = 6)
  addStyle(wb, sheet = "Yuca",style=col4,rows = (ultima_fila+11),cols = 7:9)
# # Trigo ------------------------------------------------------------------


# #Realizar solo el proceso para los trimestres
# if (mes %in% c(3, 6, 9, 12)){

#   #Leer solo la hoja de Huevos
#   data <- read.xlsx(entrada, sheet = "Trigo trimestral", colNames = TRUE,startRow = 10)

#   ultima_fila=nrow(data)
#   fila=which(data$Año==(anio-1))


#   #Correr la funcion Ovino_Caprino
#   valor_Trigo=f_Trigo(directorio,mes,anio)
#   valor_Trigo=rep(valor_Trigo,each=2)
#   trimestre=f_trimestre(mes)
#   valor_actual=NULL
#   valor_anterior=as.numeric(data[data$Año==(anio-2),"Trigo"])
#   for(i in 1:(4+trimestre)){
#   valor_actual[i]=as.numeric(valor_anterior[i]*(1+valor_Trigo[i]/100))
#   valor_anterior=c(valor_anterior,valor_actual[i])
#   }
#   tamaño=length(valor_actual)
#   tabla_trigo=cbind(valor_actual,valor_anterior[1:tamaño])
#   tabla_trigo=as.data.frame(tabla_trigo)
#   colnames(tabla_trigo)=c("valor_actual","valor_anterior")
#   tabla_trigo$variacion_anual=tabla_trigo$valor_actual/tabla_trigo$valor_anterior*100-100
#   #Identificar el numero de trimestre
#   trimestre=f_trimestre(mes)
#   tabla_trigo$estado <-  ""
#   for (i in seq(4, nrow(tabla_trigo), by = 4)) {
#     tabla_trigo$estado[i] <- (sum(tabla_trigo$valor_actual[(i-3):i]) / sum(tabla_trigo$valor_anterior[(i-3):i]))*100-100  # Realiza la suma y división
#   }


#   nuevos_datos <- data.frame(
#     Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
#     Año = c(data[fila[1]:ultima_fila,"Año"],anio),
#     Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],trimestre),
#     Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Toneladas"),
#     Trigo_valor=tabla_trigo$valor_actual,
#     Variacion.Anual=tabla_trigo$variacion_anual,
#     Estado=as.numeric(tabla_trigo$estado),
#     observaciones=rep("",tamaño),
#     Tipo=rep("",tamaño)
#   )

#   # Escribe los datos en la hoja "Ovino y Caprino trimestral"
#   writeData(wb, sheet = "Trigo trimestral", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

#   #Añadir estilos de celda
#   addStyle(wb, sheet = "Trigo trimestral",style=col1,rows = (ultima_fila+11),cols = 1:4)
#   addStyle(wb, sheet = "Trigo trimestral",style=col3,rows = (ultima_fila+11),cols = 5)
#   addStyle(wb, sheet = "Trigo trimestral",style=col5,rows = (ultima_fila+11),cols = 6:7)

# }else{
#   print("Este mes no se actualiza la hoja de Trigo")
# }
#
#
# # Sorgo y Cebada ------------------------------------------------------------------


# #Realizar solo el proceso para los trimestres
# if (mes %in% c(3, 6, 9, 12)){

#   #Leer solo la hoja de Huevos
#   data <- read.xlsx(entrada, sheet = "Sorgo y Cebada trimestral", colNames = TRUE,startRow = 10)

#   ultima_fila=nrow(data)
#   fila=which(data$Año==(anio-1))


#   #Correr la funcion Ovino_Caprino
#   valor_sorgo=f_Sorgo_Cebada(directorio,mes,anio)
#   valor_sorgo=rep(valor_sorgo,each=2)
#   trimestre=f_trimestre(mes)
#   tamaño=length(valor_actual)
#   valor_anterior=c(data[data$Año==(anio-2),"Sorgo.y.Cebada"],valor_sorgo[1:trimestre])
#   tamaño=length(valor_anterior)
#   tabla_sorgo=cbind(valor_sorgo[1:tamaño],valor_anterior)
#   tabla_sorgo=as.data.frame(tabla_sorgo)
#   colnames(tabla_sorgo)=c("valor_actual","valor_anterior")
#   tabla_sorgo$variacion_anual=tabla_sorgo$valor_actual/tabla_sorgo$valor_anterior*100-100
#   #Identificar el numero de trimestre
#   trimestre=f_trimestre(mes)
#   tabla_sorgo$estado <-  ""
#   for (i in seq(4, nrow(tabla_sorgo), by = 4)) {
#     tabla_sorgo$estado[i] <- (sum(tabla_sorgo$valor_actual[(i-3):i]) / sum(tabla_sorgo$valor_anterior[(i-3):i]))*100-100  # Realiza la suma y división
#   }


#   nuevos_datos <- data.frame(
#     Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
#     Año = c(data[fila[1]:ultima_fila,"Año"],anio),
#     Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],trimestre),
#     Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Toneladas"),
#     sorgo_valor=tabla_sorgo$valor_actual,
#     Variacion.Anual=tabla_sorgo$variacion_anual,
#     Estado=as.numeric(tabla_sorgo$estado),
#     observaciones=rep("",tamaño),
#     Tipo=rep("",tamaño)
#   )
#     # Escribe los datos en la hoja "Ovino y Caprino trimestral"
#     writeData(wb, sheet = "Sorgo y Cebada trimestral", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

#     #Añadir estilos de celda
#     addStyle(wb, sheet = "Sorgo y Cebada trimestral",style=col1,rows = (ultima_fila+11),cols = 1:4)
#     addStyle(wb, sheet = "Sorgo y Cebada trimestral",style=col3,rows = (ultima_fila+11),cols = 5)
#     addStyle(wb, sheet = "Sorgo y Cebada trimestral",style=col5,rows = (ultima_fila+11),cols = 6:7)

# }else{
#   print("Este mes no se actualiza la hoja de Sorgo y Cebada")
# }
#
#

 ## Tabaco ------------------------------------------------------------------


 ##Realizar solo el proceso para los trimestres
 #if (mes %in% c(3, 6, 9, 12)){

 #  #Leer solo la hoja de Huevos
 #  data <- read.xlsx(entrada, sheet = "Tabaco trimestral", colNames = TRUE,startRow = 10)

 #  ultima_fila=nrow(data)
 #  fila=which(data$Año==(anio-1))

 #  #Correr la funcion Ovino_Caprino
 #  valor_Tabaco=f_Tabaco(directorio,mes,anio)
 #  trimestre=f_trimestre(mes)
 #  valor_actual=NULL
 #  valor_anterior=as.numeric(data[data$Año==(anio-2),"Tabaco"])
 #  for(i in 1:(4+trimestre)){
 #    valor_actual[i]=as.numeric(valor_anterior[i]*(1+valor_Tabaco[i]/100))
 #    valor_anterior=c(valor_anterior,valor_actual[i])
 #  }
 #  tamaño=length(valor_actual)
 #  tabla_Tabaco=cbind(valor_actual,valor_anterior[1:tamaño])
 #  tabla_Tabaco=as.data.frame(tabla_Tabaco)
 #  colnames(tabla_Tabaco)=c("valor_actual","valor_anterior")
 #  tabla_Tabaco$variacion_anual=tabla_Tabaco$valor_actual/tabla_Tabaco$valor_anterior*100-100
 #  #Identificar el numero de trimestre
 #  tabla_Tabaco$estado <-  ""
 #  for (i in seq(2, nrow(tabla_Tabaco), by = 2)) {
 #    tabla_Tabaco$estado[i] <- (sum(tabla_Tabaco$valor_actual[(i-1):i]) / sum(tabla_Tabaco$valor_anterior[(i-1):i]))*100-100  # Realiza la suma y división
 #  }


 #  nuevos_datos <- data.frame(
 #    Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
 #    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
 #    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],trimestre),
 #    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Toneladas"),
 #    Tabaco_valor=tabla_Tabaco$valor_actual,
 #    Variacion.Anual=tabla_Tabaco$variacion_anual,
 #    Estado=as.numeric(tabla_Tabaco$estado),
 #    observaciones=rep("",tamaño),
 #    Tipo=rep("",tamaño)
 #  )
 #  # Escribe los datos en la hoja "Tabaco trimestral"
 #  writeData(wb, sheet = "Tabaco trimestral", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

 #  #Añadir estilos de celda
 #  addStyle(wb, sheet = "Tabaco trimestral",style=col1,rows = (ultima_fila+11),cols = 1:4)
 #  addStyle(wb, sheet = "Tabaco trimestral",style=col3,rows = (ultima_fila+11),cols = 5)
 #  addStyle(wb, sheet = "Tabaco trimestral",style=col5,rows = (ultima_fila+11),cols = 6:7)


 #}else{
 #  print("Este mes no se actualiza la hoja de Tabaco")
 #}
#
#

  # Papa ------------------------------------------------------------------

  #Leer solo la hoja de Papa
  data <- read.xlsx(wb, sheet = "Papa", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)

  fila=which(data$Año==2013)
  tamaño=(anio-2013)*12+mes

  #Correr la funcion Pollos
  valor_Papa=f_Papa(directorio,mes,anio)
  valor_Papa=valor_Papa[1:tamaño]
  valor_Papa=as.data.frame(valor_Papa)
  valor_Papa$anterior=c(data[data$Año==(2012),"Papa"],valor_Papa[1:(nrow(valor_Papa)-12),"valor_Papa"])
  valor_Papa$variacion_anual=valor_Papa$valor_Papa/valor_Papa$anterior*100-100
  valor_Papa$Estado <- ""

  for (i in seq(3, nrow(valor_Papa), by = 3)) {
    valor_Papa$Estado[i] <- (sum(valor_Papa$valor_Papa[(i-2):i]) / sum(valor_Papa$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  valor_Papa$Observaciones <- ""
  for (i in seq(6, nrow(valor_Papa), by = 6)) {
    valor_Papa$Observaciones[i] <- (sum(valor_Papa$valor_Papa[(i-5):i]) / sum(valor_Papa$anterior[(i-5):i]))*100-100  # Realiza la suma y división
  }

  valor_Papa$Tipo <- ""

  for (i in seq(12, nrow(valor_Papa), by = 12)) {
    valor_Papa$Tipo[i] <- (sum(valor_Papa$valor_Papa[(i-11):i]) / sum(valor_Papa$anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo =c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Toneladas"),
    Papa.Kilos=valor_Papa$valor_Papa,
    Variacion.Anual=valor_Papa$variacion_anual,
    Estado=as.numeric(valor_Papa$Estado),
    Observaciones=as.numeric(valor_Papa$Observaciones),
    tipo=as.numeric(valor_Papa$Tipo)
  )



  # Escribe los datos en la hoja "Papa"
  writeData(wb, sheet = "Papa", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


  addStyle(wb, sheet = "Papa",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Papa",style=col5,rows = (ultima_fila+11),cols = 5)
  addStyle(wb, sheet = "Papa",style=col3,rows = (ultima_fila+11),cols = 6)
  addStyle(wb, sheet = "Papa",style=col4,rows = (ultima_fila+11),cols = 7:9)
  # Legumbres ------------------------------------------------------------------

  #Leer solo la hoja de Legumbres
  data <- read.xlsx(wb, sheet = "Legumbres", colNames = TRUE,startRow = 10)


  ultima_fila=nrow(data)


  #Correr la funcion Pollos
  valor_Legumbres=f_Legumbres(directorio,mes,anio)
  vari=valor_Legumbres$variacion
  valor_Legumbres=as.data.frame(valor_Legumbres$vector)
  colnames(valor_Legumbres)=c("valor_Legumbres")


  valor_Legumbres$anterior=lag(valor_Legumbres$valor_Legumbres,12)
  valor_Legumbres$variacion_anual=valor_Legumbres$valor_Legumbres/valor_Legumbres$anterior*100-100
  valor_Legumbres$Estado <- ""

  for (i in seq(3, nrow(valor_Legumbres), by = 3)) {
    valor_Legumbres$Estado[i] <- (sum(valor_Legumbres$valor_Legumbres[(i-2):i]) / sum(valor_Legumbres$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  valor_Legumbres$Observaciones <- ""
  for (i in seq(6, nrow(valor_Legumbres), by = 6)) {
    valor_Legumbres$Observaciones[i] <- (sum(valor_Legumbres$valor_Legumbres[(i-5):i]) / sum(valor_Legumbres$anterior[(i-5):i]))*100-100  # Realiza la suma y división
  }

  valor_Legumbres$Tipo <- ""

  for (i in seq(12, nrow(valor_Legumbres), by = 12)) {
    valor_Legumbres$Tipo[i] <- (sum(valor_Legumbres$valor_Legumbres[(i-11):i]) / sum(valor_Legumbres$anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }



if (mes %in% c(3, 6, 9, 12)){
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo =c(data[1:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[1:ultima_fila,"Año"],anio),
    Periodicidad=c(data[1:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[1:ultima_fila,"Descripción"],"Toneladas"),
    Legumbres.Kilos=valor_Legumbres$valor_Legumbres,
    Variacion.Anual=valor_Legumbres$variacion_anual,
    Estado=as.numeric(valor_Legumbres$Estado),
    Observaciones=as.numeric(valor_Legumbres$Observaciones),
    tipo=as.numeric(valor_Legumbres$Tipo)
  )
  }else{
  valor_mes=tail(lag(data$`Legumbres.verdes.y.secas.(frijoles,.arvejas,.habas,.garbanzos,.lentejas,.etc.)`,11),1)*(1+vari/100)
  valor_Legumbres[nrow(valor_Legumbres),"valor_Legumbres"]=valor_mes
  valor_Legumbres[nrow(valor_Legumbres),"variacion_anual"]=valor_mes/tail(
                                                           lag(data$`Legumbres.verdes.y.secas.(frijoles,.arvejas,.habas,.garbanzos,.lentejas,.etc.)`,11),1)*100-100
  valor_Legumbres[nrow(valor_Legumbres)-1,"valor_Legumbres"]=tail(data$`Legumbres.verdes.y.secas.(frijoles,.arvejas,.habas,.garbanzos,.lentejas,.etc.)`,1)
  valor_Legumbres[nrow(valor_Legumbres)-1,"variacion_anual"]=tail(data$Variacion.Anual,1)

  nuevos_datos <- data.frame(
    Consecutivo =c(data[1:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[1:ultima_fila,"Año"],anio),
    Periodicidad=c(data[1:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[1:ultima_fila,"Descripción"],"Toneladas"),
    Legumbres.Kilos=valor_Legumbres$valor_Legumbres,
    Variacion.Anual=valor_Legumbres$variacion_anual,
    Estado=as.numeric(valor_Legumbres$Estado),
    Observaciones=as.numeric(valor_Legumbres$Observaciones),
    tipo=as.numeric(valor_Legumbres$Tipo),
    adicional=c(rep("",nrow(valor_Legumbres)-2),"Variación mensualización",vari)
  )
  }



  # Escribe los datos en la hoja "Legumbres"
  writeData(wb, sheet = "Legumbres", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = 11)


  addStyle(wb, sheet = "Legumbres",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Legumbres",style=col7,rows = (ultima_fila+11),cols = 5)
  addStyle(wb, sheet = "Legumbres",style=col3,rows = (ultima_fila+11),cols = 6)
  addStyle(wb, sheet = "Legumbres",style=col4,rows = (ultima_fila+11),cols = 7:9)

  # Guardar el libro --------------------------------------------------------


  if (!file.exists(salida)) {
    saveWorkbook(wb, file = salida)
  } else {
    saveWorkbook(wb, file = salida,overwrite= TRUE)
  }

}
