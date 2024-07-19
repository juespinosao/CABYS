#' @export
### pecuario

ZG_Permanentes=function(directorio,mes,anio){

  #Cargar librerias
  library(openxlsx)
  #utils

  #Crear el nombre de las carpetas del mes anterior y el actual
  if(mes==1){
    carpeta_anterior=nombre_carpeta(12,(anio-1))
    entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[12],"_",(anio-1),".xlsx")

  }else{
    carpeta_anterior=nombre_carpeta(mes-1,anio)
    entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

  }

  carpeta_actual=nombre_carpeta(mes,anio)

  #Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
    salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb <- loadWorkbook(entrada)


  # Cafe_Pergamino ------------------------------------------------------------------

    #Leer solo la hoja de Cafe_Pergamino
    data <- read.xlsx(wb, sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)


    ultima_fila=nrow(data)
    fila=which(data$Año==(anio-2))


    #Correr la funcion Cafe_Pergamino
    valor_Cafe_verde_pergamino=as.data.frame(f_Cafe_verde_pergamino(directorio,mes,anio))
    anterior_pergamino=c(data[data$Año==(anio-3),"Producción.Total.de.Café.Pergamino"],valor_Cafe_verde_pergamino[1:(nrow(valor_Cafe_verde_pergamino)-12),"produccion_total_pergamino"])
    tamaño=nrow(valor_Cafe_verde_pergamino)
    Estado <- rep("",tamaño)

    for (i in seq(3, nrow(valor_Cafe_verde_pergamino), by = 3)) {
      Estado[i] <- (sum(valor_Cafe_verde_pergamino$produccion_total_pergamino[(i-2):i]) / sum(anterior_pergamino[(i-2):i]))*100-100  # Realiza la suma y división
    }

    Observaciones <-rep("",tamaño)

    for (i in seq(12, tamaño, by = 12)) {
      Observaciones[i] <- (sum(valor_Cafe_verde_pergamino$produccion_total_pergamino[(i-11):i]) / sum(anterior_pergamino[(i-11):i]))*100-100  # Realiza la suma y división
    }
    #Crear la nueva fila
    nuevos_datos <- data.frame(
      Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
      Año = c(data[fila[1]:ultima_fila,"Año"],anio),
      Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
      Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Sacos de 60 kilogramos de CAFÉ PERGAMINO"),
      valor_Cafe_verde_pergamino,
      Variacion.Anual=valor_Cafe_verde_pergamino$produccion_total_pergamino/anterior_pergamino*100-100,
      Estado=as.numeric(Estado),
      observaciones=as.numeric(Observaciones),
      Tipo=rep("",tamaño)
    )



    # Escribe los datos en la hoja "Ganado_Cafe_Pergamino"
    writeData(wb, sheet = "Cafe Pergamino", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))

    addStyle(wb, sheet = "Cafe Pergamino", style = col1, rows = (ultima_fila+10), cols = 1:4)
    addStyle(wb, sheet = "Cafe Pergamino",style=col8,rows = (ultima_fila+10),cols = 5:8)
    addStyle(wb, sheet = "Cafe Pergamino",style=col9,rows = (ultima_fila+10),cols = 9:11)
    addStyle(wb, sheet = "Cafe Pergamino",style=col7,rows = (ultima_fila+10),cols = 12)
    addStyle(wb, sheet = "Cafe Pergamino",style= col4 ,rows = (ultima_fila+10),cols = 13)
    addStyle(wb, sheet = "Cafe Pergamino",style=col8,rows = (ultima_fila+10),cols = 14)

# Cafe_verde ------------------------------------------------------------------

  #Leer solo la hoja de Cafe_verde
  data <- read.xlsx(wb, sheet = "Cafe Verde", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))
  anterior_stocks=valor_Cafe_verde_pergamino$Valor_existencias-lag(valor_Cafe_verde_pergamino$Valor_existencias)
  anterior_stocks[1]=valor_Cafe_verde_pergamino$Valor_existencias[1]-data[data$Año==(anio-3) & data$Periodicidad==12,"Cambio.en.Existencias.de.verde.Miles.Sacos"]
  anterior_expos=c(data[data$Año==(anio-3),"Exportaciones.Totales"],valor_Cafe_verde_pergamino[1:(nrow(valor_Cafe_verde_pergamino)-12),"total_exportaciones"])
  anterior_impos=c(data[data$Año==(anio-3),"Importaciones.Totales"],valor_Cafe_verde_pergamino[1:(nrow(valor_Cafe_verde_pergamino)-12),"importaciones"])
  anterior_consumo=c(data[data$Año==(anio-3),"Consumo.Intermedio"],valor_Cafe_verde_pergamino[1:(nrow(valor_Cafe_verde_pergamino)-12),"consumo_interno"])
  anterior_produccion=c(data[data$Año==(anio-3),"Producción.café.verde"],valor_Cafe_verde_pergamino[1:(nrow(valor_Cafe_verde_pergamino)-12),"valor_produccion"])

  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Sacos de 60 kilogramos de CAFÉ VERDE"),
    valor_Cafe_verde_pergamino[,1:5],
    cambio_stock=anterior_stocks,
    Variacion_expos=valor_Cafe_verde_pergamino$total_exportaciones/anterior_expos*100-100,
    Variacion_impos=valor_Cafe_verde_pergamino$importaciones/anterior_impos*100-100,
    Variacion_consumo=valor_Cafe_verde_pergamino$consumo_interno/anterior_consumo*100-100,
    Variacion_produccion=valor_Cafe_verde_pergamino$valor_produccion/anterior_produccion*100-100,
    Estado=rep("",tamaño),
    observaciones=rep("",tamaño),
    Tipo=rep("",tamaño)
  )




  # Escribe los datos en la hoja "Cafe_verde"
  writeData(wb, sheet = "Cafe Verde", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))

  addStyle(wb, sheet = "Cafe Verde",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Cafe Verde",style=col9,rows = (ultima_fila+11),cols = 5:8)
  addStyle(wb, sheet = "Cafe Verde",style=col6,rows = (ultima_fila+11),cols = 9:14)

  # Cafetos ------------------------------------------------------------------

  #Leer solo la hoja de Cafetos
  data <- read.xlsx(wb, sheet = "Cafetos", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)
  fila=which(data$Año== anio)


  #Correr la funcion Cafetos
  valor_Cafetos=f_Cafetos(directorio,mes,anio)
  valor_Cafetos=as.data.frame(valor_Cafetos)
  valor_Cafetos$anterior=tail(lag(data$Cafetos,11),mes)
  valor_Cafetos$Estado <- ""

  if(nrow(valor_Cafetos)>2){
  for (i in seq(3, nrow(valor_Cafetos), by = 3)) {
    valor_Cafetos$Estado[i] <- (sum(valor_Cafetos$valor_Cafetos[(i-2):i]) / sum(valor_Cafetos$anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
  }else{
    valor_Cafetos$Estado <- ""
}

if(mes==1){
  nuevos_datos <- data.frame(
    Consecutivo = c((data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(anio),
    Periodicidad=c(data[fila,"Periodicidad"],mes),
    Descripcion=rep("Hectáreas Renovadas para Producción",mes),
    Cafetos.Toneladas=valor_Cafetos$valor_Cafetos,
    Variacion.Anual=valor_Cafetos$valor_Cafetos/valor_Cafetos$anterior*100-100,
    Estado=as.numeric(valor_Cafetos$Estado),
    observaciones=rep("",mes),
    Tipo=rep("",mes)
  )

}else{
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila,"Año"],anio),
    Periodicidad=c(data[fila,"Periodicidad"],mes),
    Descripcion=rep("Hectáreas Renovadas para Producción",mes),
    Cafetos.Toneladas=valor_Cafetos$valor_Cafetos,
    Variacion.Anual=valor_Cafetos$valor_Cafetos/valor_Cafetos$anterior*100-100,
    Estado=as.numeric(valor_Cafetos$Estado),
    observaciones=rep("",mes),
    Tipo=rep("",mes)
  )

}
  #Crear la nueva fila




  # Escribe los datos en la hoja "Cafetos"
  if(is.null(length(fila))){
    writeData(wb, sheet = "Cafetos", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila+10))

  }else{
    writeData(wb, sheet = "Cafetos", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))

  }

  addStyle(wb, sheet = "Cafetos",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Cafetos",style=col7,rows = (ultima_fila+10),cols = 5)
  addStyle(wb, sheet = "Cafetos",style=col3,rows = (ultima_fila+10),cols = 6)
  addStyle(wb, sheet = "Cafetos",style=col4,rows = (ultima_fila+10),cols = 7)
  # Banano ------------------------------------------------------------------

  #Leer solo la hoja de Bananos
  data <- read.xlsx(wb, sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)

  ultima_fila=nrow(data)
  fila=which(data$Año== (anio-2))
  fila_2015=which(data$Año== 2015)

  #Correr la funcion Pollos

  valor_Banano=f_Banano(directorio,mes,anio)
  tamaño=length(valor_Banano$exportaciones)
  valor_Banano$consumo_interno=tail(as.numeric(valor_Banano$consumo_interno[,1]),tamaño)

  #Crear valores necesarios
  Prom2015_exportaciones=mean(data[fila_2015,"Banano.de.Exportación.(DANE).ktes"])
  Indice_exportacion=valor_Banano$exportaciones/Prom2015_exportaciones*100
  Ponderador_expos=74.0461933276304
  Prom2015_consumo=7921383
  Indice_consumo=valor_Banano$consumo_interno/Prom2015_consumo*100
  Ponderador_consumo=25.9538066723696
  Indice_ponderado=((Indice_exportacion*Ponderador_expos)+(tail(Indice_consumo,tamaño)*Ponderador_consumo))/100
  vector_banano=cbind(valor_Banano$exportaciones,Indice_exportacion,Ponderador_expos,tail(valor_Banano$consumo_interno,tamaño),tail(Indice_consumo,tamaño),
                  Ponderador_consumo,tail(Indice_ponderado,tamaño))
  vector_banano=as.data.frame(vector_banano)
  Indice_ponderado_anterior=c(data[data$Año==(anio-3),"Indice.de.producción.ponderado"],Indice_ponderado[1:(length(Indice_ponderado)-12)])
  Expo_anterior=c(data[data$Año==(anio-3),"Banano.de.Exportación.(DANE).ktes"],valor_Banano$exportaciones[1:(length(valor_Banano$exportaciones)-12)])
  Consumo_anterior=c(data[data$Año==(anio-3),"Banano.consumo.interno(SIPSA).ton"],valor_Banano$consumo_interno[1:(length(valor_Banano$consumo_interno)-12)])
  Indice_consumo_anterior=c(data[data$Año==(anio-3),"Banano.consumo.interno(SIPSA)ÍNDICE"],Indice_consumo[1:(length(Indice_consumo)-12)])


  expo_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
    expo_trim[i] <- (sum(valor_Banano$exportaciones[(i-2):i]) / sum(Expo_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  consumo_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
   consumo_trim[i] <- (sum(valor_Banano$consumo_interno[(i-2):i]) / sum(Consumo_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  ponderado_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
    ponderado_trim[i] <- (sum(Indice_ponderado[(i-2):i]) / sum(Indice_ponderado_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  consumo_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    consumo_anual[i] <- (sum(Indice_consumo[(i-11):i]) / sum(Indice_consumo_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }

  ponderado_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    ponderado_anual[i] <- (sum(Indice_ponderado[(i-11):i]) / sum(Indice_ponderado_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }


  consumo_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    consumo_anual[i] <- (sum(valor_Banano$consumo_interno[(i-11):i]) / sum(Consumo_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }

  expo_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    expo_anual[i] <- (sum(valor_Banano$exportaciones[(i-11):i]) / sum(Expo_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }

  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=rep("Toneladas",tamaño),
    Banano.Kilos=vector_banano,
    indice_Variacion_Anual=Indice_ponderado/Indice_ponderado_anterior*100-100,
    exportacion_Variacion_Anual=valor_Banano$exportaciones/Expo_anterior*100-100,
    consumo_Variacion_Anual=tail(valor_Banano$consumo_interno,tamaño)/tail(Consumo_anterior,tamaño)*100-100,
    indice_Variacion_Anual2=Indice_ponderado/Indice_ponderado_anterior*100-100,
    Expos_trim=as.numeric(expo_trim),
    Consumo_trim=as.numeric(consumo_trim),
    Indice_trim=as.numeric(ponderado_trim),
    Tipo=rep("",tamaño),
    Var_anual_indice_consumo=as.numeric(consumo_anual),
    Var_anual_indice_ponderado=as.numeric(ponderado_anual),
    Var_anual_consumo=as.numeric(consumo_anual),
    Var_anual_expos=as.numeric(expo_anual)
  )




  # Escribe los datos en la hoja "Banano Total(Expos+Interno)"
  writeData(wb, sheet = "Banano Total(Expos+Interno)", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+11))

  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col1,rows = (ultima_fila+12),cols = 1:4)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 5:7)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col7,rows = (ultima_fila+12),cols = 8)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 9:11)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col3,rows = (ultima_fila+12),cols = 12:19)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 20:21)
  addStyle(wb, sheet = "Banano Total(Expos+Interno)",style=col3,rows = (ultima_fila+12),cols = 22:23)

  # Platano ------------------------------------------------------------------

  #Leer solo la hoja de Platanos
  data <- read.xlsx(wb, sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)

  ultima_fila=nrow(data)
  fila=which(data$Año== (anio-2))


  #Correr la funcion Pollos
  valor_Platano=f_Platano(directorio,mes,anio)
  valor_Platano$consumo_interno=tail(as.numeric(valor_Platano$consumo_interno[,1]),tamaño)

  #Crear valores necesarios

  Prom2015_exportaciones=7963.12480916667
  Indice_exportacion=valor_Platano$exportaciones/Prom2015_exportaciones*100
  Ponderador_expos=5.76126001657917
  Prom2015_consumo=21568029.67
  Indice_consumo=valor_Platano$consumo_interno/Prom2015_consumo*100
  Ponderador_consumo=94.2387399834208
  Indice_ponderado=((Indice_exportacion*Ponderador_expos)+(Indice_consumo*Ponderador_consumo))/100
  vector_Platano=cbind(valor_Platano$exportaciones,Indice_exportacion,Ponderador_expos,valor_Platano$consumo_interno,Indice_consumo,
                  Ponderador_consumo,Indice_ponderado)
  vector_Platano=as.data.frame(vector_Platano)
  Indice_ponderado_anterior=c(data[data$Año==(anio-3),"ÍNDICE.de.producción.ponderado"],Indice_ponderado[1:(length(Indice_ponderado)-12)])
  Expo_anterior=c(data[data$Año==(anio-3),"Plátano.de.Exportación"],valor_Platano$exportaciones[1:(length(valor_Platano$exportaciones)-12)])
  Consumo_anterior=c(data[data$Año==(anio-3),"Plátano.consumo.interno(SIPSA).ton"],valor_Platano$consumo_interno[1:(length(valor_Platano$consumo_interno)-12)])
  Indice_consumo_anterior=c(data[data$Año==(anio-3),"Platano.consumo.interno(SIPSA)ÍNDICE"],Indice_consumo[1:(length(Indice_consumo)-12)])
  Indice_exportacion_anterior=c(data[data$Año==(anio-3),"Plátano.de.Exportación.ÍNDICE"],Indice_exportacion[1:(length(Indice_exportacion)-12)])


  expo_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
    expo_trim[i] <- (sum(valor_Platano$exportaciones[(i-2):i]) / sum(Expo_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  consumo_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
    consumo_trim[i] <- (sum(valor_Platano$consumo_interno[(i-2):i]) / sum(Consumo_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  ponderado_trim <- rep("",length(Indice_exportacion))

  for (i in seq(3, length(Indice_exportacion), by = 3)) {
    ponderado_trim[i] <- (sum(Indice_ponderado[(i-2):i]) / sum(Indice_ponderado_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  consumo_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    consumo_anual[i] <- (sum(Indice_consumo[(i-11):i]) / sum(Indice_consumo_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }




  expo_anual <- rep("",length(Indice_exportacion))

  for (i in seq(12, length(Indice_exportacion), by = 12)) {
    expo_anual[i] <- (sum(Indice_exportacion[(i-11):i]) / sum(Indice_exportacion_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  tamaño=length(valor_Platano$exportaciones)

  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=rep("Toneladas",tamaño),
    Platano.Kilos=vector_Platano,
    indice_Variacion_Anual=Indice_ponderado/Indice_ponderado_anterior*100-100,
    exportacion_Variacion_Anual=valor_Platano$exportaciones/Expo_anterior*100-100,
    consumo_Variacion_Anual=valor_Platano$consumo_interno/Consumo_anterior*100-100,
    indice_Variacion_Anual2=Indice_ponderado/Indice_ponderado_anterior*100-100,
    Expos_trim=as.numeric(expo_trim),
    Consumo_trim=as.numeric(consumo_trim),
    Indice_trim=as.numeric(ponderado_trim),
    Tipo=as.numeric(expo_anual),
    Total_Interno=as.numeric(consumo_anual),
    var_anual=rep("",tamaño)
  )




  # Escribe los datos en la hoja "Plátano Total(Expos+Interno)"
  writeData(wb, sheet = "Plátano Total(Expos+Interno)", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+11))

  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col1,rows = (ultima_fila+12),cols = 1:4)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col6,rows = (ultima_fila+12),cols = 5)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 6:7)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col7,rows = (ultima_fila+12),cols = 8)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 9:11)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col3,rows = (ultima_fila+12),cols = 12:18)
  addStyle(wb, sheet = "Plátano Total(Expos+Interno)",style=col2,rows = (ultima_fila+12),cols = 19:20)
  # Frutas citricas ------------------------------------------------------------------

  #Leer solo la hoja de Frutas
  data <- read.xlsx(wb, sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)

  ultima_fila=nrow(data)
  fila=which(data$Año==2013)
  fila_anio_ant=which(data$Año== (anio-1))
  fila_2anio_ant=which(data$Año== (anio-2))
  fila_anterior=max(which(data$Mes== mes))


  #Correr la funcion Pollos
  valor_Frutas=f_Frutas(directorio,mes,anio)
exportaciones=valor_Frutas$variacion
consumo_interno=valor_Frutas$vector
IPP=valor_Frutas$IPP
nuevos_datos=data.frame(
consecutivo=data[ultima_fila,1]+1,
periodo=anio,
Mes=mes
)
  # Escribe los datos en la hoja "Frutas Total(Expos+Interno)"
writeData(wb, sheet = "Frutas Citricas", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+3))
writeData(wb, sheet = "Frutas Citricas", x = exportaciones[,1],colNames = FALSE,startCol = "D", startRow = (fila_anio_ant[1]+2))
writeData(wb, sheet = "Frutas Citricas", x =as.numeric(consumo_interno[,1]),colNames = FALSE,startCol = "E", startRow = (fila[1]+2))
writeData(wb, sheet = "Frutas Citricas", x = exportaciones[,3],colNames = FALSE,startCol = "J", startRow = (fila_anio_ant[1]+2))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("D",ultima_fila+3,"/D",fila_anterior+2,"*100-100") ,startCol = "G", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("E",ultima_fila+3,"/E",fila_anterior+2,"*100-100") ,startCol = "H", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("E",ultima_fila+3,"*R",ultima_fila+3,"/100") ,startCol = "K", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("J",ultima_fila+3,"/J",fila_anterior+2,"*100-100") ,startCol = "M", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("K",ultima_fila+3,"/K",fila_anterior+2,"*100-100") ,startCol = "N", startRow = (ultima_fila[1]+3))
writeData(wb, sheet = "Frutas Citricas", x = exportaciones[,5],colNames = FALSE,startCol = "P", startRow = (fila_anio_ant[1]+2))
writeData(wb, sheet = "Frutas Citricas", x = IPP,colNames = FALSE,startCol = "R", startRow = (fila_2anio_ant[1]+2))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("P",ultima_fila+3,"/P",fila_anterior+2,"*100-100") ,startCol = "T", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("R",ultima_fila+3,"/R",fila_anterior+2,"*100-100") ,startCol = "V", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("((X3/12)/100)*(D",ultima_fila+3,"/AVERAGE(D125:D136)*100)") ,startCol = "X", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("((Y3/12)/100)*(E",ultima_fila+3,"/AVERAGE(E125:E136)*100)") ,startCol = "Y", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("SUM(X",ultima_fila+3,":Y",ultima_fila+3,")") ,startCol = "Z", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("X",ultima_fila+3,"/X",fila_anterior+2,"*100-100") ,startCol = "AB", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("Y",ultima_fila+3,"/Y",fila_anterior+2,"*100-100") ,startCol = "AC", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("Z",ultima_fila+3,"/Z",fila_anterior+2,"*100-100") ,startCol = "AD", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("Z",ultima_fila+3) ,startCol = "AG", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AG",ultima_fila+3,"/AG",fila_anterior+2,"*100-100") ,startCol = "AI", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AD",ultima_fila+3,"-AI",ultima_fila+3) ,startCol = "AJ", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("((AL3/12)/100)*(J",ultima_fila+3,"/AVERAGE(J125:J136)*100)") ,startCol = "AL", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("((AM3/12)/100)*(K",ultima_fila+3,"/AVERAGE(K125:K136)*100)") ,startCol = "AM", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("SUM(AL",ultima_fila+3,":AN",ultima_fila+3,")") ,startCol = "AO", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AL",ultima_fila+3,"/AL",fila_anterior+2,"*100-100") ,startCol = "AQ", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AM",ultima_fila+3,"/AM",fila_anterior+2,"*100-100") ,startCol = "AR", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AO",ultima_fila+3,"/AO",fila_anterior+2,"*100-100") ,startCol = "AS", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AL",ultima_fila+3) ,startCol = "AU", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AO",ultima_fila+3) ,startCol = "AV", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AU",ultima_fila+3,"/AU",fila_anterior+2,"*100-100") ,startCol = "AW", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AV",ultima_fila+3,"/AV",fila_anterior+2,"*100-100") ,startCol = "AX", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AX",ultima_fila+3,"-AS",ultima_fila+3) ,startCol = "AY", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("X",ultima_fila+3,"/Z",ultima_fila+3,"*100") ,startCol = "BB", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("Y",ultima_fila+3,"/Z",ultima_fila+3,"*100") ,startCol = "BD", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("SUM(BB",ultima_fila+3,":BD",ultima_fila+3,")") ,startCol = "BF", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BB",fila_anterior+2,"*AB",ultima_fila+3,"/100") ,startCol = "BH", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BD",fila_anterior+2,"*AC",ultima_fila+3,"/100") ,startCol = "BI", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("SUM(BH",ultima_fila+3,":BI",ultima_fila+3,")") ,startCol = "BJ", startRow = (ultima_fila[1]+3))
writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BJ",ultima_fila+3,"-AI",ultima_fila+3) ,startCol = "BK", startRow = (ultima_fila[1]+3))

writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("'Frutas Citricas'!AG",ultima_fila+3) ,startCol = "G", startRow = ultima_fila+11)

  addStyle(wb, sheet = "Frutas Citricas",style=col1,rows = (ultima_fila+3),cols = 1:3)
  addStyle(wb, sheet = "Frutas Citricas",style=col8,rows = (ultima_fila+3),cols = c(4,5,10,11))
  addStyle(wb, sheet = "Frutas Citricas",style=col4,rows = (ultima_fila+3),cols = c(7,8,13:63))

  trimestre=f_trimestre(mes)
  if(mes %in% c(3,6,9,12)){
    if(trimestre==1){
      fila=which(data[,45]==(anio-1))
      fila_f=fila+6
    }else{
      fila=which(data[,45]==(anio))
      fila_f=fila+1+trimestre
    }
    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AVERAGE(AG",(ultima_fila[1]+1),":AG",ultima_fila+3,")") ,startCol = "BO", startRow = (fila_f))
    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AVERAGE(AV",(ultima_fila[1]+1),":AV",ultima_fila+3,")") ,startCol = "BP", startRow = (fila_f))
    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BO",(fila_f[1]),"/BO",(fila_f[1]-4),"*100-100") ,startCol = "BQ", startRow = (fila_f))
    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BP",(fila_f[1]),"/BP",(fila_f[1]-4),"*100-100") ,startCol = "BR", startRow = (fila_f))
  }else{

  }

  if(mes==12){
    fila_anual=which(data$Anual==anio-1)
    writeData(wb, sheet = "Frutas Citricas", x = anio,colNames = FALSE,startCol = "BX", startRow = (fila_anual[1]+3))

    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("AVERAGE(AG",(ultima_fila[1]-11+3),":AG",ultima_fila+3,")") ,startCol = "BY", startRow = (fila_anual[1]+3))
    writeFormula(wb, sheet ="Frutas Citricas" , x = paste0("BY",(fila_anual[1]+3),"/BY",(fila_anual[1]+2),"*100-100") ,startCol = "BZ", startRow = (fila_anual[1]+3))
  }else{

  }

  # Otras Frutas ------------------------------------------------------------------

  #Leer solo la hoja de Frutas
  data <- read.xlsx(wb, sheet = "Otras frutas.", colNames = TRUE,startRow = 2)

  ultima_fila=nrow(data)
  fila=which(data$Año==2013)
  fila_anio_ant=which(data$Año== (anio-1))
  fila_2anio_ant=which(data$Año== (anio-2))
  fila_anterior=max(which(data$Mes== mes))



  # Escribe los datos en la hoja "Frutas Total(Expos+Interno)"
  writeData(wb, sheet = "Otras frutas.", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+3))
  writeData(wb, sheet = "Otras frutas.", x = exportaciones[,2],colNames = FALSE,startCol = "D", startRow = (fila_anio_ant[1]+2))
  writeData(wb, sheet = "Otras frutas.", x =as.numeric(consumo_interno[,2]),colNames = FALSE,startCol = "E", startRow = (fila[1]+2))
  writeData(wb, sheet = "Otras frutas.", x = exportaciones[,4],colNames = FALSE,startCol = "J", startRow = (fila_anio_ant[1]+2))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("D",ultima_fila+3,"/D",fila_anterior+2,"*100-100") ,startCol = "G", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("E",ultima_fila+3,"/E",fila_anterior+2,"*100-100") ,startCol = "H", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("E",ultima_fila+3,"*R",ultima_fila+3,"/100") ,startCol = "K", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("J",ultima_fila+3,"/J",fila_anterior+2,"*100-100") ,startCol = "M", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("K",ultima_fila+3,"/K",fila_anterior+2,"*100-100") ,startCol = "N", startRow = (ultima_fila[1]+3))
  writeData(wb, sheet = "Otras frutas.", x = exportaciones[,6],colNames = FALSE,startCol = "P", startRow = (fila_anio_ant[1]+2))
  writeData(wb, sheet = "Otras frutas.", x = IPP,colNames = FALSE,startCol = "R", startRow = (fila_2anio_ant[1]+2))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("P",ultima_fila+3,"/P",fila_anterior+2,"*100-100") ,startCol = "T", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("R",ultima_fila+3,"/R",fila_anterior+2,"*100-100") ,startCol = "V", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("((X3/12)/100)*(D",ultima_fila+3,"/AVERAGE(D125:D136)*100)") ,startCol = "X", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("((Y3/12)/100)*(E",ultima_fila+3,"/AVERAGE(E125:E136)*100)") ,startCol = "Y", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("SUM(X",ultima_fila+3,":Y",ultima_fila+3,")") ,startCol = "Z", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("X",ultima_fila+3,"/X",fila_anterior+2,"*100-100") ,startCol = "AB", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("Y",ultima_fila+3,"/Y",fila_anterior+2,"*100-100") ,startCol = "AC", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("Z",ultima_fila+3,"/Z",fila_anterior+2,"*100-100") ,startCol = "AD", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("Z",ultima_fila+3) ,startCol = "AG", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AG",ultima_fila+3,"/AG",fila_anterior+2,"*100-100") ,startCol = "AI", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AD",ultima_fila+3,"-AI",ultima_fila+3) ,startCol = "AJ", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("((AL3/12)/100)*(J",ultima_fila+3,"/AVERAGE(J125:J136)*100)") ,startCol = "AL", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("((AM3/12)/100)*(K",ultima_fila+3,"/AVERAGE(K125:K136)*100)") ,startCol = "AM", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("SUM(AL",ultima_fila+3,":AN",ultima_fila+3,")") ,startCol = "AO", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AL",ultima_fila+3,"/AL",fila_anterior+2,"*100-100") ,startCol = "AQ", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AM",ultima_fila+3,"/AM",fila_anterior+2,"*100-100") ,startCol = "AR", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AO",ultima_fila+3,"/AO",fila_anterior+2,"*100-100") ,startCol = "AS", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AL",ultima_fila+3) ,startCol = "AU", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AO",ultima_fila+3) ,startCol = "AV", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AU",ultima_fila+3,"/AU",fila_anterior+2,"*100-100") ,startCol = "AW", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AV",ultima_fila+3,"/AV",fila_anterior+2,"*100-100") ,startCol = "AX", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("AX",ultima_fila+3,"-AS",ultima_fila+3) ,startCol = "AY", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("X",ultima_fila+3,"/Z",ultima_fila+3,"*100") ,startCol = "BB", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("Y",ultima_fila+3,"/Z",ultima_fila+3,"*100") ,startCol = "BD", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("SUM(BB",ultima_fila+3,":BD",ultima_fila+3,")") ,startCol = "BF", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("BB",fila_anterior+2,"*AB",ultima_fila+3,"/100") ,startCol = "BH", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("BD",fila_anterior+2,"*AC",ultima_fila+3,"/100") ,startCol = "BI", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("SUM(BH",ultima_fila+3,":BI",ultima_fila+3,")") ,startCol = "BJ", startRow = (ultima_fila[1]+3))
  writeFormula(wb, sheet ="Otras frutas." , x = paste0("BJ",ultima_fila+3,"-AI",ultima_fila+3) ,startCol = "BK", startRow = (ultima_fila[1]+3))

  writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("'Otras frutas.'!AG",ultima_fila+3) ,startCol = "H", startRow = ultima_fila+11)

  addStyle(wb, sheet = "Otras frutas.",style=col1,rows = (ultima_fila+3),cols = 1:3)
  addStyle(wb, sheet = "Otras frutas.",style=col8,rows = (ultima_fila+3),cols = c(4,5,10,11))
  addStyle(wb, sheet = "Otras frutas.",style=col4,rows = (ultima_fila+3),cols = c(7,8,13:63))

  if(mes %in% c(3,6,9,12)){
    if(trimestre==1){
      fila=which(data[,45]==(anio-1))
      fila_f=fila+6
    }else{
      fila=which(data[,45]==(anio))
      fila_f=fila+1+trimestre
    }
    writeFormula(wb, sheet ="Otras frutas." , x = paste0("AVERAGE(AG",(ultima_fila[1]+1),":AG",ultima_fila+3,")") ,startCol = "BO", startRow = (fila_f))
    writeFormula(wb, sheet ="Otras frutas." , x = paste0("AVERAGE(AV",(ultima_fila[1]+1),":AV",ultima_fila+3,")") ,startCol = "BP", startRow = (fila_f))
    writeFormula(wb, sheet ="Otras frutas." , x = paste0("BO",(fila_f[1]),"/BO",(fila_f[1]-4),"*100-100") ,startCol = "BQ", startRow = (fila_f))
    writeFormula(wb, sheet ="Otras frutas." , x = paste0("BP",(fila_f[1]),"/BP",(fila_f[1]-4),"*100-100") ,startCol = "BR", startRow = (fila_f))
  }else{

  }

  if(mes==12){
    fila_anual=which(data$Anual==anio-1)
    writeData(wb, sheet = "Otras frutas.", x = anio,colNames = FALSE,startCol = "BX", startRow = (fila_anual[1]+3))

    writeFormula(wb, sheet ="Otras frutas." , x = paste0("AVERAGE(AG",(ultima_fila[1]-11+3),":AG",ultima_fila+3,")") ,startCol = "BY", startRow = (fila_anual[1]+3))
    writeFormula(wb, sheet ="Otras frutas." , x = paste0("BY",(fila_anual[1]+3),"/BY",(fila_anual[1]+2),"*100-100") ,startCol = "BZ", startRow = (fila_anual[1]+3))
  }else{

  }

# Fruto de Palma ------------------------------------------------------------------

  #Leer solo la hoja de Palma
  data <- read.xlsx(wb, sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-1))


  #Correr la funcion Palma
  valor_Palma=f_Palma(directorio,mes,anio)
  Palma_anterior=c(data[data$Año==(anio-2),"Frutode.palma"],valor_Palma$fruto[1:mes])
  tamaño=length(Palma_anterior)


  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(valor_Palma$fruto[(i-2):i]) / sum(Palma_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  Observaciones <-rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(valor_Palma$fruto[(i-11):i]) / sum(Palma_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"MILES DE TONELADAS"),
    Palma.Toneladas=valor_Palma$fruto[1:tamaño],
    Variacion.Anual=valor_Palma$fruto[1:tamaño]/Palma_anterior*100-100,
    Estado=as.numeric(Estado),
    observaciones=as.numeric(Observaciones),
    Tipo=rep("",tamaño)
  )




  # Escribe los datos en la hoja "Fruto de Palma"
  writeData(wb, sheet = "Fruto de Palma", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))

  writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("'Fruto de Palma'!E",ultima_fila+10) ,startCol = "E", startRow = ultima_fila+13)

  addStyle(wb, sheet = "Fruto de Palma",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Fruto de Palma",style=col4,rows = (ultima_fila+10),cols = 5:6)
  addStyle(wb, sheet = "Fruto de Palma",style=col3,rows = (ultima_fila+10),cols = 7)
  addStyle(wb, sheet = "Fruto de Palma",style=col6,rows = (ultima_fila+10),cols = 8)
  # Aceite de Palma ------------------------------------------------------------------

  #Leer solo la hoja de Palma
  data <- read.xlsx(wb, sheet = "Aceite de palma", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)

  Palma_anterior=c(data[data$Año==(anio-2),"Aceite.de.palma"],valor_Palma$aceite[1:mes])
  tamaño=length(Palma_anterior)


  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(valor_Palma$aceite[(i-2):i]) / sum(Palma_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  Observaciones <-rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(valor_Palma$aceite[(i-11):i]) / sum(Palma_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"MILES DE TONELADAS"),
    Palma.Toneladas=valor_Palma$aceite[1:tamaño],
    Variacion.Anual=valor_Palma$aceite[1:tamaño]/Palma_anterior*100-100,
    Estado=as.numeric(Estado),
    observaciones=as.numeric(Observaciones),
    Tipo=rep("",tamaño)
  )




  # Escribe los datos en la hoja "Aceite de Palma"
  writeData(wb, sheet = "Aceite de palma", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))


  addStyle(wb, sheet = "Aceite de palma",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Aceite de palma",style=col7,rows = (ultima_fila+10),cols = 5)
  addStyle(wb, sheet = "Aceite de palma",style=col4,rows = (ultima_fila+10),cols = 6)
  addStyle(wb, sheet = "Aceite de palma",style=col3,rows = (ultima_fila+10),cols = 7)
  addStyle(wb, sheet = "Aceite de palma",style=col6,rows = (ultima_fila+10),cols = 8)
  # Cacao ------------------------------------------------------------------

  #Leer solo la hoja de Cacaos
  data <- read.xlsx(wb, sheet = "Cacao", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))


  #Correr la funcion Pollos
  valor_Cacao=f_Cacao(directorio,mes,anio)
  valor_Cacao=na.omit(valor_Cacao)
  Cacao_anterior=c(data[data$Año==(anio-3),"Cacao"],valor_Cacao[1:(length(valor_Cacao)-12)])

  tamaño=length(Cacao_anterior)


  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(valor_Cacao[(i-2):i]) / sum(Cacao_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }

  Observaciones <- rep("",tamaño)

  for (i in seq(6, tamaño, by = 6)) {
    Observaciones[i] <- (sum(valor_Cacao[(i-5):i]) / sum(Cacao_anterior[(i-5):i]))*100-100  # Realiza la suma y división
  }


  Tipo <-rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Tipo[i] <- (sum(valor_Cacao[(i-11):i]) / sum(Cacao_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }

  if((mes+24)==tamaño){
    nuevos_datos <- data.frame(
      Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
      Año = c(data[fila[1]:ultima_fila,"Año"],anio),
      Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
      Descripcion=c(data[fila[1]:ultima_fila,"Descripcion"],"Toneladas"),
      Palma.Toneladas=valor_Cacao,
      Variacion.Anual=valor_Cacao/Cacao_anterior*100-100,
      Estado=as.numeric(Estado),
      observaciones=as.numeric(Observaciones),
      Tipo=as.numeric(Tipo)
    )
  }else{
    valor_Cacao=c(valor_Cacao,0)
    Cacao_anterior=c(Cacao_anterior,0)
    nuevos_datos <- data.frame(
      Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
      Año = c(data[fila[1]:ultima_fila,"Año"],anio),
      Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
      Descripcion=c(data[fila[1]:ultima_fila,"Descripcion"],"Toneladas"),
      Palma.Toneladas=valor_Cacao,
      Variacion.Anual=valor_Cacao/Cacao_anterior*100-100,
      Estado=as.numeric(c(Estado,"")),
      observaciones=as.numeric(c(Observaciones,"")),
      Tipo=as.numeric(c(Tipo,""))
    )
  }
  #Crear la nueva fila

  # Escribe los datos en la hoja "Cacao"
  writeData(wb, sheet = "Cacao", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))

if((mes+24)==tamaño){

}else{
  valor_meses=tail(which(mes==data$Periodicidad),3)
  writeFormula(wb, sheet ="Cacao" , x = paste0("AVERAGE(E",valor_meses[1]+9,",E",valor_meses[2]+9,",E",valor_meses[3]+9,")") ,startCol = "E", startRow = ultima_fila+10)
  writeFormula(wb, sheet ="Cacao" , x = paste0("E",ultima_fila+10,"/E",valor_meses[3]+9,"*100-100") ,startCol = "F", startRow = ultima_fila+10)

  }



  writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("'Cacao'!E",ultima_fila+10) ,startCol = "I", startRow = ultima_fila+13)

  addStyle(wb, sheet = "Cacao",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Cacao",style=col7,rows = (ultima_fila+10),cols = 5)
  addStyle(wb, sheet = "Cacao",style=col4,rows = (ultima_fila+10),cols = c(6:7,9))
  addStyle(wb, sheet = "Cacao",style=col3,rows = (ultima_fila+10),cols = 8)
  # Flores ------------------------------------------------------------------

  #Leer solo la hoja de Palma
  data <- read.xlsx(wb, sheet = "Flores", colNames = TRUE,startRow = 10)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))


  #Correr la funcion Palma
  valor_Flores=f_Flores(directorio,mes,anio)

  Flores_anterior=c(data[data$Año==(anio-3),"Flores"],valor_Flores[1:(length(valor_Flores)-12)])

  tamaño=length(Flores_anterior)


  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(valor_Flores[(i-2):i]) / sum(Flores_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }


  Observaciones <-rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(valor_Flores[(i-11):i]) / sum(Flores_anterior[(i-11):i]))*100-100  # Realiza la suma y división
  }
  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila[1]:ultima_fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = c(data[fila[1]:ultima_fila,"Año"],anio),
    Periodicidad=c(data[fila[1]:ultima_fila,"Periodicidad"],mes),
    Descripcion=c(data[fila[1]:ultima_fila,"Descripción"],"Exportaciones en miles de millones de pesos constantes"),
    Palma.Toneladas=valor_Flores,
    Variacion.Anual=valor_Flores/Flores_anterior*100-100,
    Estado=as.numeric(Estado),
    observaciones=as.numeric(Observaciones),
    Tipo=rep("",tamaño)
  )





  # Escribe los datos en la hoja "Flores"
  writeData(wb, sheet = "Flores", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+10))


  addStyle(wb, sheet = "Flores",style=col1,rows = (ultima_fila+11),cols = 1:4)
  addStyle(wb, sheet = "Flores",style=col4,rows = (ultima_fila+11),cols = 5:8)
  # Caña de Azucar ------------------------------------------------------------------

  #Leer solo la hoja de Palma
  data <- read.xlsx(wb, sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)
  fila=which(data$Año==anio)



  #Correr la funcion Palma
  valor_Caña_Azucar=f_Caña_azucar(directorio,mes,anio)
  valor_actual_caña=tail(lag(data$Caña.de.Azúcar,11),1)*(1+valor_Caña_Azucar$variacion/100)
  if(mes>length(valor_Caña_Azucar$vector)){
    vector_caña=c(valor_Caña_Azucar$vector,valor_actual_caña)

  }else{
    #valor_Caña_Azucar$vector[mes]=valor_actual_caña
    vector_caña=valor_Caña_Azucar$vector
  }


  Caña_anterior=tail(lag(data$Caña.de.Azúcar,11),mes)

  tamaño=length(Caña_anterior)

  Estado <- rep("",tamaño)
if(tamaño>2){
  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(vector_caña[(i-2):i]) / sum(Caña_anterior[(i-2):i]))*100-100  # Realiza la suma y división
  }
} else{

}



  nuevos_datos <- data.frame(
    Consecutivo = c(data[fila,"Consecutivo"],(data[ultima_fila, "Consecutivo"] + 1)),
    Año = rep(anio,mes),
    Periodicidad=c(1:mes),
    Descripcion="Toneladas",
    Leche.Toneladas=vector_caña,
    Variacion.Anual=vector_caña/Caña_anterior*100-100,
    Estado=as.numeric(Estado),
    observaciones=if (mes==12) {
      c(rep("",11),sum(vector_caña)/sum(Caña_anterior)*100-100)
    } else {
      rep("",mes)
    },
    Tipo=rep("",mes),
    adicional=c(rep("",(mes-1)),valor_Caña_Azucar$variacion)
  )


  # Escribe los datos en la hoja "Caña de Azúcar"
  writeData(wb, sheet = "Caña de Azúcar", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (fila[1]+9))

  writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("'Caña de Azúcar'!E",ultima_fila+10) ,startCol = "F", startRow = ultima_fila+13)

  addStyle(wb, sheet = "Caña de Azúcar",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Caña de Azúcar",style=col7,rows = (ultima_fila+10),cols = 5)
  addStyle(wb, sheet = "Caña de Azúcar",style=col4,rows = (ultima_fila+10),cols = 6:8)
  # Panela ------------------------------------------------------------------

  #Leer solo la hoja de Palma
  data <- read.xlsx(wb, sheet = "Panela", colNames = TRUE,startRow = 9)

  ultima_fila=nrow(data)
  fila=which(data$Año==(anio-2))


  #Correr la funcion Palma
  valor_Panela=f_Panela(directorio,mes,anio)
  data_Panela=data %>%
              filter(Año==(anio-2) |Año==(anio-1)) %>%
              select(Año,Periodicidad,Panela)
  data_Panela=data_Panela %>%
              group_by(Periodicidad)%>%
              summarise(promedio=mean(Panela))

  participacion_mes=data_Panela[mes,"promedio"]/sum(data_Panela$promedio)*100
  valor_actual=valor_Panela*(participacion_mes$promedio/100)

  #Crear la nueva fila
  nuevos_datos <- data.frame(
    Consecutivo = (data[ultima_fila, "Consecutivo"] + 1),
    Año = anio,
    Periodicidad=mes,
    Descripcion="Miles de Toneladas",
    Palma.Toneladas=valor_actual,
    Variacion.Anual=valor_actual/tail(lag(data$Panela,11),1)*100-100,
    Estado=if (mes %in% c(3, 6, 9, 12)) {
      (valor_actual+sum(data[(ultima_fila-1):ultima_fila,"Panela"]))/
        (sum(tail(lag(data$Panela,11),3)))*100-100

    } else {
      " "
    },
    observaciones=if (mes==12) {
      (valor_actual+sum(filter(data, Año == anio)[["Panela"]]))/
        (sum(tail(lag(data$Panela,11),12)))*100-100
    } else {
      " "
    },
    Tipo=""
  )





  # Escribe los datos en la hoja "Panela"
  writeData(wb, sheet = "Panela", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila+10))
  addStyle(wb, sheet = "Panela",style=col1,rows = (ultima_fila+10),cols = 1:4)
  addStyle(wb, sheet = "Panela",style=col7,rows = (ultima_fila+10),cols = 5)
  addStyle(wb, sheet = "Panela",style=col4,rows = (ultima_fila+10),cols = 6:8)
#  # Algodon ------------------------------------------------------------------
#
#  #Leer solo la hoja de Palma
#  data <- read.xlsx(wb, sheet = "Algodón Trimestral", colNames = TRUE,startRow = 11)
#
#  ultima_fila=nrow(data)
#
#
#
#  #Correr la funcion Palma
#  valor_Algodon=f_Algodon(directorio,mes,anio)
#  trimestre=f_trimestre(mes)
#  semestre=f_semestre(mes)
#  if (trimestre %in% c(1,3)){
#    valor_trimestre=tail(lag(data$Algodón,3),1)*(1+valor_Algodon/100)
#
#    nuevos_datos <- data.frame(
#      Consecutivo = (data[ultima_fila, "consecutivo"] + 1),
#      Año = anio,
#      Periodicidad=trimestre,
#      Descripcion="Toneladas",
#      Palma.Toneladas=valor_trimestre,
#      Variacion.Anual=valor_trimestre/tail(lag(data$Algodón,3),1)*100-100,
#      Estado=if (semestre==4) {
#        (valor_trimestre+sum(filter(data, Año == anio)[["Algodón"]]))/
#          (sum(tail(lag(data$Algodón,3),4)))*100-100
#      } else {
#        ""
#      },
#      observaciones="",
#      Tipo=""
#    )
#
#  }else{
#
#    valor_trimestre=tail(lag(data$Algodón,3),2)*(1+valor_Algodon/100)
#
#    nuevos_datos <- data.frame(
#      Consecutivo = c(data[ultima_fila, "consecutivo"],data[ultima_fila, "consecutivo"] + 1),
#      Año = c(data[ultima_fila, "Año"],anio),
#      Periodicidad=c(data[ultima_fila, "Periodicidad"],trimestre),
#      Descripcion=c(data[ultima_fila, "Descripcion"],"Toneladas"),
#      Palma.Toneladas=valor_trimestre,
#      Variacion.Anual=valor_trimestre/tail(lag(data$Algodón,3),2)*100-100,
#      Estado=if (semestre==4) {
#        (valor_trimestre+sum(filter(data, Año == anio)[["Algodón"]]))/
#          (sum(tail(lag(data$Algodón,3),4)))*100-100
#      } else {
#        ""
#      },
#      observaciones=rep("",2),
#      Tipo=rep("",2)
#    )
#  }
#
#
#
#  #Crear la nueva fila
#
#
#
#
#
#  # Escribe los datos en la hoja "Algodón Trimestral"
#  writeData(wb, sheet = "Algodón Trimestral", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila+12))
#
#  #Añadir estilos de celda
#  addStyle(wb, sheet = "Algodón Trimestral",style=col1,rows = (ultima_fila+12),cols = 1:4)
#  addStyle(wb, sheet = "Algodón Trimestral",style=col4,rows = (ultima_fila+12),cols = 5:7)
#
#

# Areas en desarrollo -----------------------------------------------------

data <- read.xlsx(wb, sheet = "Áreas en desarrollo", colNames = TRUE,startRow = 12)


ultima_fila=nrow(data)
fila_ant=which(data$Año==(anio-1) & data$Periodicidad==mes)

nuevos_datos <- data.frame(
  Consecutivo = data[ultima_fila,1]+1,
  Año = anio,
  Periodicidad=mes,
  Descripcion="hectáreas"
)


# Escribe los datos en la hoja "Flores"
writeData(wb, sheet="Áreas en desarrollo", x = nuevos_datos,colNames = FALSE,startCol = "A", startRow = (ultima_fila+13))




formulas <- c(paste0("E",ultima_fila+13,"/(SUM(E133:E144)/12)*100"),
                   paste0("F",ultima_fila+13,"/(SUM(F133:F144)/12)*100"),
                   paste0("G",ultima_fila+13,"/(SUM(G133:G144)/12)*100"),
                   paste0("H",ultima_fila+13,"/(SUM(H133:H144)/12)*100"),
                   paste0("I",ultima_fila+13,"/(SUM(I133:I144)/12)*100"),
                   paste0("(J",ultima_fila+13,"*J10)+(K",ultima_fila+13,"*K10)+(L",ultima_fila+13,"*L10)+(N",ultima_fila+13,"*N10)+(M",ultima_fila+13,"*M10)"),
                   paste0("O",ultima_fila+13,"/O",fila_ant+12,"*100-100")) ## skip header row

for (i in 10:16) {
  writeFormula(wb, sheet ="Áreas en desarrollo" , x = formulas[i-9] ,startCol = i, startRow = ultima_fila+13)
}
if (mes %in% c(3,6,9,12)){
  writeFormula(wb, sheet ="Áreas en desarrollo" , x = paste0("SUM(O",ultima_fila+11,":O",ultima_fila+13,")/SUM(O",fila_ant+10,":O",fila_ant+12,")*100-100") ,startCol = "Q", startRow = ultima_fila+13)
}


addStyle(wb, sheet = "Áreas en desarrollo",style=col1,rows = (ultima_fila+13),cols = 1:4)
addStyle(wb, sheet = "Áreas en desarrollo",style=col7,rows = (ultima_fila+13),cols = c(5:7,9:11,14))
addStyle(wb, sheet = "Áreas en desarrollo",style=col6,rows = (ultima_fila+13),cols = 15)
addStyle(wb, sheet = "Áreas en desarrollo",style=col2,rows = (ultima_fila+13),cols = c(8,12:13))
addStyle(wb, sheet = "Áreas en desarrollo",style=col3,rows = (ultima_fila+13),cols = 16)
addStyle(wb, sheet = "Áreas en desarrollo",style=col4,rows = (ultima_fila+13),cols = 17)



# Guardar el libro --------------------------------------------------------

if (!file.exists(salida)) {
  saveWorkbook(wb, file = salida)
} else {
  saveWorkbook(wb, file = salida,overwrite= TRUE)
}


}
