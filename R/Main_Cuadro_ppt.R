#' @export


Cuadros_ppt<-function(directorio,mes,anio){

  #Cargar librerias
  library(openxlsx)
  library(dplyr)
  #Crear el nombre de las carpetas del mes anterior y el actual
  carpeta_actual=nombre_carpeta(mes,anio)
  if(mes==1){
    carpeta_anterior=nombre_carpeta(12,(anio-1))
    entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/Cuadros_Agropecuario_",nombres_meses[12],"_",(anio-1),".xlsx")

  }else{
    carpeta_anterior=nombre_carpeta(mes-1,anio)
    entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/Cuadros_Agropecuario_",nombres_meses[mes-1],"_",anio,".xlsx")

  }

  carpeta_actual=nombre_carpeta(mes,anio)

  #Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
  salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/Cuadros_Agropecuario_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb <- loadWorkbook(entrada)


  funcion_cuadro=function(nombre){
    exportacion=as.numeric(data[c(fila[1]:nrow(data)),nombre])
    var_anual=exportacion/lag(exportacion,12)*100-100
    exportacion_ant=lag(exportacion,12)
    tamaño=length(exportacion)
    Estado <- rep("",tamaño)

    for (i in seq(3, tamaño, by = 3)) {
      Estado[i] <- (sum(exportacion[(i-2):i]) / sum(exportacion_ant[(i-2):i]))*100-100  # Realiza la suma y división
    }
    Estado=as.numeric(Estado)
    Observaciones <- rep("",tamaño)

    for (i in seq(12, tamaño, by = 12)) {
      Observaciones[i] <- (sum(exportacion[(i-11):i]) / sum(exportacion_ant[(i-11):i]))*100-100  # Realiza la suma y división
    }
    Observaciones=as.numeric(Observaciones)

    cuadro_expo=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                           Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])
    return(cuadro_expo)
  }
  funcion_cuadro2=function(nombre){
    exportacion=as.numeric(data[c(fila[1]:nrow(data)),nombre])
    exportacion_ant=lag(exportacion,12)
    tamaño=length(exportacion)
    Estado <- rep("",tamaño)

    for (i in seq(3, tamaño, by = 3)) {
      Estado[i] <- (sum(exportacion[(i-2):i]) / sum(exportacion_ant[(i-2):i]))*100-100  # Realiza la suma y división
    }
    Estado=as.numeric(Estado)
    Observaciones <- rep("",tamaño)

    for (i in seq(12, tamaño, by = 12)) {
      Observaciones[i] <- (sum(exportacion[(i-11):i]) / sum(exportacion_ant[(i-11):i]))*100-100  # Realiza la suma y división
    }

    Observaciones=as.numeric(Observaciones)
    var_anual=exportacion
    cuadro_expo=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                           Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])
    return(cuadro_expo)
  }
  trim_rom=f_trim_rom(mes)
  semestre_nombre=f_semestre_nombre(mes)
  # Café -----------------------------------------------------------


  writeData(wb, sheet = "Café", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Café", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Café", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Café", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Café", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Café", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))

  cuadro_expo=funcion_cuadro("Exportaciones.Totales")
  cuadro_impo=funcion_cuadro("Importaciones.Totales")
  cuadro_consumo=funcion_cuadro("Consumo.Intermedio")
  cuadro_existencias_v=funcion_cuadro2("VARIACIÓN.Existencias.de.café.verde.miles.sacos")
  cuadro_existencias_p=funcion_cuadro2("Existencias.de.pergamino.Miles.Sacos")
  cuadro_produccion_v=funcion_cuadro("Producción.café.verde")
  cuadro_produccion_p=funcion_cuadro("Producción.Total.de.Café.Pergamino")
  cuadro_precios_interno=f_Cafe_precio_interno_ppt(directorio,mes,anio)
  cuadro_precios_internacional=f_Cafe_precio_internacional_ppt(directorio,mes,anio)
  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Cafetos", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))
  cuadro_cafetos=funcion_cuadro("Cafetos")
  cuadro_produccion_cafetos=f_produccion_cafetos(directorio,mes,anio)

nuevos_datos=bind_rows(cuadro_expo,cuadro_impo,cuadro_consumo,cuadro_existencias_v,cuadro_existencias_p,
                   cuadro_produccion_v,cuadro_produccion_p,cuadro_precios_interno,cuadro_precios_internacional,
                   cuadro_cafetos,cuadro_produccion_cafetos)
writeData(wb, sheet = "Café", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)








  # Maiz -----------------------------------------------------------


  writeData(wb, sheet = "Maiz", x = paste0(semestre_nombre," semestre ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Maiz", x = paste0(semestre_nombre," semestre ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Maiz", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Maiz", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Maiz", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Maiz", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)

  writeData(wb, sheet = "Maiz", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 7)
  writeData(wb, sheet = "Maiz", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 7)
  writeData(wb, sheet = "Maiz", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 7)
  writeData(wb, sheet = "Maiz", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 7)
  writeData(wb, sheet = "Maiz", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 7)
  writeData(wb, sheet = "Maiz", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 7)

  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Maíz", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Maiz")
  cuadro_productos=f_producto_maiz(directorio,mes,anio)
  colnames(cuadro_productos)=colnames(cuadro_produccion)
  cuadro_inferior=f_maiz_complemento(directorio,mes,anio)
  cuadro_importaciones=cuadro_inferior[[1]]
  cuadro_ipp=cuadro_inferior[[2]]
  colnames(cuadro_ipp)=colnames(cuadro_produccion)
  cuadro_precio=cuadro_inferior[[3]]

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_productos)
  writeData(wb, sheet = "Maiz", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)
  writeData(wb, sheet = "Maiz", x = cuadro_importaciones,colNames = FALSE,startCol = "D", startRow = 8)
  nuevos_datos2=bind_rows(cuadro_ipp,cuadro_precio)
  writeData(wb, sheet = "Maiz", x = nuevos_datos2,colNames = FALSE,startCol = "D", startRow = 11)
  setRowHeights(wb,sheet ="Maiz",rows = c(9,10),heights = 0)



  # Arroz -----------------------------------------------------------


  writeData(wb, sheet = "Arroz", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Arroz", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Arroz", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Arroz", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Arroz", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Arroz", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


    data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Arroz", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Arroz")
  cuadro_inferior=f_arroz_complemento(directorio,mes,anio)

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_inferior)
  writeData(wb, sheet = "Arroz", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)

  setRowHeights(wb,sheet ="Arroz",rows = c(10,11),heights = 0)


  # Papa -----------------------------------------------------------


  writeData(wb, sheet = "Papa", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Papa", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Papa", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Papa", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Papa", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Papa", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Papa", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Papa")
  cuadro_inferior=f_Papa_complemento(directorio,mes,anio)

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_inferior)
  writeData(wb, sheet = "Papa", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)

  # Hortalizas -----------------------------------------------------------


  writeData(wb, sheet = "Hortalizas", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Hortalizas", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Hortalizas", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Hortalizas", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Hortalizas", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Hortalizas", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Hortalizas", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Hortalizas")
  cuadro_inferior=f_Hortalizas_complemento(directorio,mes,anio)

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_inferior)
  writeData(wb, sheet = "Hortalizas", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 5)

  setRowHeights(wb,sheet ="Hortalizas",rows = c(8:11),heights = 0)

  # Yuca -----------------------------------------------------------


  writeData(wb, sheet = "Yuca", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Yuca", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Yuca", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Yuca", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Yuca", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Yuca", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Yuca", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Yuca")
  cuadro_inferior=f_Yuca_complemento(directorio,mes,anio)

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_inferior)
  writeData(wb, sheet = "Yuca", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 9)

  setRowHeights(wb,sheet ="Yuca",rows = c(4:7),heights = 0)

  # Frijol -----------------------------------------------------------


  writeData(wb, sheet = "Frijol", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Frijol", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Frijol", x = paste0(trim_rom," ",anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Frijol", x = paste0(trim_rom," ",anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Frijol", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Frijol", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Legumbres", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))

  cuadro_produccion=funcion_cuadro("Legumbres.verdes.y.secas.(frijoles,.arvejas,.habas,.garbanzos,.lentejas,.etc.)")
  cuadro_inferior=f_Frijol_complemento(directorio,mes,anio)

  nuevos_datos=bind_rows(cuadro_produccion,cuadro_inferior)
  writeData(wb, sheet = "Frijol", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)


  # Banano -----------------------------------------------------------


  writeData(wb, sheet = "Banano", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Banano", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Banano", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Banano", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Banano", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Banano", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data$Año==(anio-3))


  cuadro_produccion=funcion_cuadro("Banano.consumo.interno(SIPSA).ton")
  cuadro_inferior=f_Banano_complemento(directorio,mes,anio)
  cuadro_exportaciones=funcion_cuadro("Banano.de.Exportación.(DANE).ktes")
  cuadro_total=funcion_cuadro2("Variación.anual.Banano.total")

  writeData(wb, sheet = "Banano", x = cuadro_produccion,colNames = FALSE,startCol = "D", startRow = 5)
  writeData(wb, sheet = "Banano", x = cuadro_inferior[[1]],colNames = FALSE,startCol = "D", startRow = 6)
  writeData(wb, sheet = "Banano", x = cuadro_exportaciones,colNames = FALSE,startCol = "D", startRow = 9)
  writeData(wb, sheet = "Banano", x = cuadro_inferior[[2]],colNames = FALSE,startCol = "D", startRow = 10)
  writeData(wb, sheet = "Banano", x = cuadro_total,colNames = FALSE,startCol = "D", startRow = 12)


  # Platano -----------------------------------------------------------


  writeData(wb, sheet = "Platano", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Platano", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Platano", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Platano", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Platano", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Platano", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data$Año==(anio-3))


  cuadro_produccion=funcion_cuadro("Plátano.consumo.interno(SIPSA).ton")
  cuadro_exportaciones=funcion_cuadro("Plátano.de.Exportación")
  cuadro_inferior=f_Platano_complemento(directorio,mes,anio)
  cuadro_total=funcion_cuadro2("Variación.anual.plátano.total")
  nuevos_datos=bind_rows(cuadro_produccion,cuadro_exportaciones,cuadro_inferior,cuadro_total)

  writeData(wb, sheet = "Platano", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 5)

  # Frutas citricas -----------------------------------------------------------


  writeData(wb, sheet = "Frutas citricas", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Frutas citricas", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Frutas citricas", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Frutas citricas", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Frutas citricas", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Frutas citricas", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)
  fila=which(data$Año==(anio-3))


  cuadro_produccion=funcion_cuadro("SIPSA")
  cuadro_exportaciones=funcion_cuadro("Expos.Ktes")
  cuadro_frutas=f_Frutas_complemento(directorio,mes,anio)
  cuadro_total=funcion_cuadro("Frutas.citricas.Ktes.2.272")
  nuevos_datos=bind_rows(cuadro_produccion,cuadro_exportaciones,cuadro_frutas,cuadro_total)

  writeData(wb, sheet = "Frutas citricas", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 5)


  # Otras frutas -----------------------------------------------------------


  writeData(wb, sheet = "Otras frutas", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Otras frutas", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Otras frutas", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Otras frutas", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Otras frutas", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Otras frutas", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Otras frutas.", colNames = TRUE,startRow = 2)
  fila=which(data$Año==(anio-3))


  cuadro_produccion=funcion_cuadro("SIPSA")
  cuadro_exportaciones=funcion_cuadro("Expos.Ktes")
  cuadro_total=funcion_cuadro("Frutas.citricas.Ktes.7.516")
  nuevos_datos=bind_rows(cuadro_produccion,cuadro_exportaciones,cuadro_frutas,cuadro_total)

  writeData(wb, sheet = "Otras frutas", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 5)

  # Fruto de palma -----------------------------------------------------------


  writeData(wb, sheet = "Fruto de palma", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Fruto de palma", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Fruto de palma", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Fruto de palma", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Fruto de palma", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Fruto de palma", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))

  cuadro_fruto=funcion_cuadro("Frutode.palma")
  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Aceite de palma", colNames = TRUE,startRow = 9)
  cuadro_aceite=funcion_cuadro("Aceite.de.palma")
  cuadro_palma=f_Palma_complemento(directorio,mes,anio)
  nuevos_datos=bind_rows(cuadro_fruto,cuadro_aceite,cuadro_palma)

  writeData(wb, sheet = "Fruto de palma", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)


  # Cacao -----------------------------------------------------------


  writeData(wb, sheet = "Cacao", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Cacao", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Cacao", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Cacao", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Cacao", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Cacao", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Cacao", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))
  cuadro_cacao=funcion_cuadro("Cacao")
  cuadro_inferior=f_Cacao_complemento(directorio,mes,anio)
  nuevos_datos=bind_rows(cuadro_cacao,cuadro_inferior)

  writeData(wb, sheet = "Cacao", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)


  # Flores -----------------------------------------------------------


  writeData(wb, sheet = "Flores", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Flores", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Flores", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Flores", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Flores", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Flores", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Flores", colNames = TRUE,startRow = 10)
  fila=which(data$Año==(anio-3))
  cuadro_Flores=funcion_cuadro("Flores")
  cuadro_inferior=f_Flores_complemento(directorio,mes,anio)
  colnames(cuadro_Flores)=colnames(cuadro_inferior)
  nuevos_datos=bind_rows(cuadro_Flores,cuadro_inferior)

  writeData(wb, sheet = "Flores", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)


  # Caña de azucar -----------------------------------------------------------


  writeData(wb, sheet = "Caña de azucar", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Caña de azucar", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Caña de azucar", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Caña de azucar", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Caña de azucar", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Caña de azucar", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))
  cuadro_Caña_azucar=funcion_cuadro("Caña.de.Azúcar")
  cuadro_inferior=f_Caña_azucar_complemento(directorio,mes,anio)
  nuevos_datos=bind_rows(cuadro_Caña_azucar,cuadro_inferior)

  writeData(wb, sheet = "Caña de azucar", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)


  # Panela -----------------------------------------------------------


  writeData(wb, sheet = "Panela", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Panela", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Panela", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Panela", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "G", startRow = 3)
  writeData(wb, sheet = "Panela", x = paste0(anio-1),colNames = FALSE,startCol = "H", startRow = 3)
  writeData(wb, sheet = "Panela", x = paste0(anio),colNames = FALSE,startCol = "I", startRow = 3)


  data <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx"), sheet = "Panela", colNames = TRUE,startRow = 9)
  fila=which(data$Año==(anio-3))
  cuadro_Panela=funcion_cuadro("Panela")
  cuadro_inferior=f_Panela_complemento(directorio,mes,anio)
  nuevos_datos=bind_rows(cuadro_Panela,cuadro_inferior)

  writeData(wb, sheet = "Panela", x = nuevos_datos,colNames = FALSE,startCol = "D", startRow = 4)

  # Pesca -----------------------------------------------------------

  if(mes %in% c(3,6,9,12)){
  writeData(wb, sheet = "Pesca", x = paste0(trim_rom,anio-1," / ",trim_rom,anio-2),colNames = FALSE,startCol = "D", startRow = 3)
  writeData(wb, sheet = "Pesca", x = paste0(trim_rom,anio," / ",trim_rom,anio-1),colNames = FALSE,startCol = "E", startRow = 3)
  writeData(wb, sheet = "Pesca", x = paste0(anio-1),colNames = FALSE,startCol = "F", startRow = 3)
  writeData(wb, sheet = "Pesca", x = paste0(anio),colNames = FALSE,startCol = "G", startRow = 3)



  cuadro_inferior=f_Pesca_complemento(directorio,mes,anio)

  writeData(wb, sheet = "Pesca", x = cuadro_inferior,colNames = FALSE,startCol = "D", startRow = 5)
  setRowHeights(wb,sheet ="Pesca",rows = c(4),heights = 0)
  }else{

}
  # Guardar el libro --------------------------------------------------------


  if (!file.exists(salida)) {
    saveWorkbook(wb, file = salida)
  } else {
    saveWorkbook(wb, file = salida,overwrite= TRUE)
  }
}
