#' @export
# Flores
# Cargar la biblioteca readxl

f_Flores_complemento<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)



  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")


  #IPP

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPP","NOMBRE"]

  IPP <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                   sheet = "5.1")

  #Identificar la fila donde esta la palabra Periodo
  n_col=which(IPP == "CODIGO",arr.ind = TRUE)[, "col"]

  dos_digitos=anio %% 100
  #identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie
  columna1=max(which(grepl(paste0(nombres_siglas[1],"-",(dos_digitos-2)),IPP),arr.ind = TRUE))
  columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),IPP),arr.ind = TRUE)

  fila1=which(grepl("01962",IPP[,n_col]),arr.ind = TRUE)


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


  #rosas
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Flores <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                     sheet = "PNK")

  n_fila=which(Flores == "011101",arr.ind = TRUE)[,"row"]
  n_col1=which(Flores== paste0((anio-2)," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Flores== paste0((anio)," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Rosas=as.numeric(Flores[n_fila[1],(n_col1[1]:n_col2[1])])
  var_anual=Rosas/lag(Rosas,12)*100-100
  Rosas_ant=lag(Rosas,12)
  tamaño=length(Rosas)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Rosas[(i-2):i]) / sum(Rosas_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Rosas[(i-11):i]) / sum(Rosas_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_Rosas=data.frame(var_anual[c(12+mes)],var_anual[c(24+mes)],Estado[c(12+mes)],
                                  Estado[c(24+mes)],Observaciones[c(12+mes)],Observaciones[c(24+mes)])


#Claveles

  n_fila=which(Flores == "011102",arr.ind = TRUE)[,"row"]
  n_col1=which(Flores== paste0((anio-2)," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Flores== paste0((anio)," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Claveles=as.numeric(Flores[n_fila[1],(n_col1[1]:n_col2[1])])
  var_anual=Claveles/lag(Claveles,12)*100-100
  Claveles_ant=lag(Claveles,12)
  tamaño=length(Claveles)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Claveles[(i-2):i]) / sum(Claveles_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Claveles[(i-11):i]) / sum(Claveles_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_Claveles=data.frame(var_anual[c(12+mes)],var_anual[c(24+mes)],Estado[c(12+mes)],
                          Estado[c(24+mes)],Observaciones[c(12+mes)],Observaciones[c(24+mes)])


  #Pompones

  n_fila=which(Flores == "011103",arr.ind = TRUE)[,"row"]
  n_col1=which(Flores== paste0((anio-2)," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Flores== paste0((anio)," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Pompones=as.numeric(Flores[n_fila[1],(n_col1[1]:n_col2[1])])
  var_anual=Pompones/lag(Pompones,12)*100-100
  Pompones_ant=lag(Pompones,12)
  tamaño=length(Pompones)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Pompones[(i-2):i]) / sum(Pompones_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Pompones[(i-11):i]) / sum(Pompones_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_Pompones=data.frame(var_anual[c(12+mes)],var_anual[c(24+mes)],Estado[c(12+mes)],
                             Estado[c(24+mes)],Observaciones[c(12+mes)],Observaciones[c(24+mes)])


  nuevos_datos=bind_rows(cuadro_Rosas,cuadro_Claveles,cuadro_Pompones,cuadro_Valor_IPP)
  return(nuevos_datos)
}
