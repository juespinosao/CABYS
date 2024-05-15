#' @export
# Yuca


f_Yuca_complemento<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")

  #IPP

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

  fila1=which(grepl("01592",IPP[,n_col]),arr.ind = TRUE)


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

  #IPC

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPC_HIST","NOMBRE"]
  hojas=excel_sheets(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo))
  dos_digitos=anio %% 100
  tamaño=36+mes
  hoja_final <- hojas[grepl(dos_digitos, hojas) ]

  IPC <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                   sheet = hoja_final)


  n_col1=which(IPC == (anio-3),arr.ind=TRUE)[, "col"]


  fila1=which(grepl("01170600",IPC[,1]),arr.ind = TRUE)

  #Tomar el valor que nos interesa
  Valor_IPC=as.data.frame(t(IPC[c(fila1[1]),c(n_col1[1]:(n_col1[1]+tamaño-1))]))
  tamaño=nrow(Valor_IPC)
  Valor_IPC=as.numeric(Valor_IPC[,1])

  var_anual=Valor_IPC/lag(Valor_IPC,12)*100-100
  Valor_IPC_ant=lag(Valor_IPC,12)
  tamaño=length(Valor_IPC)
  Estado <- rep("",tamaño)

  for (i in seq(3, tamaño, by = 3)) {
    Estado[i] <- (sum(Valor_IPC[(i-2):i]) / sum(Valor_IPC_ant[(i-2):i]))*100-100  # Realiza la suma y división
  }
  Estado=as.numeric(Estado)
  Observaciones <- rep("",tamaño)

  for (i in seq(12, tamaño, by = 12)) {
    Observaciones[i] <- (sum(Valor_IPC[(i-11):i]) / sum(Valor_IPC_ant[(i-11):i]))*100-100  # Realiza la suma y división
  }
  Observaciones=as.numeric(Observaciones)

  cuadro_Valor_IPC=data.frame(var_anual[c(24+mes)],var_anual[c(36+mes)],Estado[c(24+mes)],
                              Estado[c(36+mes)],Observaciones[c(24+mes)],Observaciones[c(36+mes)])


  colnames(cuadro_Valor_IPP)=colnames(cuadro_Valor_IPC)
  nuevos_datos=bind_rows(cuadro_Valor_IPP,cuadro_Valor_IPC)
  return(nuevos_datos)
}
