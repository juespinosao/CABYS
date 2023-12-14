

f_Fenavi<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)



# Encasetamiento ----------------------------------------------------------



  # Especifica la ruta del archivo de Excel
  Encasetamiento_pollito <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENAVI/Encasetamiento",anio,".xlsx"),sheet = "POLLITO")

  n_fila=which(Encasetamiento_pollito == "MES",arr.ind = TRUE)[, "row"]
  fila_tabla1=as.numeric(which(Encasetamiento_pollito == "ENERO",arr.ind = TRUE)[, "row"])
  fila_tabla2=as.numeric(which(Encasetamiento_pollito == "SUBTOTAL",arr.ind = TRUE)[, "row"])
  columna=which(grepl(anio,Encasetamiento_pollito[n_fila,]),arr.ind = TRUE)
  Tabla1=as.numeric(Encasetamiento_pollito[fila_tabla1:(fila_tabla2-1),columna])


  Encasetamiento_pollita <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENAVI/Encasetamiento",anio,".xlsx"),sheet = "POLLITA")

  n_fila=which(Encasetamiento_pollita == "MES",arr.ind = TRUE)[, "row"]
  fila_tabla1=as.numeric(which(Encasetamiento_pollita == "ENERO",arr.ind = TRUE)[, "row"])
  fila_tabla2=as.numeric(which(Encasetamiento_pollita == "SUBTOTAL",arr.ind = TRUE)[, "row"])
  columna=which(grepl(anio,Encasetamiento_pollita[n_fila,]),arr.ind = TRUE)
  Tabla2=as.numeric(Encasetamiento_pollita[fila_tabla1:(fila_tabla2-1),columna])



# Aves de postura ---------------------------------------------------------



  # Especifica la ruta del archivo de Excel
  Postura <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/FENAVI/Invetarios-de-aves-",tolower(nombres_siglas[mes]),"-",anio,".xlsx"))
 fila_col=which(grepl("Mes",as.data.frame(t(Postura))),arr.ind = TRUE)
  n_fila=which(grepl("postura",as.data.frame(t(Postura))),arr.ind = TRUE)
  fila_tabla=as.numeric(which(Postura == "Ene",arr.ind = TRUE)[, "row"])
  des=which.min(abs(n_fila-fila_tabla))
  columna=which(grepl(anio,Postura[fila_col[1],]),arr.ind = TRUE)


  Tabla=as.data.frame(Postura[fila_tabla[des]:(fila_tabla[des]+mes-1),columna])


  Valor_Postura=as.numeric(Tabla[,1])





# agrupar datos -----------------------------------------------------------
  longitud_maxima <- max(length(Tabla1), length(Tabla2), length(Valor_Postura))
  Tabla1 <- c(Tabla1, rep(NA, longitud_maxima - length(Tabla1)))
  Tabla2 <- c(Tabla2, rep(NA, longitud_maxima - length(Tabla2)))
  Valor_Postura <- c(Valor_Postura, rep(NA, longitud_maxima - length(Valor_Postura)))

  valor_Fenavi=data.frame(Tabla1,Tabla2,Valor_Postura)





  # Importaciones -----------------------------------------------------------

  archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]


  dos_digitos <- anio %% 100
  # Especifica la ruta del archivo de Excel
  Impor <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Importaciones ",mes_0[mes],"_",dos_digitos," - copia.xlsx"),
                     sheet = "PNK")

  n_col1=which(Impor== paste0((anio)," 01"),arr.ind = TRUE)[,"col"]
  n_col2=which(Impor== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]
  n_fila1=which(Impor== "020502",arr.ind = TRUE)[,"row"]
  n_fila2=which(Impor == "020400",arr.ind = TRUE)[,"row"]
  n_fila3=which(Impor == "210105",arr.ind = TRUE)[,"row"]
  n_fila4=which(Impor == "010102",arr.ind = TRUE)[,"row"]
  n_fila5=which(Impor == "210600",arr.ind = TRUE)[,"row"]

  #Tomar el valor que nos interesa
  Valor_importaciones=as.data.frame(t(Impor[c(n_fila1,n_fila2,n_fila3,n_fila4,n_fila5),(n_col1[1]:n_col2[1])]))
  Valor_importaciones=as.data.frame(lapply(Valor_importaciones, as.numeric))





  # Exportaciones -----------------------------------------------------------



  # Especifica la ruta del archivo de Excel
  # Especifica la ruta del archivo de Excel
  Expor <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Exportaciones ",mes_0[mes],"-",anio," - copia.xlsx"),
                      sheet = "PNK")


  n_fila1=which(Expor == "020502",arr.ind = TRUE)[,"row"]
  n_col1=which(Expor== paste0(anio," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Expor== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

  n_fila2=which(Expor == "020400",arr.ind = TRUE)[,"row"]
  n_fila3=which(Expor == "210105",arr.ind = TRUE)[,"row"]


  #Tomar el valor que nos interesa
  Valor_exportaciones=as.data.frame(t(Expor[c(n_fila1,n_fila2,n_fila3),(n_col1[1]:n_col2[1])]))
  Valor_exportaciones=as.data.frame(lapply(Valor_exportaciones, as.numeric))

  valor_Fenavi=cbind(valor_Fenavi,Valor_importaciones,Valor_exportaciones)
  colnames(valor_Fenavi)=c("POLLITOS","POLLITAS","AVES_POSTURA","IMPO_HUEVOS","AVES_GALLUS","CARNE_DESPOJOS","IMPO_MAIZ","RESIDUOS_GRASA","EXPO_HUEVOS","EXPO_GALLOS","EXPO_CARNE_DESPOJOS")



  return(valor_Fenavi)
}
