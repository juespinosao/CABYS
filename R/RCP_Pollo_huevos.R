#' @export

f_Fenavi<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)
  #utils

  carpeta=nombre_carpeta(mes,anio)



# Encasetamiento ----------------------------------------------------------

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="FENAVI2","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Encasetamiento_pollito <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENAVI/",archivo),sheet = "POLLITO")

  n_fila=which(Encasetamiento_pollito == "MES",arr.ind = TRUE)[, "row"]
  fila_tabla1=as.numeric(which(Encasetamiento_pollito == "ENERO",arr.ind = TRUE)[, "row"])
  fila_tabla2=as.numeric(which(Encasetamiento_pollito == "SUBTOTAL"|Encasetamiento_pollito == "TOTAL",arr.ind = TRUE)[, "row"])
  columna=which(grepl(anio,Encasetamiento_pollito[n_fila,]),arr.ind = TRUE)
  columna <- which(grepl(paste0((anio-1), "|", (anio)), Encasetamiento_pollito[n_fila, ]), arr.ind = TRUE)
  Tabla1=Encasetamiento_pollito[c(fila_tabla1:(fila_tabla2[2]-1)),columna[1]]
  Tabla1=Tabla1[-c((fila_tabla2[1]-fila_tabla1[1]+1))]
  Tabla2=Encasetamiento_pollito[fila_tabla1:(fila_tabla2[1]-1),columna[2]]

  vector1=na.omit(c(Tabla1,Tabla2))

  Encasetamiento_pollita <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENAVI/",archivo),sheet = "POLLITA")

  n_fila=which(Encasetamiento_pollita == "MES",arr.ind = TRUE)[, "row"]
  fila_tabla1=as.numeric(which(Encasetamiento_pollita == "ENERO",arr.ind = TRUE)[, "row"])
  fila_tabla2=as.numeric(which(Encasetamiento_pollita == "SUBTOTAL" |Encasetamiento_pollita == "TOTAL",arr.ind = TRUE)[, "row"])
  columna <- which(grepl(paste0((anio-1), "|", (anio)), Encasetamiento_pollita[n_fila, ]), arr.ind = TRUE)

  Tabla1=Encasetamiento_pollita[c(fila_tabla1:(fila_tabla2[2]-1)),columna[1]]
  Tabla1=Tabla1[-c((fila_tabla2[1]-fila_tabla1[1]+1))]
  Tabla2=Encasetamiento_pollita[fila_tabla1:(fila_tabla2[1]-1),columna[2]]

  vector2=na.omit(c(Tabla1,Tabla2))

  vector2=na.omit(vector2)


# Aves de postura ---------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="FENAVI3","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Postura <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/FENAVI/",archivo))
 fila_col=which(grepl("Mes",as.data.frame(t(Postura))),arr.ind = TRUE)
  n_fila=which(grepl("postura",as.data.frame(t(Postura))),arr.ind = TRUE)
  fila_tabla=as.numeric(which(Postura == "Ene",arr.ind = TRUE)[, "row"])
  fila_tabla2=as.numeric(which(Postura == "Dic",arr.ind = TRUE)[, "row"])
  des=which.min(abs(n_fila-fila_tabla))
  des2=which.max(abs(n_fila-fila_tabla2))
  columna <- which(grepl(paste0((anio-1), "|", (anio)), Postura[fila_col[1],]), arr.ind = TRUE)


  Tabla3=as.data.frame(Postura[fila_tabla[des]:(fila_tabla2[des2]),columna])


  vector3=NULL
  for (i in 1:ncol(Tabla3)) {
    vector3=c(vector3,as.vector(Tabla3[,i]))
  }
  vector3=na.omit(vector3)



# agrupar datos -----------------------------------------------------------
  longitud_maxima <- max(length(vector1), length(vector2), length(vector3))
  vector1 <- c(vector1, rep(NA, longitud_maxima - length(vector1)))
  vector2 <- c(vector2, rep(NA, longitud_maxima - length(vector2)))
  vector3 <- c(vector3, rep(NA, longitud_maxima - length(vector3)))

  valor_Fenavi=data.frame(vector1,vector2,vector3)





  # Importaciones -----------------------------------------------------------

  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Importaciones","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Impor <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                     sheet = "PNK")

  n_col1=which(Impor== paste0((anio-1)," 01"),arr.ind = TRUE)[,"col"]
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

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]

  Expor <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "PNK")


  n_fila1=which(Expor == "020502",arr.ind = TRUE)[,"row"]
  n_col1=which(Expor== paste0(anio-1," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Expor== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

  n_fila2=which(Expor == "020400",arr.ind = TRUE)[,"row"]
  n_fila3=which(Expor == "210105",arr.ind = TRUE)[,"row"]


  #Tomar el valor que nos interesa
  Valor_exportaciones=as.data.frame(t(Expor[c(n_fila1,n_fila2,n_fila3),(n_col1[1]:n_col2[1])]))
  Valor_exportaciones=as.data.frame(lapply(Valor_exportaciones, as.numeric))
valor_Fenavi=valor_Fenavi[1:nrow(Valor_exportaciones),]

  valor_Fenavi=cbind(valor_Fenavi,Valor_importaciones,Valor_exportaciones)
  colnames(valor_Fenavi)=c("POLLITOS","POLLITAS","AVES_POSTURA","IMPO_HUEVOS","AVES_GALLUS","CARNE_DESPOJOS","IMPO_MAIZ","RESIDUOS_GRASA","EXPO_HUEVOS","EXPO_GALLOS","EXPO_CARNE_DESPOJOS")



  return(valor_Fenavi)
}
