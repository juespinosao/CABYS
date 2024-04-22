#' @export
# Ganado_Precios
# Cargar la biblioteca readxl

f_Precios<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)
  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel



# IPP ---------------------------------------------------------------------

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPP","NOMBRE"]

  Precios <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                        sheet = "1.1")

  #Identificar la fila donde esta la palabra Periodo
  n_col=which(Precios == "CODIGO",arr.ind = TRUE)[, "col"]

 dos_digitos=anio %% 100
  #identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie
 if(mes==1){
   columna1=which(grepl(paste0(nombres_siglas[12],"-",(dos_digitos-1)),Precios),arr.ind = TRUE)
   columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),Precios),arr.ind = TRUE)
 }else{
   columna1=which(grepl(paste0(nombres_siglas[mes-1],"-",dos_digitos),Precios),arr.ind = TRUE)
   columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),Precios),arr.ind = TRUE)
 }
columna1=max(columna1)

  fila1=which(grepl("02111",Precios[,n_col]),arr.ind = TRUE)
  fila2=which(grepl("21111",Precios[,n_col]),arr.ind = TRUE)
  fila3=which(grepl("02100",Precios[,n_col]),arr.ind = TRUE)
  fila4=which(grepl("21113",Precios[,n_col]),arr.ind = TRUE)
  fila5=which(grepl("02211",Precios[,n_col]),arr.ind = TRUE)
  fila6=which(grepl("02151",Precios[,n_col]),arr.ind = TRUE)
  fila7=which(grepl("21121",Precios[,n_col]),arr.ind = TRUE)
  fila8=which(grepl("02310",Precios[,n_col]),arr.ind = TRUE)
  fila9=which(grepl("23319",Precios[,n_col]),arr.ind = TRUE)




  #Tomar el valor que nos interesa
  Valor_Precios=as.data.frame(t(Precios[c(fila1,fila2,fila3,fila4,fila5,fila6,fila7,fila8,fila9),c(columna1,columna2[1])]))

  Valor_Precios=as.data.frame(lapply(Valor_Precios, as.numeric))


# Porkcolombia ------------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="PorkColombia","NOMBRE"]


  Porkcol_pie <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Porkcolombia/",archivo),
                               sheet = "Precio nacional pie ")

  #Identificar la fila donde esta la palabra Periodo
  ubi=which(Porkcol_pie == "Mes",arr.ind = TRUE)



  #identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie
  columna1=which(grepl(anio,Porkcol_pie[ubi[1],]),arr.ind = TRUE)
  fila1=which(Porkcol_pie[,ubi[2]]=="Ene",arr.ind = TRUE)[,"row"]
   #Filtrar la fila del mes de interes

  Porkcol_canal <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Porkcolombia/",archivo),
                            sheet = "Precio nacional canal caliente")

  #Identificar la fila donde esta la palabra Periodo
  ubi=which(Porkcol_canal == "Mes",arr.ind = TRUE)



  #identificar las columna donde dice total general y peso en canal
  #identificar las columna donde dice total general y peso en canal
  columna2=which(grepl(anio,Porkcol_canal[ubi[1],]),arr.ind = TRUE)
  fila2=which(Porkcol_canal[,ubi[2]]=="Ene",arr.ind = TRUE)[,"row"]
  #Filtrar la fila del mes de interes



  #Tomar el valor que nos interesa
  Valor_Porkcol=cbind(Porkcol_pie[fila1:(fila1+mes-1),columna1],Porkcol_canal[fila2:(fila2+mes-1),columna2])

  Valor_Porkcol=as.data.frame(lapply(Valor_Porkcol, as.numeric))
  colnames(Valor_Porkcol)=c("pie","canal")



  # USP -----------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Leche_USP3","NOMBRE"]

  Precio_sinbon <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Leche/USP/",archivo),
                            sheet = "PRECIO SIN BONIF Res 0017-12")

  fila_fuente=which(grepl("Fuente",as.data.frame(t(Precio_sinbon))))
  col_per=which(Precio_sinbon=="Periodo",arr.ind = TRUE)[,"col"]
  col_nal=which(Precio_sinbon=="Nacional",arr.ind = TRUE)[,"col"]
  Precio_sinbon[,col_per[1]]=as.numeric(Precio_sinbon[,col_per[1]])
  Precio_sinbon[,col_per[1]] <- as.Date(Precio_sinbon[,col_per[1]], origin = "1899-12-30")
  Precio_sinbon[,col_per[1]]=format(Precio_sinbon[,col_per[1]], "%Y-%m")
  fila1=which(Precio_sinbon==paste0(anio,"-01"))
  Valor_sinbon=as.numeric(Precio_sinbon[fila1:(fila_fuente[length(fila_fuente)]-1),col_nal[1]])





  Precio_total <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Leche/USP/",archivo),
                           sheet = "PRECIO ($) TOTAL Res 0017-2012")

  fila_fuente=which(grepl("Fuente",as.data.frame(t(Precio_total))))
  col_per=which(Precio_total=="Periodo",arr.ind = TRUE)[,"col"]
  col_nal=which(Precio_total=="Nacional",arr.ind = TRUE)[,"col"]
  Precio_total[,col_per[1]]=as.numeric(Precio_total[,col_per[1]])
  Precio_total[,col_per[1]] <- as.Date(Precio_total[,col_per[1]], origin = "1899-12-30")
  Precio_total[,col_per[1]]=format(Precio_total[,col_per[1]], "%Y-%m")
  fila1=which(Precio_total==paste0(anio,"-01"))
  Valor_total=as.numeric(Precio_total[fila1:(fila_fuente[length(fila_fuente)]-1),col_nal[1]])

valor_usp=as.data.frame(cbind(Valor_sinbon,Valor_total))

valor_usp=as.data.frame(lapply(valor_usp, as.numeric))



# IPC -----------------------------------------------------------


archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPC","NOMBRE"]


IPC <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                     sheet = "8")

#Identificar la fila donde esta la palabra Periodo
n_col1=which(IPC == "Código",arr.ind = TRUE)[, "col"]

n_col2=which(IPC == "Índice",arr.ind=TRUE)[, "col"]


fila1=which(grepl("01140500",IPC[,n_col1]),arr.ind = TRUE)
fila2=which(grepl("01120300",IPC[,n_col1]),arr.ind = TRUE)



#Tomar el valor que nos interesa
Valor_IPC=as.data.frame(t(IPC[c(fila1,fila2),n_col2]))

Valor_IPC=as.data.frame(lapply(Valor_IPC, as.numeric))


  return(list(Valor_Precios,Valor_Porkcol,valor_usp,Valor_IPC))

}
