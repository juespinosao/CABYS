#' @export
# Frutas
# Cargar la biblioteca readxl

f_Frutas<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # Exportaciones ------------------------------------------------------------------
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Frutas <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "TOTAL EXPO_KTES")


  n_fila=which(Frutas == "010499" |Frutas == "010403",arr.ind = TRUE)[,"row"]
  n_fila=c(n_fila[[1]],n_fila[[2]])
  n_col_1=which(Frutas== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]
  n_col_2=which(Frutas== paste0((anio-1)," ",1),arr.ind = TRUE)[,"col"]



  Frutas2 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "CTES FOBPES")
  n_fila_2=which(Frutas2 == "010499" |Frutas2 == "010403",arr.ind = TRUE)[,"row"]
  n_fila_2=c(n_fila[[1]],n_fila[[2]])
  n_col_1_2=which(Frutas2== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]
  n_col_2_2=which(Frutas2== paste0((anio-1)," 01"),arr.ind = TRUE)[,"col"]

  Frutas3 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "IP_EXPO")


   valor_exportaciones=as.data.frame(cbind(t(Frutas[n_fila,n_col_2[1]:n_col_1[1]]),t(Frutas2[n_fila_2,n_col_2_2[1]:n_col_1_2[1]]),t(Frutas3[n_fila_2,n_col_2_2[1]:n_col_1_2[1]])))
   for (i in 1:6) {
     valor_exportaciones[,i]=as.numeric(valor_exportaciones[,i])
   }






  # Consumo interno ---------------------------------------------------------
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]
  Frutas <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/",archivo))

  columna_fila=which(grepl("Frutas citricas con exclusion plazas",Banano),arr.ind = TRUE)
  columna1=which(grepl("Frutas citricas retropolado",Frutas),arr.ind = TRUE)
  columna2=which(grepl("Otras frutas retropolado",Frutas),arr.ind = TRUE)
  fila1=min(which(Banano[,columna_fila[1]-3]==2013,arr.ind = TRUE)[,"row"])
  fila2=min(which(Banano[,columna_fila[1]-3]==anio,arr.ind = TRUE)[,"row"])


  Valor_Frutas=as.data.frame(na.omit(Frutas[fila1:(fila2+mes-1),c(columna1[1],columna2[1])]))




# IPP ---------------------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPP","NOMBRE"]

  IPP <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                       sheet = "3.1")

  n_col=which(IPP == "CODIGO",arr.ind = TRUE)[, "col"]

  dos_digitos=anio %% 100
  #identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie



  columna1=max(which(grepl(paste0(nombres_siglas[1],"-",(dos_digitos-2)),IPP),arr.ind = TRUE))
  columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),IPP),arr.ind = TRUE)

  fila1=which(grepl("01310",IPP[,n_col]),arr.ind = TRUE)


  #Tomar el valor que nos interesa
  Valor_IPP=as.data.frame(t(IPP[c(fila1[1]),c(columna1:columna2[1])]))
  tamaño=nrow(Valor_IPP)
  filas=c(seq(1, tamaño, by = 2),tamaño)
  Valor_IPP=as.numeric(Valor_IPP[filas,1])
  # Agrupar datos -----------------------------------------------------------


  return(list(variacion = valor_exportaciones, vector = Valor_Frutas,IPP=Valor_IPP))
}
