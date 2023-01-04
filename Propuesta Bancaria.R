#C?lculo de INGRESOS_12M
#rm(list=ls()) 

library(readxl)
setwd("C:/Users/DLSR_33/Desktop/00 NAF-DRV ESG")

df_carvg<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "CAR_VIGENTE",range = "A2:AZ11",
              col_names = T)
df_carvg<-df_carvg[,-1]
df_gastos<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "GASTOS_PYA",range = "A2:AZ11",
                   col_names = T)
df_gastos <- df_gastos[,-1]
df_margfin<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "MARGEN_FIN",range = "A2:AZ11",
                   col_names = T)
df_margfin <- df_margfin[,-1]
df_carvc<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "CAR_VENCIDA",range = "A2:AZ11",
                   col_names = T)
df_carvc<-df_carvc[,-1]
df_estpre<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "EST_PREVENTIVA",range = "A2:AZ11",
                   col_names = T)
df_estpre<-df_estpre[,-1]
df_befnt<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "BEF_NET",range = "A2:AZ11",
                   col_names = T)
df_befnt<-df_befnt[,-1]
df_actt<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "ACT_TOL",range = "A2:AZ11",
                   col_names = T)
df_actt<-df_actt[,-1]
df_actc<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "ACT_CIR",range = "A2:AZ11",
                   col_names = T)
df_actc<-df_actc[,-1]
df_pasc<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "PAS_CIR",range = "A2:AZ11",
                   col_names = T)
df_pasc<-df_pasc[,-1]
df_icap<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "ICAP",range = "B4:AZ13",
                   col_names = T)
#####Calculo de razones####
df_ccv <- data.frame(matrix(ncol = 51, nrow = 9))#Crear marco de datos.
for (f in 1:9){                    #for para las filas.
  for (c in 2:51){                  #for para las columnas.
    df_ccv[f,c] <- (df_carvg[f,c]-df_carvg[f,c-1])/df_carvg[f,c-1]   #Rend. trimestre c trimestre.
  }}

df_raef <- as.data.frame(matrix(data=0,nrow=9,ncol=51))
for (c in 1:51){
  for (f in 1:9){
    df_raef[f,c] <- df_gastos[f,c]/df_margfin[f,c]
  }
}

df_imor <- as.data.frame(matrix(data=0,nrow=9,ncol=51))
for (c in 1:51){
  for (f in 1:9){
    df_imor[f,c] <- df_carvc[f,c]/(df_carvg[f,c]+df_carvc[f,c])
  }
}

df_icob <- as.data.frame(matrix(data=0,nrow=9,ncol=51))
for (c in 1:51){
  for (f in 1:9){
    df_icob[f,c] <- -(df_estpre[f,c]/df_carvc[f,c])
  }
}

df_roa <- as.data.frame(matrix(data=0,nrow=9,ncol=51))
for (c in 1:51){
  for (f in 1:9){
    df_roa[f,c] <- df_befnt[f,c]/df_actt[f,c]
  }
}

df_rl <- as.data.frame(matrix(data=0,nrow=9,ncol=51))
for (c in 1:51){
  for (f in 1:9){
    df_rl[f,c] <- df_actc[f,c]/df_pasc[f,c]
  }
}
####Acomodo de base de datos ####              

#ccv
emisoras <- (df_ccv$Emisora)
fechas <- colnames(df_ccv)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_ccv) <- emisoras
colnames(df_ccv) <- fechas_Formato

#raef
emisoras <- (df_raef$Emisora)
fechas <- colnames(df_raef)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_raef) <- emisoras
colnames(df_raef) <- fechas_Formato


#imor
emisoras <- (df_imor$Emisora)
fechas <- colnames(df_imor)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_imor) <- emisoras
colnames(df_imor) <- fechas_Formato

#icob
emisoras <- (df_icob$Emisora)
fechas <- colnames(df_icob)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_icob) <- emisoras
colnames(df_icob) <- fechas_Formato

#roa
emisoras <- (df_roa$Emisora)
fechas <- colnames(df_roa)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_roa) <- emisoras
colnames(df_roa) <- fechas_Formato

#rl
emisoras <- (df_rl$Emisora)
fechas <- colnames(df_rl)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_rl) <- emisoras
colnames(df_rl) <- fechas_Formato

#icap
emisoras <- (df_icap$Emisora)
fechas <- colnames(df_icap)

fechas_Formato <- as.Date(as.numeric(fechas),origin="1899-12-30")
rownames(df_icap) <- emisoras
colnames(df_icap) <- fechas_Formato



##### Promedios ####
#Promedios de CCV
promedios_ccv <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
  for (r in 1:9){ #Promedios ccv
    for (c in 1:27) { 
      promedios_ccv[r,c] <- mean(as.numeric(df_ccv[r,(1:24)+c]), na.rm=TRUE)
  }
  }
  
#Promedios de raef
promedios_raef <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios raef
  for (c in 1:27) { 
    promedios_raef[r,c] <- mean(as.numeric(df_raef[r,(1:24)+c]), na.rm=TRUE)
  }
}

#Promedios de IMOR
promedios_imor <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios imor
  for (c in 1:27) { 
    promedios_imor[r,c] <- mean(as.numeric(df_imor[r,(1:24)+c]), na.rm=TRUE)
  }
}

#Promedios de ICOB
promedios_icob <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios icob
  for (c in 1:27) { 
    promedios_icob[r,c] <- mean(as.numeric(df_icob[r,(1:24)+c]), na.rm=TRUE)
  }
}

#Promedios de ROA
promedios_roa <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios roa
  for (c in 1:27) {
    promedios_roa[r,c] <- mean(as.numeric(df_roa[r,(1:24)+c]), na.rm=TRUE)
  }
}

#Promedios de rl
promedios_rl <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios rl
  for (c in 1:27) { 
    promedios_rl[r,c] <- mean(as.numeric(df_rl[r,(1:24)+c]), na.rm=TRUE)
  }
}

#rm=ls()
#Promedios de ICAP
promedios_icap <- as.data.frame(matrix(data=0,nrow=9,ncol=27))
for (r in 1:9){ #Promedios icap
  for (c in 1:27) { 
    promedios_icap[r,c] <- mean(as.numeric(df_icap[r,(1:24)+c]), na.rm=TRUE)
  }
}

##### Evaluaci?n ####
#evaluacion_ccv[,27] <- 0
#Evaluaci?n de las emisoras

#CCV
evaluacion_ccv <- matrix(data=0, ncol=9, nrow=27)
evaluacion_ccv <- ifelse(df_ccv[,25:51]>promedios_ccv[,1:27],1,0)


#RAEF
evaluacion_raef <- matrix(data=0, ncol=9, nrow=27)
evaluacion_raef <- ifelse(df_raef[,25:51]>promedios_raef[,1:27],1,0)

#IMOR
evaluacion_imor <- matrix(data=0, ncol=9, nrow=27)
evaluacion_imor <- ifelse(df_imor[,25:51]<promedios_imor[,1:27],1,0)


#ICOB
evaluacion_icob <- matrix(data=0, ncol=9, nrow=27)
evaluacion_icob <- ifelse(df_icob[,25:51]>promedios_icob[,1:27],1,0)

#ROA
evaluacion_roa <- matrix(data=0, ncol=9, nrow=27)
evaluacion_roa <- ifelse(df_roa[,25:51]>promedios_roa[,1:27],1,0)


#RL
evaluacion_rl <- matrix(data=0, ncol=9, nrow=27)
evaluacion_rl <- ifelse(df_rl[,25:51]>promedios_rl[,1:27],1,0)


#ICAP
evaluacion_icap <- matrix(data=0, ncol=9, nrow=27)
evaluacion_icap <- ifelse(df_icap[,25:51]>10.5,ifelse(df_icap[,25:51]>promedios_icap[,1:27],1,0))


####Sumas de evaluacion####
#ccv
suma_ccv <- rep(0,9)
suma_ccv <- ifelse(evaluacion_ccv[,1]==0,apply(evaluacion_ccv[,1:27],MARGIN=1,FUN=sum),
                                  apply(evaluacion_ccv[,1:27],MARGIN=1,FUN=sum))
#raef
suma_raef <- rep(0,9)
suma_raef <- ifelse(evaluacion_raef[,1]==0,apply(evaluacion_raef[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_raef[,1:27],MARGIN=1,FUN=sum))
#imor
suma_imor <- rep(0,9)
suma_imor <- ifelse(evaluacion_imor[,1]==0,apply(evaluacion_imor[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_imor[,1:27],MARGIN=1,FUN=sum))
#ICOB
suma_icob <- rep(0,9)
suma_icob <- ifelse(evaluacion_icob[,1]==0,apply(evaluacion_icob[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_icob[,1:27],MARGIN=1,FUN=sum))
#ROA
suma_roa <- rep(0,9)
suma_roa <- ifelse(evaluacion_roa[,1]==0,apply(evaluacion_roa[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_roa[,1:27],MARGIN=1,FUN=sum))
#RL
suma_rl <- rep(0,9)
suma_rl <- ifelse(evaluacion_rl[,1]==0,apply(evaluacion_rl[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_rl[,1:27],MARGIN=1,FUN=sum))
#ICAP
suma_icap <- rep(0,9)
suma_icap <- ifelse(evaluacion_icap[,1]==0,apply(evaluacion_icap[,1:27],MARGIN=1,FUN=sum),
                   apply(evaluacion_icap[,1:27],MARGIN=1,FUN=sum))


####Datos generales y calificaci?n final####

#CCV
promedio_ccv <- mean(suma_ccv,na.rm=TRUE)
maximo_ccv <- max(suma_ccv,na.rm=TRUE)
minimo_ccv <- min(suma_ccv,na.rm=TRUE)
calificacion_ccv <- ((suma_ccv-minimo_ccv)/(maximo_ccv-minimo_ccv)*(10-1)+1)

#raef
promedio_raef <- mean(suma_raef,na.rm=TRUE)
maximo_raef <- max(suma_raef,na.rm=TRUE)
minimo_raef <- min(suma_raef,na.rm=TRUE)
calificacion_raef <- ((suma_raef-minimo_raef)/(maximo_raef-minimo_raef)*(10-1)+1)

#IMOR
promedio_imor <- mean(suma_imor,na.rm=TRUE)
maximo_imor <- max(suma_imor,na.rm=TRUE)
minimo_imor <- min(suma_imor,na.rm=TRUE)
calificacion_imor <- ((suma_imor-minimo_imor)/(maximo_imor-minimo_imor)*(10-1)+1)

#ICOB
promedio_icob <- mean(suma_icob,na.rm=TRUE)
maximo_icob <- max(suma_icob,na.rm=TRUE)
minimo_icob <- min(suma_icob,na.rm=TRUE)
calificacion_icob <- ((suma_icob-minimo_icob)/(maximo_icob-minimo_icob)*(10-1)+1)

#roa
promedio_roa <- mean(suma_roa,na.rm=TRUE)
maximo_roa <- max(suma_roa,na.rm=TRUE)
minimo_roa<- min(suma_roa,na.rm=TRUE)
calificacion_roa <- ((suma_roa-minimo_roa)/(maximo_roa-minimo_roa)*(10-1)+1)

#RL
promedio_rl <- mean(suma_rl,na.rm=TRUE)
maximo_rl <- max(suma_rl,na.rm=TRUE)
minimo_rl <- min(suma_rl,na.rm=TRUE)
calificacion_rl <- ((suma_rl-minimo_rl)/(maximo_rl-minimo_rl)*(10-1)+1)

#ICAP
promedio_icap <- mean(suma_icap,na.rm=TRUE)
maximo_icap <- max(suma_icap,na.rm=TRUE)
minimo_icap <- min(suma_icap,na.rm=TRUE)
calificacion_icap <- ((suma_icap-minimo_icap)/(maximo_icap-minimo_icap)*(10-1)+1)


####Concentrado####

concentrado<-cbind(calificacion_ccv,calificacion_icap,calificacion_icob,calificacion_imor,calificacion_raef,calificacion_rl,calificacion_roa)
prom_concentrado <- as.data.frame(matrix(data=0,nrow=9,ncol=1))
prom_concentrado <- rowMeans(concentrado)
maximo_concentrado <- max(prom_concentrado,na.rm=TRUE)
minimo_concentrado <- min(prom_concentrado,na.rm=TRUE)
calificacion_final <- ((prom_concentrado-minimo_concentrado)/(maximo_concentrado-minimo_concentrado)*(10-1)+1)
m<-read_xlsx("Propuesta Bancaria.xlsx",sheet = "CONCENTRADO",range = "A3:A12",
                   col_names = T)
concentrado <-cbind(em,calificacion_ccv,calificacion_icap,calificacion_icob,calificacion_imor,calificacion_raef,
                    calificacion_rl,calificacion_roa,prom_concentrado,calificacion_final)
concentrado
