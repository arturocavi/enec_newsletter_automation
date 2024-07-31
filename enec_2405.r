#Ficha ENEC
#Incluyendo Valor de la producción por sector y tipo de obra
#23/02/2024 (Enconding UTF-8)


#Obtención de Datos para Gráficas y Texto de la ENEC para Ficha y Boletín
#Fuente: INEGI, ENEC Datos Abiertos Por Entidad Federativa

#Limpiar variables
rm(list=ls())


########## LIBRERIA ##########
library(dplyr)
library(tidyverse)
library(openxlsx)


########## DIRECTORIO ##########
#MODIFICACIÓN DE DATOS AUTOMÁTICA (la que se usa siempre, a menos que el mes en que se corra no coincida con el mes de publicación de INEGI)
#Determinar el mes y el año de la encuesta, bajo el supuesto que se publica dos meses después
#Hacer adecuaciones en caso de que se ejecute el código en una fecha distinta a la publicación de INEGI (ver "MODIFICACIÓN DE DATOS MANUAL" 4 líneas abajo)
mes = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],6,7))
año = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],1,4))

#MODIFICACIÓN DE DATOS MANUAL (la que casi no se usa, solo se corre cuando no coincida con el mes de publicación de INEGI)
#Mantener esta sección comentada. Solo usar en caso de desface de mes con INEGI. Por ejemplo, pruebas para modificar el código; o cuando no se corrió el mes que tocaba.
# año=año #a mano
# mes=mes-1 #a mano



#Ponerle cero al mes2
if (mes < 10) {
  mes2 = paste0("0",mes)
  # Si el mes es menor a 10, le pegamos un 0 delante.
} else {
  mes2 = as.character(mes)
}

###
#Directorio
directorio = paste0("C:/Users/arturo.carrillo/Documents/ENEC/",año," ",mes2)
setwd(directorio)


########## PERIODOS DE TÍTULOS DE GRÁFICAS ##########

#Nombres de meses
meses = data.frame(c("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"))
colnames(meses) = "MES" #Cambia nombre de la columna

#Periodos de títulos
periodo1=paste0(meses[mes,1]," 2018-",paste(meses[mes,1],año))
periodo2=paste0("enero 2018-",paste(meses[mes,1],año))
periodo3=paste(meses[mes,1],año)


########## NÚMEROS CÁRDINALES ##########
#Números cardinales DEL 1 AL 20 en masculino y femenino
num_cardinales=data.frame(matrix(c("primer","segundo","tercer","cuarto","quinto","sexto","séptimo","octavo","noveno","décimo","decimoprimer","decimosegundo","decimotercer","decimocuarto","decimoquinto","decimosexto","decimoséptimo","decimoctavo","decimonoveno","vigésimo","primera","segunda","tercera","cuarta","quinta","sexta","séptima","octava","novena","décima","decimoprimera","decimosegunda","decimotercera","decimocuarta","decimoquinta","decimosexta","decimoséptima","decimoctava","decimonovena","vigésima"),nrow=20))
colnames(num_cardinales)=c("MAS","FEM")


########## NOMBRES DE ENTIDADES FEDERATIVAS ##########
#Nombres comunes de Entidades Federativas, los que en verdad se usan
nombre_ef=data.frame(c("Nacional", "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Coahuila", "Colima", "Chiapas", "Chihuahua", "Ciudad de México", "Durango", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "Estado de México", "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"))
colnames(nombre_ef) = "Nombre_EF"


########## DESCARGAR BASE DE INEGI ##########
temp = tempfile() #crear archivo temporal
#Descargar archivo del INEGI:
download.file("https://www.inegi.org.mx/contenidos/programas/enec/2018/datosabiertos/enec_mensual_csv.zip",temp)

#Nombre de la base a buscar en el archivo de INEGI
nombre_base_csv=paste0("enec_absoluto_entidad_2018_",año,".csv")

#Descomprimir archivo y leer base
base_csv = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")
unlink(temp) #borrar archivo  temporal (el descargado de INEGI)

# #MODIFICACIÓN DE DATOS MANUAL (la que casi no se usa, solo se corre cuando no coincida con el mes de publicación de INEGI)
# #Mantener esta sección comentada. Solo usar en caso de desface de mes con INEGI. Por ejemplo, pruebas para modificar el código; o cuando no se corrió el mes que tocaba.
# nombre_base_csv=paste0("enec_absoluto_entidad_2018_",año,".csv")
# base_csv=read.csv(nombre_base_csv)


########## DATOS DEL MISMO MES PARA TODOS LOS AÑOS DE JALISCO (EN MILLONES DE PESOS CORRIENTES) ##########
#Datos mensuales de Jalisco de valor de la produccion en la entidad a precios corrientes
vp_men_jal=filter(base_csv,CODIGO_ENTIDAD==14) %>%
  select(ANIO,MES,VP=VP_TOTAL_ENTIDAD) #VP_TOTAL_ENTIDAD: Valor de producción total según la localización de las obras (Miles de pesos corrientes)

vp_mismes=vp_men_jal[vp_men_jal$MES==mes,]
if (mes<10){
  vp_mismes$MES=paste0(rep(0,nrow(vp_mismes)),vp_mismes$MES)
}
vp_mismes=mutate(vp_mismes,FECHA=paste(vp_mismes$ANIO,vp_mismes$MES, sep="/"))

graf_mes=select(vp_mismes,FECHA,VP)
graf_mes$VP=graf_mes$VP/1000 #Para que sean millones


########## TEXTO DESCRIPCIÓN 1: graf_mes ##########
descripcion1 = paste0("De acuerdo con datos de la Encuesta Nacional de Empresas Constructoras (ENEC), la producción de la industria de la construcción en Jalisco se ubicó en ",
                      format(round(graf_mes[nrow(graf_mes),2]),big.mark=","),
                      " millones de pesos corrientes durante el mes de ", meses[mes,1]," de ", año,
                      ", ubicándose por ")
if (graf_mes[nrow(graf_mes),2] > graf_mes[nrow(graf_mes)-1,2]) {
  # Si la producción de este mes es mayor a la del mismo mes del año pasado...
  descripcion1 = paste0(descripcion1,"arriba")
  ind = "aumenta" # Etiqueta para el título del boletín
} else {
  # Si la producción de este mes es menor a la del mismo mes del año pasado...
  descripcion1 = paste0(descripcion1,"debajo")
  ind = "disminuye" # Etiqueta para el título del boletín
}
descripcion1 = paste0(descripcion1," de la cifra de ",meses[mes,1]," de ",año-1," que alcanzó los ",
                      format(round(graf_mes[nrow(graf_mes)-1,2]),big.mark=","),
                      " millones de pesos corrientes. Esto representó ")

# El porcentaje de incremento/decremento no está presente en esta tabla, así que hay que calcularlo
porcentaje = abs((graf_mes[nrow(graf_mes),2] - graf_mes[nrow(graf_mes)-1,2]) / graf_mes[nrow(graf_mes)-1,2])
porcentaje = paste0(format(round(porcentaje*100,1),nsmall=1),"%")

if (graf_mes[nrow(graf_mes),2] > graf_mes[nrow(graf_mes)-1,2]) {
  # Si la producción de este mes es mayor a la del mismo mes del año pasado...
  descripcion1 = paste0(descripcion1,"un aumento de ",porcentaje)
} else {
  # Si la producción de este mes es menor a la del mismo mes del año pasado...
  descripcion1 = paste0(descripcion1,"una disminución de ",porcentaje)
}
descripcion1 = paste0(descripcion1," respecto al mismo mes del año anterior.")






########## DATOS HISTORICOS MENSUALES DE PRODUCCION DE JALISCO EN MILLONES DE PESOS A PRECIOS CORRIENTES ##########
#Cambiar Meses de numero a letras
vp_men_jal$MES=gsub(10,"OCT",vp_men_jal$MES)
vp_men_jal$MES=gsub(11,"NOV",vp_men_jal$MES)
vp_men_jal$MES=gsub(12,"DIC",vp_men_jal$MES)
vp_men_jal$MES=gsub(1,"ENE",vp_men_jal$MES)
vp_men_jal$MES=gsub(2,"FEB",vp_men_jal$MES)
vp_men_jal$MES=gsub(3,"MAR",vp_men_jal$MES)
vp_men_jal$MES=gsub(4,"ABR",vp_men_jal$MES)
vp_men_jal$MES=gsub(5,"MAY",vp_men_jal$MES)
vp_men_jal$MES=gsub(6,"JUN",vp_men_jal$MES)
vp_men_jal$MES=gsub(7,"JUL",vp_men_jal$MES)
vp_men_jal$MES=gsub(8,"AGO",vp_men_jal$MES)
vp_men_jal$MES=gsub(9,"SEP",vp_men_jal$MES)

#Datos historicos mensuales desde 2018
dat_his_men=vp_men_jal$VP
dhm=length(dat_his_men)-12
#Promedio de historicos mensuales
#Crear base vacia
pro_his=matrix(data=0,dhm,1)
pro_his=as.data.frame(pro_his)
names(pro_his)="PROM"

for (i in 1:(dhm)){
  pro_his[i,1]=mean(dat_his_men[(i+1):(i+12)])
}

#Fechas para graficar
#Crear base vacía
fecha_gra=matrix(data="",dhm,2)
fecha_gra=as.data.frame(fecha_gra)
fecha_gra=data.frame(lapply(fecha_gra, as.character), stringsAsFactors=FALSE)
names(fecha_gra)=c("ANIO","MES")

vp_his=vp_men_jal[vp_men_jal$ANIO>2018,]

for (i in 1:dhm){
  if(vp_his[i,2]=="ENE"){
    fecha_gra[i,1]=vp_his[i,1]
  }
  fecha_gra[i,2]=vp_his[i,2]
}

graf_his=cbind.data.frame(fecha_gra,VP=vp_his$VP/1000,pro_his/1000)


########## TEXTO DESCRIPCIÓN 2: graf_his ##########
descripcion2 = "Con base en cifras originales a precios corrientes, la producción de la industria de la construcción en Jalisco "

# No contamos con el porcentaje en esta tabla, por lo que tenemos que calcularlo.
porcentaje = abs((graf_his[nrow(graf_his),3] - graf_his[nrow(graf_his)-1,3]) / graf_his[nrow(graf_his)-1,3])
porcentaje = paste0(format(round(porcentaje*100,1),nsmall=1),"%")

# Para revisar cuántos registros consecutivos son mayores a su anterior
ainc = 0 #antes incrementos
while (graf_his[nrow(graf_his)-ainc,4] > graf_his[nrow(graf_his)-ainc-1,4]) {
  ainc = ainc + 1
}

# Para revisar cuántos registros consecutivos son menores a su anterior
adec = 0 #antes decrementos
while (as.numeric(graf_his[nrow(graf_his)-adec,4]) < as.numeric(graf_his[nrow(graf_his)-adec-1,4])) {
  adec = adec + 1
}

# Para revisar cuántos registros consecutivos son mayores a su anterior hubo antes de la primera caída
ainc1pa = 1 # (ainc 1 periodo antes)
while (graf_his[nrow(graf_his)-ainc1pa,4] > graf_his[nrow(graf_his)-ainc1pa-1,4]) {
  ainc1pa = ainc1pa + 1
}

# Para revisar cuántos registros consecutivos son menores a su anterior hubo antes de la primera subida
adec1pa = 1 # (adec 1 periodo antes)
while (as.numeric(graf_his[nrow(graf_his)-adec1pa,4]) < as.numeric(graf_his[nrow(graf_his)-adec1pa-1,4])) {
  adec1pa = adec1pa + 1
}


if (graf_his[nrow(graf_his),3] > graf_his[nrow(graf_his)-1,3]) {
  # Si hubo incremento con respecto al mes inmediato anterior
  descripcion2 = paste0(descripcion2,"aumentó ",porcentaje," en ",meses[mes,1]," de ",año,
                        " respecto al mes inmediato anterior. ")
  
  if (graf_his[nrow(graf_his),4] > graf_his[nrow(graf_his)-1,4]) { 
    # Si hubo incremento en el promedio anual con respecto al mes inmediato anterior
    if (ainc == 1) {
      if (adec1pa > 5) {
        # Si es el primer mes que cambia después de 6 o más meses en el mismo sentido
        descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses subió por primera vez después de  ",adec1pa,
                              " meses a la baja, pasando de ")
      }
      else {
        descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses presentó un incremento, pasando de ")
      }
    }
    else if (ainc <= 20) {
      descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses subió por ",num_cardinales[ainc,1],
                            " mes consecutivo, pasando de ")
    }
    else if (ainc > 20) {
      descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses mantiene su tendencia creciente, pasando de ")
    }
  }
  else if (graf_his[nrow(graf_his),4] < graf_his[nrow(graf_his)-1,4]) {
    # Si hubo disminuciones en el promedio anual con respecto al mes inmediato anterior
    if (adec == 1) {
      if (ainc1pa > 5) {
        # Si es el primer mes que cambia después de 6 o más meses en el mismo sentido
        descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses bajó por primera vez después de ", ainc1pa,
                              " meses consecutivos al alza, pasando de ")
      }
      else {
        descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses presentó una caída, pasando de ")
      }
    }
    else if (adec <= 20) {
      descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses bajó por ",num_cardinales[adec,1],
                            " mes consecutivo, pasando de ")
    }
    else if (adec > 20) {
      descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses mantiene su tendencia decreciente, pasando de ")
    }
  }
} else {
  # Si hubo una disminución con respecto al mes inmediato anterior
  descripcion2 = paste0(descripcion2,"disminuyó ",porcentaje," en ",meses[mes,1]," de ",año,
                        " respecto al mes inmediato anterior. ")
  
  if (graf_his[nrow(graf_his),4] < graf_his[nrow(graf_his)-1,4]) {
    # Si hubo decremento en el promedio anual con respecto al mes inmediato anterior
    if (adec == 1) {
      if (ainc1pa > 5) {
        # Si es el primer mes que cambia después de 6 o más meses al alza
        descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses bajó por primera vez después de ", ainc1pa,
                              " meses consecutivos al alza, pasando de ")
      }
      else {
        descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses presentó una caída, pasando de ")
      }
    }
    else if (adec <= 20) {
      descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses bajó por ",num_cardinales[adec,1],
                            " mes consecutivo, pasando de ")
    }
    else if (adec > 20) {
      descripcion2 = paste0(descripcion2,"Asimismo, el comportamiento promedio de los últimos doce meses mantiene su tendencia decreciente, pasando de ")
    }
  } else {
    # Si hubo incremento en el promedio anual con respecto al mes inmediato anterior...
    if (ainc == 1) {
      if (adec1pa > 5) {
        # Si es el primer mes que cambia después de 6 o más meses a la baja
        descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses subió por primera vez después de ", adec1pa,
                              " meses consecutivos a la baja, pasando de ")
      }
      else {
        descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses presentó un incremento, pasando de ")
      }
    }
    else if (ainc <= 20){
      descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses subió por ",num_cardinales[ainc,1],
                            " mes consecutivo, pasando de ")
    }
    else if (ainc > 20){
      descripcion2 = paste0(descripcion2,"Sin embargo, el comportamiento promedio de los últimos doce meses mantiene su tendencia creciente, pasando de ")
    }
    
  }
  
}

# Si el mes de la información es enero, nos aseguramos que el mes pasado sea diciembre del año pasado
antmes = mes - 1
antaño = año
if (antmes == 0){
  antmes = 12
  antaño = antaño-1
  }

if (graf_his[nrow(graf_his),4] == graf_his[nrow(graf_his)-1,4]) {
  descripcion2 = paste0(descripcion2,"No obstante, el comportamiento promedio de los últimos doce meses se mantuvo sin cambios.")
} else {
  descripcion2 = paste0(descripcion2,format(round(graf_his[nrow(graf_his)-1,4]),big.mark=","),
                        " millones de pesos corrientes en ",meses[antmes,1]," de ",antaño," a ",
                        format(round(graf_his[nrow(graf_his),4]),big.mark=","),
                        " millones de pesos corrientes en ",meses[mes,1],
                        " de ",año,".")
}


########## PARTICIPATION PORCENTUAL POR ENTIDAD FEDERATIVA DEL VALOR DE LA PRODUCCION NACIONAL ##########
#Datos
base_estados=filter(base_csv,ANIO==año,MES==mes) %>%
  select(ENTIDAD,VP=VP_TOTAL_ENTIDAD)
base_estados=base_estados[1:33,]

#Unir Nombre Común de Entidades Federativas
base_estados=cbind.data.frame(nombre_ef,VP=base_estados[,2])

#Participacion porcentual del ultimo dato de valor de la produccion
par_base_estados=base_estados[2:33,2]/base_estados[1,2]

#Base de grafica para ranking de participacion
graf_rank_dis=cbind.data.frame(Entidad=nombre_ef[2:33,1],Porcentaje=par_base_estados)
graf_rank_dis[,'Porcentaje']=as.numeric(as.character(graf_rank_dis[,'Porcentaje']))
graf_rank_dis=graf_rank_dis[order(graf_rank_dis$Porcentaje),]
graf_rank_dis[,'Porcentaje']=graf_rank_dis[,'Porcentaje']*100


########## TEXTO DESCRIPCIÓN 3: graf_rank_dis ##########
#Variables
mesp = meses[mes,1] #mes en palabras
comp = which(graf_rank_dis$Entidad=="Jalisco") #extraemos el número de fila de Jalisco
rankjal = round(graf_rank_dis[comp,2],1) #ranking para Jalisco



descripcion3 = paste0("Durante ",mesp," de ",año,", Jalisco concentró ",rankjal,"%",
                      " del total del valor de la producción de la industria de la construcción en el país, ubicando a Jalisco en el ")

# Revisamos que el lugar de Jalisco esté dentro del Top 10 para usar números cardinales.
# De lo contrario, se usarán números ordinales para expresarlo.Importante, porque el orden cambia.
# El lugar se calcula como 32-comp+1 porque el orden es de menor a mayor.
if (32-comp+1 <= 10) {
  descripcion3 = paste0(descripcion3,num_cardinales[32-comp+1,1]," lugar a nivel nacional")
} else {
  descripcion3 = paste0(descripcion3,"lugar ",32-comp+1," a nivel nacional")
}

# Revisamos ahora si Jalisco está en el top 5. De lo contrario, se acaba el párrafo aquí.
# Además, si Jalisco también es primer lugar, el párrafo también se acaba aquí.
if (((32-comp+1) > 5) | ((32-comp+1) == 1)) {
  descripcion3 = paste0(descripcion3,".")
} else {
  
  # Ciclo para integrar al texto los estados que tuvieron un mejor desempeño que Jalisco
  descripcion3 = paste0(descripcion3,", por debajo de ")
  
  for (i in 1:(32-comp)) {
    
    if (i == (32-comp)) {
      descripcion3 = paste0(descripcion3,graf_rank_dis[32-i+1,1],".")
    } else if (i == (32-comp-1)) {
      descripcion3 = paste0(descripcion3,graf_rank_dis[32-i+1,1]," y ")
    } else {
      descripcion3 = paste0(descripcion3,graf_rank_dis[32-i+1,1],", ")
    }
    
  }
  
}


########## VARIACIÓN ANUAL DE LA PRODUCCIÓN MENSUAL POR ENTIDAD FEDERATIVA ##########
#Datos
#Ultimo dato
base_ud=filter(base_csv,ANIO==año,MES==mes) %>%
  select(DESCRIPCION_ACTIVIDAD,ANIO,MES,ENTIDAD,VP=VP_TOTAL_ENTIDAD)
base_ud=base_ud[1:33,]
vp_ult_dat=base_ud[,5]
#Año anterior
base_aa=filter(base_csv,ANIO==año-1,MES==mes) %>%
  select(DESCRIPCION_ACTIVIDAD,ANIO,MES,ENTIDAD,VP=VP_TOTAL_ENTIDAD)
base_aa=base_aa[1:33,]
vp_año_ant=base_aa[,5]

#Leer nombres de entidades
vp_ult_ant=cbind.data.frame(nombre_ef,ULT=vp_ult_dat,ANT=vp_año_ant,VAR=vp_ult_dat/vp_año_ant-1)

#Variación porcentual anual del dato mensual
varporanu_datmen=vp_ult_ant$VAR

#Base de grafica para ranking de variación porcentual de la producción
graf_rank_var=cbind.data.frame(Entidad=nombre_ef[,1],Variación=varporanu_datmen)
graf_rank_var[,'Variación']=as.numeric(as.character(graf_rank_var[,'Variación']))
graf_rank_var=graf_rank_var[order(graf_rank_var$Variación),]
graf_rank_var[,'Variación']=graf_rank_var[,'Variación']*100


########## TEXTO DESCRIPCIÓN 4: graf_rank_var##########
#Ranking sin nacional
rank_var_sn=cbind.data.frame(Entidad=nombre_ef[,1],Variación=varporanu_datmen)
rank_var_sn[,'Variación']=as.numeric(as.character(rank_var_sn[,'Variación']))
rank_var_sn=filter(rank_var_sn,Entidad!="Nacional")
rank_var_sn=arrange(rank_var_sn,-Variación)
rank_var_sn[,'Variación']=rank_var_sn[,'Variación']*100

#Extraemos el número de fila de Jalisco
lug = which(rank_var_sn$Entidad=="Jalisco")

#Valores de cambio Estatal y Nacional
camjal=round(rank_var_sn[lug,2],1) #cambio en Jalisco
camnac=round(graf_rank_var[graf_rank_var$Entidad=="Nacional",2],1) #cambio nacional

#Dirección del cambio
if(camjal>0){
  dircam="aumentó"
} else if(camjal<0){
  dircam="disminuyó"
} else{
  dircam="no presentó crecimiento"
  }

#Comparación respecto a la variación nacional
if(camjal>camnac){
  compnac="por arriba de "
} else if(camjal<camnac){
  compnac="por debajo de "
} else{
  compnac="en línea con "
}

descripcion4 = paste0("Durante ",meses[mes,1]," de ",año,
                      ", la producción mensual de la industria de la construcción estatal, que ",dircam," ",
                      round(abs(camjal),2),"% a tasa anual, ubicó al Estado ",compnac,"la variación nacional, colocando a Jalisco en el ")

# Revisamos que el lugar de Jalisco esté dentro del Top 10 para usar números cardinales.
# De lo contrario, se usarán números ordinales para expresarlo.Importante, porque el orden de las palabras cambia.
if (lug<=10) {
  descripcion4 = paste0(descripcion4,num_cardinales[lug,1]," lugar",
                        " respecto al resto de las entidades del país en cuanto a crecimiento del valor de la producción.")
} else {
  descripcion4 = paste0(descripcion4,"lugar ",lug,
                        " respecto al resto de las entidades del país en cuanto a crecimiento del valor de la producción.")
}


########## VARIACIÓN MENSUAL DE LA PRODUCCIÓN MENSUAL POR ENTIDAD FEDERATIVA ##########

#Datos
#Mes anterior
if (mes==1){
  #caso especial inicio de inicio de año
  base_ma=filter(base_csv,ANIO==año-1,MES==12) %>%
    select(DESCRIPCION_ACTIVIDAD,ANIO,MES,ENTIDAD,VP=VP_TOTAL_ENTIDAD)
  base_ma=base_ma[1:33,]
  vp_mes_ant=base_ma[,5]
} else{
  base_ma=filter(base_csv,ANIO==año,MES==mes-1) %>%
    select(DESCRIPCION_ACTIVIDAD,ANIO,MES,ENTIDAD,VP=VP_TOTAL_ENTIDAD)
  base_ma=base_ma[1:33,]
  vp_mes_ant=base_ma[,5]
}


#Leer nombres de entidades
vp_ult_mant=cbind.data.frame(nombre_ef,ULT=vp_ult_dat,ANT=vp_mes_ant,VAR=vp_ult_dat/vp_mes_ant-1)

#Variación porcentual mensual del dato mensual
varpormen_datmen=vp_ult_mant$VAR

#Base de grafica para ranking de variación porcentual de la producción
graf_rank_varmen=cbind.data.frame(Entidad=nombre_ef[,1],Variación=varpormen_datmen)
graf_rank_varmen[,'Variación']=as.numeric(as.character(graf_rank_varmen[,'Variación']))
graf_rank_varmen=graf_rank_varmen[order(graf_rank_varmen$Variación),]
graf_rank_varmen[,'Variación']=graf_rank_varmen[,'Variación']*100


########## TEXTO DESCRIPCIÓN 5: graf_rank_varmen ##########
#Ranking sin nacional
rank_var_sn=cbind.data.frame(Entidad=nombre_ef[,1],Variación=varpormen_datmen)
rank_var_sn[,'Variación']=as.numeric(as.character(rank_var_sn[,'Variación']))
rank_var_sn=filter(rank_var_sn,Entidad!="Nacional")
rank_var_sn=arrange(rank_var_sn,-Variación)
rank_var_sn[,'Variación']=rank_var_sn[,'Variación']*100

#Extraemos el número de fila de Jalisco
lug = which(rank_var_sn$Entidad=="Jalisco")

#Valores de cambio Estatal y Nacional
camjal=round(rank_var_sn[lug,2],1) #cambio en Jalisco
camnac=round(graf_rank_varmen[graf_rank_varmen$Entidad=="Nacional",2]) #cambio nacional

#Dirección del cambio
if(camjal>0){
  dircam="aumentó"
} else if(camjal<0){
  dircam="disminuyó"
} else{
  dircam="no presentó crecimiento"
}

#Comparación respecto a la variación nacional
if(camjal>camnac){
  compnac="por arriba de "
} else if(camjal<camnac){
  compnac="por debajo de "
} else{
  compnac="en línea con "
}

descripcion5 = paste0("Durante ",meses[mes,1]," de ",año,
                      ", la producción mensual de la industria de la construcción estatal, que ",dircam," ",
                      format(abs(camjal), nsmall = 1),"% a tasa mensual, ubicó al Estado ",compnac,
                      "la variación nacional, colocando a Jalisco en el ")

# Revisamos que el lugar de Jalisco esté dentro del Top 10 para usar números cardinales.
# De lo contrario, se usarán números ordinales para expresarlo.Importante, porque el orden de las palabras cambia.
if (lug<=10) {
  descripcion5 = paste0(descripcion5,num_cardinales[lug,1]," lugar",
                        " respecto al resto de las entidades del país en cuanto a crecimiento mensual del valor de la producción.")
} else {
  descripcion5 = paste0(descripcion5,"lugar ",lug,
                        " respecto al resto de las entidades del país en cuanto a crecimiento mensual del valor de la producción.")
}



########## DATOS POR SECTOR CONTRATANTE (EN MILLONES DE PESOS CORRIENTES) ##########
#Datos mensuales de Jalisco de valor de la produccion por sector contratente
vp_sec=filter(base_csv,CODIGO_ENTIDAD==14) %>%
  select(ANIO,MES,SECTOR_PUBLICO,SECTOR_PRIVADO) %>%
  mutate(POR_SEC_PUB = SECTOR_PUBLICO/(SECTOR_PUBLICO+SECTOR_PRIVADO))

#Convertir en millones de pesos
vp_sec$SECTOR_PUBLICO = vp_sec$SECTOR_PUBLICO/1000
vp_sec$SECTOR_PRIVADO = vp_sec$SECTOR_PRIVADO/1000

#Cambiar Meses de numero a letras
vp_sec$MES=gsub(10,"OCT",vp_sec$MES)
vp_sec$MES=gsub(11,"NOV",vp_sec$MES)
vp_sec$MES=gsub(12,"DIC",vp_sec$MES)
vp_sec$MES=gsub(1,"ENE",vp_sec$MES)
vp_sec$MES=gsub(2,"FEB",vp_sec$MES)
vp_sec$MES=gsub(3,"MAR",vp_sec$MES)
vp_sec$MES=gsub(4,"ABR",vp_sec$MES)
vp_sec$MES=gsub(5,"MAY",vp_sec$MES)
vp_sec$MES=gsub(6,"JUN",vp_sec$MES)
vp_sec$MES=gsub(7,"JUL",vp_sec$MES)
vp_sec$MES=gsub(8,"AGO",vp_sec$MES)
vp_sec$MES=gsub(9,"SEP",vp_sec$MES)

#Fechas para graficar
dhm=nrow(vp_sec) #Datos historicos mensuales
#Crear base vacía
fecha_gra=matrix(data="",dhm,2)
fecha_gra=as.data.frame(fecha_gra)
fecha_gra=data.frame(lapply(fecha_gra, as.character), stringsAsFactors=FALSE)
names(fecha_gra)=c("ANIO","MES")

for (i in 1:dhm){
  if(vp_sec[i,2]=="ENE"){
    fecha_gra[i,1]=vp_sec[i,1]
  }
  fecha_gra[i,2]=vp_sec[i,2]
}

graf_sec=cbind.data.frame(fecha_gra,vp_sec[,3:5])



########## TEXTO DESCRIPCIÓN 6: graf_sec ##########

#Variables a necesitar

#Construccion sector publico
c_sp <- base_csv %>% filter(ANIO == año, MES == mes, ENTIDAD == 'Jalisco') %>%
  select(SECTOR_PUBLICO) %>% .[1,1]
c_sp <- c_sp/1000

##Construccion sector publico mismo mes año anterior
c_sp_Ant_Anual <- base_csv %>% filter(ANIO == año-1, MES == mes, ENTIDAD == 'Jalisco') %>%
  select(SECTOR_PUBLICO) %>% .[1,1]
c_sp_Ant_Anual <- c_sp_Ant_Anual/1000

c_sp_var_Anual <- ((c_sp/c_sp_Ant_Anual)-1)*100


#Construccion sector privado
c_spri <- base_csv %>% filter(ANIO == año, MES == mes, ENTIDAD == 'Jalisco') %>%
  select(SECTOR_PRIVADO) %>% .[1,1]
c_spri <- c_spri/1000

##Construccion sector privado mismo mes año anterior
c_spri_Ant_Anual <- base_csv %>% filter(ANIO == año-1, MES == mes, ENTIDAD == 'Jalisco') %>%
  select(SECTOR_PRIVADO) %>% .[1,1]
c_spri_Ant_Anual <- c_spri_Ant_Anual/1000

c_spri_var_Anual <- ((c_spri/c_spri_Ant_Anual)-1)*100


#Porcentaje el sector publico
por_c_sp <- (c_sp/(c_spri+c_sp))*100
por_c_sp_ant_anual <- (c_sp_Ant_Anual/(c_spri_Ant_Anual+c_sp_Ant_Anual))*100

#texto
descripcion6 <- paste0('Durante ', mesp, ' de ' , año, ', la producción ',
                       'mensual de la industria de la construcción estatal ',
                       'contratada por el sector público de los tres niveles de gobierno ')

if(c_sp > c_sp_Ant_Anual){
  descripcion6 <- paste0(descripcion6, 'aumentó ', round(c_sp_var_Anual,1),
                         '% respecto al año anterior, pasando de ',
                         format(round(c_sp_Ant_Anual,0), big.mark = ','),
                         ' millones de pesos corrientes en ', mesp, ' de ',
                         año-1, ' a ', format(round(c_sp,0), big.mark = ','),
                         ' millones de pesos corrientes en ', mesp, ' de ',
                         año,'.')
  if(c_spri > c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Además, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado aumentó ',
                           round(c_spri_var_Anual,1), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else if(c_spri < c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado disminuyó ',
                           abs(round(c_spri_var_Anual,1)), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else{
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado se mantuvo ',
                           'en línea con la cifra del mismo mes del año anterior',
                           round(c_spri_var_Anual,1), ' millones de pesos corrientes.')
  }
  
  
}else if(c_sp < c_sp_Ant_Anual){
  descripcion6 <- paste0(descripcion6, 'disminuyó ', abs(round(c_sp_var_Anual,1)),
                         '% respecto al año anterior, pasando de ',
                         format(round(c_sp_Ant_Anual,0), big.mark = ','),
                         ' millones de pesos corrientes en ', mesp, ' de ',
                         año-1, ' a ', format(round(c_sp,0), big.mark = ','),
                         ' millones de pesos corrientes en ', mesp, ' de ',
                         año,'.')
  
  if(c_spri > c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado aumentó ',
                           round(c_spri_var_Anual,1), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else if(c_spri < c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Además, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado disminuyó ',
                           abs(round(c_spri_var_Anual,1)), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else{
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado se mantuvo ',
                           'en línea con la cifra del mismo mes del año anterior',
                           round(c_spri_var_Anual,1), ' millones de pesos corrientes.')
  }
  
  
}else{
  descripcion6 <- paste0(descripcion6, 'se mantuvo sin cambios respecto a la',
                         ' cifra del mismo mes del año anterior, ',
                         format(round(c_sp,0), big.mark = ','),
                         ' millones de pesos corrientes.')
  if(c_spri > c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado aumentó ',
                           round(c_spri_var_Anual,1), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else if(c_spri < c_spri_Ant_Anual){
    descripcion6 <- paste0(descripcion6, ' Sin embargo, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado disminuyó ',
                           abs(round(c_spri_var_Anual,1)), '% respectó al año ',
                           'anterior, pasando de ',
                           format(round(c_spri_Ant_Anual,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, 
                           ' de ', año-1, ' a ', format(round(c_spri,0), big.mark = ','),
                           ' millones de pesos corrientes en ', mesp, ' de ',
                           año, '.')
  }else{
    descripcion6 <- paste0(descripcion6, ' Además, la producción mensual ',
                           'de la industria de la construcción estatal ',
                           'contratada por el sector privado se mantuvo ',
                           'en línea con la cifra del mismo mes del año anterior',
                           round(c_spri_var_Anual,1), ' millones de pesos corrientes.')
  }
  
}


descripcion6 <- paste0(descripcion6, ' Con estas cifras el porcentaje que representa ',
                       'el sector público respecto al total del valor generado ',
                       'en la entidad pasó de ',round( por_c_sp_ant_anual,1), '% en ',
                       mesp, ' de ', año-1, ' a ', round(por_c_sp,1), '% en ',
                       mesp, ' de ', año, '.')
mesesAbr <- c('ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP',
              'OCT', 'NOV', 'DIC')

graf_sec <- base_csv %>% filter(ENTIDAD == 'Jalisco') %>% 
  select(c(ANIO, MES, SECTOR_PUBLICO, SECTOR_PRIVADO)) %>%
  mutate(Por_Sector_Publico = (SECTOR_PUBLICO/(SECTOR_PRIVADO+SECTOR_PUBLICO))) %>%
  mutate(MES = mesesAbr[MES]) %>%
  mutate(SECTOR_PUBLICO  = round(SECTOR_PUBLICO/1000,0), SECTOR_PRIVADO = round(SECTOR_PRIVADO /1000,0))

graf_sec[,1] <- fecha_gra[,1]
graf_sec[,2] <- fecha_gra[,2]

# for (i in 1:nrow(graf_sec)) {
#   if(! i %in% seq(1,nrow(graf_sec), 12))
#      graf_sec[i,1] <- " "
# }


########## DATOS POR TIPO DE OBRA (EN MILLONES DE PESOS CORRIENTES) ##########
#Datos mensuales de Jalisco de valor de la produccion por tipo de obra
vp_obra=filter(base_csv,CODIGO_ENTIDAD==14) %>%
  select(ANIO,MES,EDIFICACION,TRANSPORTE_Y_URBANIZACION,AGUA_RIEGO_Y_SANEAMIENTO,ELECTRICIDAD_Y_TELECOMUN,
         PETROLEO_Y_PETROQUIMICA,OTRAS_CONSTRUCCIONES) %>%
  mutate(OTRAS = ELECTRICIDAD_Y_TELECOMUN + PETROLEO_Y_PETROQUIMICA + OTRAS_CONSTRUCCIONES) %>%
  select(-c(ELECTRICIDAD_Y_TELECOMUN,PETROLEO_Y_PETROQUIMICA,OTRAS_CONSTRUCCIONES)) %>% filter(MES==mes) %>%
  mutate(across(c(EDIFICACION:OTRAS),~.x/1000))
# mutate(across(c(EDIFICACION,TRANSPORTE_Y_URBANIZACION,AGUA_RIEGO_Y_SANEAMIENTO,OTRAS),~.x/1000))

#Fecha con formato para graficar
if (mes<10){
  vp_obra$MES=paste0(rep(0,nrow(vp_obra)),vp_obra$MES)
}
vp_obra=mutate(vp_obra,FECHA=paste(vp_obra$ANIO,vp_obra$MES, sep="/"))

#Reorganizar columnas
graf_obra=vp_obra[,c(7,3:6)]




########## TEXTO DESCRIPCIÓN 7: graf_obra ##########
#PENDIENTE
#Datos mensuales de Jalisco de valor de la produccion por tipo de obra
vp_obra=filter(base_csv,CODIGO_ENTIDAD==14) %>%
  select(ANIO,MES,EDIFICACION,TRANSPORTE_Y_URBANIZACION,AGUA_RIEGO_Y_SANEAMIENTO,ELECTRICIDAD_Y_TELECOMUN,
         PETROLEO_Y_PETROQUIMICA,OTRAS_CONSTRUCCIONES) %>%
  mutate(OTRAS = ELECTRICIDAD_Y_TELECOMUN + PETROLEO_Y_PETROQUIMICA + OTRAS_CONSTRUCCIONES) %>%
  select(-c(ELECTRICIDAD_Y_TELECOMUN,PETROLEO_Y_PETROQUIMICA,OTRAS_CONSTRUCCIONES)) %>% filter(MES==mes) %>%
  mutate(across(c(EDIFICACION:OTRAS),~.x/1000))
# mutate(across(c(EDIFICACION,TRANSPORTE_Y_URBANIZACION,AGUA_RIEGO_Y_SANEAMIENTO,OTRAS),~.x/1000))

#Fecha con formato para graficar
if (mes<10){
  vp_obra$MES=paste0(rep(0,nrow(vp_obra)),vp_obra$MES)
}
vp_obra=mutate(vp_obra,FECHA=paste(vp_obra$ANIO,vp_obra$MES, sep="/"))

#Reorganizar columnas
graf_obra=vp_obra[,c(7,3:6)]

#variables a necesitar

Edi_diff <- (vp_obra %>% select(EDIFICACION) %>% .[nrow(.),1])-(vp_obra %>% select(EDIFICACION) %>% .[nrow(.)-1,1])
Edi_var <- ((vp_obra %>% select(EDIFICACION) %>% .[nrow(.),1])/(vp_obra %>% select(EDIFICACION) %>% .[nrow(.)-1,1])-1)*100
T_U_diff <- (vp_obra %>% select(TRANSPORTE_Y_URBANIZACION) %>% .[nrow(.),1])-(vp_obra %>% select(TRANSPORTE_Y_URBANIZACION) %>% .[nrow(.)-1,1])
T_U_var <- ((vp_obra %>% select(TRANSPORTE_Y_URBANIZACION) %>% .[nrow(.),1])/(vp_obra %>% select(TRANSPORTE_Y_URBANIZACION) %>% .[nrow(.)-1,1])-1)*100
Agua_diff <- (vp_obra %>% select(AGUA_RIEGO_Y_SANEAMIENTO) %>% .[nrow(.),1])-(vp_obra %>% select(AGUA_RIEGO_Y_SANEAMIENTO) %>% .[nrow(.)-1,1])
Agua_var <- ((vp_obra %>% select(AGUA_RIEGO_Y_SANEAMIENTO) %>% .[nrow(.),1])/(vp_obra %>% select(AGUA_RIEGO_Y_SANEAMIENTO) %>% .[nrow(.)-1,1])-1)*100


obras_var <- data.frame(
  Tipo = c('edificación',
           'transporte y urbanización',
           'agua, riego y saneamiento'),
  var = c(Edi_var, T_U_var, Agua_var),
  dif = c(Edi_diff, T_U_diff, Agua_diff)
)

obras_var <- obras_var %>% arrange(desc(dif))
prod_construc <- vp_obra[nrow(vp_obra),] %>% mutate(total = EDIFICACION + TRANSPORTE_Y_URBANIZACION + AGUA_RIEGO_Y_SANEAMIENTO + OTRAS) %>%
  select(total) %>% .[1,1]
prod_construc_ant_anual <- vp_obra[nrow(vp_obra)-1,] %>% mutate(total = EDIFICACION + TRANSPORTE_Y_URBANIZACION + AGUA_RIEGO_Y_SANEAMIENTO + OTRAS) %>%
  select(total) %>% .[1,1]

prod_construc_diff <- prod_construc - prod_construc_ant_anual

prod_construc_var_anual <- ((prod_construc/prod_construc_ant_anual)-1)*100

Otras_Produc <- vp_obra %>% select(OTRAS) %>% .[nrow(.),1]

Otras_Produc_Ant <- vp_obra %>% select(OTRAS) %>% .[nrow(.)-1,1]

Otras_Produc_diff <- Otras_Produc - Otras_Produc_Ant

Otras_Produc_var <- ((Otras_Produc/Otras_Produc_Ant)-1)*100

#Texto

descripcion7 <- paste0('La producción de la industria de la construcción ',
                       'en Jalisco presentó ')

if(prod_construc > prod_construc_ant_anual){
  descripcion7 <- paste0(descripcion7, 'un aumento anual de ',
                         round(prod_construc_diff,0), ' millones de pesos corrientes (',
                         round(prod_construc_var_anual,1), '%) en ', mesp, ' de ', año,
                         ',') 
  if(any(c(Edi_var,T_U_var, Agua_var) > 0)){
    descripcion7 <- paste0(descripcion7, ' las obras de ',
                           obras_var[[1,1]], ' presentaron un incremento anual de ',
                           round(obras_var[[1,3]],0), ' millones de pesos corrientes ',
                           '(', round(obras_var[[1,2]],1),'%), de esta forma este tipo de ',
                           'obra destaca en la aportación al crecimiento de la variación global ',
                           'del valor de la producción en la entidad.')
    if(obras_var[[2,3]] > 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un avance anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }else if(obras_var[[2,3]] < 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un retroceso anual de ',
                             abs(round(obras_var[[2,3]],0)), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    } else{
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron una constante anual manteniendose en ',
                             abs(round(obras_var[[2,3]],0)), ' millones de pesos corrientes, mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'tampoco presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }
  }else{
    descripcion7 <- paste0(descripcion7, ' las obras de ',
                           obras_var[[1,1]], ' presentaron un decremento anual de ',
                           abs(round(obras_var[[1,3]],0)), ' millones de pesos corrientes ',
                           '(', round(obras_var[[1,2]],1),'%).')
    if(obras_var[[2,3]] > 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un avance anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }else if(obras_var[[2,3]] < 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un retroceso anual de ',
                             abs(round(obras_var[[2,3]],0)), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    } else{
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron una constante anual manteniendose en ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes, mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'tampoco presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }
  }  
  
  
}else if(prod_construc < prod_construc_ant_anual){
  descripcion7 <- paste0(descripcion7, 'una caída anual de ',
                         abs(round(prod_construc_diff,0)), ' millones de pesos corrientes (',
                         round(prod_construc_var_anual,1), '%) en ', mesp, ' de ', año,
                         ',') 
  if(any(c(Edi_var,T_U_var, Agua_var) > 0)){
    descripcion7 <- paste0(descripcion7, ' las obras de ',
                           obras_var[[1,1]], ' presentaron un incremento anual de ',
                           round(obras_var[[1,3]],0), ' millones de pesos corrientes ',
                           '(', round(obras_var[[1,2]],1),'%), de esta forma este tipo de ',
                           'obra permitió que la variación global ',
                           'del valor de la producción en la entidad cayera a menor tasa.')
    if(obras_var[[2,3]] > 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un avance anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }else if(obras_var[[2,3]] < 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un retroceso anual de ',
                             abs(round(obras_var[[2,3]],0)), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    } else{
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron una constante anual manteniendose en ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes, mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'tampoco presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }
  }else{
    descripcion7 <- paste0(descripcion7, ' las obras de ',
                           obras_var[[1,1]], ' presentaron un decremento anual de ',
                           abs(round(obras_var[[1,3]],0)), ' millones de pesos corrientes ',
                           '(', round(obras_var[[1,2]],1),'%).')
    if(obras_var[[2,3]] > 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un avance anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs( round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }else if(obras_var[[2,3]] < 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un retroceso anual de ',
                             abs(round(obras_var[[2,3]],0)), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%). de esta forma ',
                               'este tipo de obra destaca en la aportación a la caída en el ',
                               'valor de la producción global del estado.')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    } else{
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron una constante anual manteniendose en ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes, mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               abs(round(obras_var[[3,3]],0)), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'tampoco presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }
  }     
  
  
  
  
}else{
  descripcion7 <- paste0(descripcion7, 'en ',mesp, ' de ', año,' una constante anual al mantener el mismo valor ',
                         'de ', round(prod_construc,0), ' respecto al mismo mes del año previo,') 
  
  if(any(c(Edi_var,T_U_var, Agua_var) > 0)){
    descripcion7 <- paste0(descripcion7, ' las obras de ',
                           obras_var[[1,1]], ' presentaron un incremento anual de ',
                           round(obras_var[[1,3]],0), ' millones de pesos corrientes ',
                           '(', round(obras_var[[1,2]],1),'%).')
    if(obras_var[[2,3]] > 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un avance anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }else if(obras_var[[2,3]] < 0){
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron un retroceso anual de ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes (',
                             round(obras_var[[2,2]],1), '%), mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' también presentaron un decremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'no presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    } else{
      descripcion7 <- paste0(descripcion7, ' Por su parte, las obras de ', 
                             obras_var[[2,1]], ' presentaron una constante anual manteniendose en ',
                             round(obras_var[[2,3]],0), ' millones de pesos corrientes, mientras que ')
      if(obras_var[[3,3]] > 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un incremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else if(obras_var[[3,3]] < 0){
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               ' presentaron un decremento de ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes (',
                               round(obras_var[[3,2]],1), '%).')
      }else{
        descripcion7 <- paste0(descripcion7, 'las obras de ', obras_var[[3,1]],
                               'tampoco presentaron variación anual manteniendose en ',
                               round(obras_var[[3,3]],0), ' millones de pesos corrientes.')
      }
    }
  }   
  
}


descripcion7 <- paste0(descripcion7, ' Otras construcciones provocaron ')

if(Otras_Produc > Otras_Produc_Ant){
  
  descripcion7 <- paste0(descripcion7, 'un aumento de ',
                         round(Otras_Produc_diff,0), ' millones de pesos corrientes (',
                         round(Otras_Produc_var,1), '%).')
  
}else if(Otras_Produc < Otras_Produc_Ant){
  
  descripcion7 <- paste0(descripcion7, 'una disminución de ',
                         abs(round(Otras_Produc_diff,0)), ' millones de pesos corrientes (',
                         round(Otras_Produc_var,1), '%).')
}



if(Otras_Produc_diff > obras_var[[1,3]]){
  descripcion7 <- str_remove(descripcion7,
             ', de esta forma este tipo de obra destaca en la aportación al crecimiento de la variación global del valor de la producción en la entidad'
             )
  descripcion7 <- str_sub(descripcion7, 1,str_length(descripcion7)-1)
  descripcion7 <- paste0(descripcion7, ', destacando este tipo de obras en la aportación al crecimiento de la variación global del valor de la producción en la entidad.')
}




vp_obra <- vp_obra[,c(7,3:6)]
names(vp_obra) <- str_replace_all(names(vp_obra), '_', ' ') %>%
  str_to_title() %>% str_replace_all('Y', 'y') %>%
  str_replace_all('Edificacion', 'Edificación') %>% str_replace_all('Transporte y Urbanizacion', 'Transporte y urbanización') %>% 
  str_replace_all('Agua Riego y Saneamiento', 'Agua riego y saneamiento')

########## Exportar a Excel ##########
# Titulo Ficha
titulof = paste0("Producción de la industria de la construcción en Jalisco de ",meses[mes,1]," de ",año)

# Título gráfica 1
tgraf1 = paste0("Valor de la producción de la industria de la construcción en Jalisco, cifras de ",meses[mes,1],
                  " en millones de pesos corrientes,")
#Fuente y titulos
ft=data.frame(c(tgraf1,
                "Valor de la producción de la industria de la construcción en Jalisco, cifras mensuales en millones de pesos corrientes,",
                "Distribución porcentual de la producción de la industria de la construcción por entidad federativa,",
                "Variación porcentual anual de la producción de la industria de la construcción por entidad federativa,",
                "Variación porcentual mensual de la producción de la industria de la construcción por entidad federativa,",
                "Valor de la producción por sector contratante en millones de pesos corrientes y porcentaje que representa el sector público respecto al valor generado en Jalisco,",
                "Valor de la producción por tipo de obra en millones de pesos corrientes,",
                "Fuente: IIEG, con información de INEGI."))
colnames(ft)= "FT"


#Notas
nota=data.frame(c("Nota: El promedio se refiere al de los últimos doce meses.",
                  "Nota: La variación porcentual anual de se refiere a la variación de la producción mensual respecto al mismo mes del año anterior en pesos corrientes con cifras originales.",
                  "Nota: La variación porcentual mensual de se refiere a la variación de la producción mensual respecto al mes inmediato anterior en pesos corrientes con cifras originales.",
                  "Nota: El sector público considera a todas las obras que realiza una empresa constructora por encargo de una dependencia gubernamental en cualquiera de sus tres niveles: Federal, Estatal y Municipal. El sector privado contempla todas las obras que realiza una empresa constructora para cualquier particular o entidad privada.",
                  "Nota: Las obras de edificación, incluyen giros tales como vivienda, edificios industriales, comerciales y de servicios, escuelas, hospitales y clínicas además de obras y trabajos auxiliares para la edificación. Las obras de agua, riego y saneamiento incluyen giros tales como sistemas de agua potable y drenaje, presas y obras de riego, obras y trabajos auxiliares para agua, riego y saneamiento. Las obras de transporte y urbanización incluyen giros tales como obras de transporte en ciudades y urbanización, carreteras, caminos y puentes, obras ferroviarias, infraestructura marítimo y fluvial, obras y trabajos auxiliares para transporte. Las otras obras incluyen giros tales como infraestructura para la generación y distribución de electricidad, infraestructura para telecomunicaciones, refinerías y plantas petroleras, oleoductos y gaseoductos, instalaciones en edificaciones, montaje de estructuras, trabajos de albañilería y acabados, así como obras y trabajos auxiliares para otras construcciones."))
colnames(nota)= "Nota"

#Nombres de variables
nombre_var=data.frame(c("Año","Mes","Fecha","Valor de la producción","Promedio","Variación","Variación promedio","Entidad","Distribución porcentual",
                        "Sector público","Sector privado","Porcentaje sector público","Edificación","Transporte y urbanización","Agua, riego y saneamiento","Otras construcciones"))
colnames(nombre_var)="VARIABLE"

names(graf_mes)=nombre_var[3:4,1]
names(graf_his)=nombre_var[c(1,2,4,5),1]
names(graf_rank_dis)=nombre_var[8:9,1]
names(graf_rank_var)=nombre_var[c(8,6),1]
names(graf_rank_varmen)=nombre_var[c(8,6),1]
names(graf_sec)=nombre_var[c(1,2,10:12),1]
names(graf_obra)=nombre_var[c(3,13:16),1]


wb=createWorkbook("IIEG DIEEF")
addWorksheet(wb, "MES")
titulo=paste(ft[1,1],periodo1)
writeData(wb, sheet=1, titulo, startCol=1, startRow=1)
writeData(wb, sheet=1, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=1, graf_mes, startCol=1, startRow=5)

addWorksheet(wb, "HIS")
titulo=paste(ft[2,1],periodo2)
writeData(wb, sheet=2, titulo, startCol=1, startRow=1)
writeData(wb, sheet=2, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=2, nota[1,1], startCol=1, startRow=3)
writeData(wb, sheet=2, graf_his, startCol=1, startRow=5)

addWorksheet(wb, "RANKDIS")
titulo=paste(ft[3,1],periodo3)
writeData(wb, sheet=3, titulo, startCol=1, startRow=1)
writeData(wb, sheet=3, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=3, graf_rank_dis, startCol=1, startRow=5)

addWorksheet(wb, "RANKVPP")
titulo=paste(ft[4,1],periodo3)
writeData(wb, sheet=4, titulo, startCol=1, startRow=1)
writeData(wb, sheet=4, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=4, nota[3,1], startCol=1, startRow=3)
writeData(wb, sheet=4, graf_rank_var, startCol=1, startRow=5)

addWorksheet(wb, "RANKVMEN")
titulo=paste(ft[5,1],periodo3)
writeData(wb, sheet=5, titulo, startCol=1, startRow=1)
writeData(wb, sheet=5, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=5, nota[4,1], startCol=1, startRow=3)
writeData(wb, sheet=5, graf_rank_varmen, startCol=1, startRow=5)

addWorksheet(wb, "SECTOR")
titulo=paste(ft[6,1],periodo2)
writeData(wb, sheet=6, titulo, startCol=1, startRow=1)
writeData(wb, sheet=6, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=6, nota[4,1], startCol=1, startRow=3)
writeData(wb, sheet=6, graf_sec, startCol=1, startRow=5)

addWorksheet(wb, "OBRA")
titulo=paste(ft[7,1],periodo1)
writeData(wb, sheet=7, titulo, startCol=1, startRow=1)
writeData(wb, sheet=7, ft[8,1], startCol=1, startRow=2)
writeData(wb, sheet=7, nota[5,1], startCol=1, startRow=3)
writeData(wb, sheet=7, vp_obra, startCol=1, startRow=5)


# Texto
addWorksheet(wb, "Texto")
writeData(wb, sheet=8, "Título ficha:", startCol=1, startRow=1)
writeData(wb, sheet=8, titulof, startCol=1, startRow=2)
###
writeData(wb, sheet=8, "MES", startCol=1, startRow=4)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=5)
writeData(wb, sheet=8, descripcion1, startCol=2, startRow=5)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=6)
writeData(wb, sheet=8, paste(ft[1,1],periodo1), startCol=2, startRow=6)
###
writeData(wb, sheet=8, "HIS", startCol=1, startRow=8)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=9)
writeData(wb, sheet=8, descripcion2, startCol=2, startRow=9)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=10)
writeData(wb, sheet=8, paste(ft[2,1],periodo2), startCol=2, startRow=10)
writeData(wb, sheet=8, "Nota:", startCol=1, startRow=11)
writeData(wb, sheet=8, nota[1,1], startCol=2, startRow=11)
###
writeData(wb, sheet=8, "RANKDIS", startCol=1, startRow=13)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=14)
writeData(wb, sheet=8, descripcion3, startCol=2, startRow=14)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=15)
writeData(wb, sheet=8, paste(ft[3,1],periodo3), startCol=2, startRow=15)
###
writeData(wb, sheet=8, "RANKVPP", startCol=1, startRow=17)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=18)
writeData(wb, sheet=8, descripcion4, startCol=2, startRow=18)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=19)
writeData(wb, sheet=8, paste(ft[4,1],periodo3), startCol=2, startRow=19)
writeData(wb, sheet=8, "Nota:", startCol=1, startRow=20)
writeData(wb, sheet=8, nota[2,1], startCol=2, startRow=20)
###
writeData(wb, sheet=8, "RANKVMEN", startCol=1, startRow=22)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=23)
writeData(wb, sheet=8, descripcion5, startCol=2, startRow=23)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=24)
writeData(wb, sheet=8, paste(ft[5,1],periodo3), startCol=2, startRow=24)
writeData(wb, sheet=8, "Nota:", startCol=1, startRow=25)
writeData(wb, sheet=8, nota[3,1], startCol=2, startRow=25)
###
writeData(wb, sheet=8, "SECTOR", startCol=1, startRow=27)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=28)
writeData(wb, sheet=8, descripcion6, startCol=2, startRow=28)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=29)
writeData(wb, sheet=8, paste(ft[6,1],periodo3), startCol=2, startRow=29)
writeData(wb, sheet=8, "Nota:", startCol=1, startRow=30)
writeData(wb, sheet=8, nota[4,1], startCol=2, startRow=30)
###
writeData(wb, sheet=8, "OBRA", startCol=1, startRow=32)
writeData(wb, sheet=8, "Texto:", startCol=1, startRow=33)
writeData(wb, sheet=8, descripcion7, startCol=2, startRow=33)
writeData(wb, sheet=8, "Gráfica:", startCol=1, startRow=34)
writeData(wb, sheet=8, paste(ft[7,1],periodo3), startCol=2, startRow=34)
writeData(wb, sheet=8, "Nota:", startCol=1, startRow=35)
writeData(wb, sheet=8, nota[5,1], startCol=2, startRow=35)


nombre_wb=paste0("ENEC_R-Excel_",año,"_",mes2,".xlsx")
saveWorkbook(wb, nombre_wb, overwrite = TRUE)
