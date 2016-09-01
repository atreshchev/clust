# Multivariate dynamic clusterization approach
# based on adaptive measure DTW+CORT
# including preliminary factor model analysis.

# Mostly russian comments are below.

#=================== ЗАГРУЗКА БИБЛИОТЕК ====================
  library("wmtsa")
  library("splus2R")
  library("ifultools")
  library("MASS")
  library("pdc")
  library("cluster")	# silhouette()
  library("TSclust")
  library("dtw")	# dtw()  
  options(java.parameters = "-Xmx1024m")
  library("XLConnect")	# loadWorkbook(), getSheets(), readWorksheet(), writeWorksheetToFile()
  library("corrplot")	# corrplot()
  library("ape")	# as.phylo()
  library("sparcl")	# ColorDendrogram()
  source("http://addictedtor.free.fr/packages/A2R/lastVersion/R/code.R") # A2Rplot()
  library("ggdendro")	# ggdendrogram()
  library("parallel")

  
#=================== ПОДГОТОВКА ПРОЕКТА ====================  
# rm(list = ls()) # удаление всех переменных
# rm(list = ls()[!grepl("var.name", ls())]) # удаление переменных, имя которых начинается с var.name
# gc()			  # очистка памяти


#===================== НАСТРОЙКИ ПРОЕКТА ====================  
cores.cnt <- detectCores() # Число ядер вычислительной машины

excel.directory	     <- "C:/Users/KR!K/Documents/R"
excel.file_name_full <- "Vars12.6 NFT (Import).xlsx"	# ИД для всех банков (полная выборка)
excel.file_name_good <- "Vars12.6 NGST (Import).xlsx"	# ИД для действующих банков

# Число показателей в рассматриваемой комбинации
Vars.comb.low.cnt  <- 9
Vars.comb.high.cnt <- 9


#================== ИМПОРТ ИСХОДНЫХ ДАННЫХ ==================  
Import.excel <- function(in.dir, in.fname)
{
  setwd(in.dir) # Установка рабочей директории
  excel.book <- loadWorkbook(in.fname, create = FALSE) # Загрузка в ОП книги excel
  
  Vars.name_list     <- getSheets(excel.book)  # Список наименований листов в книге, соответствующий наименованиям показателей
  Vars.name_list.cnt <- length(Vars.name_list) # Количество листов в книге = общее количество показателей

  Vars  <- list() # Вектор данных, содержащий в каждом элементе таблицу исходных данных по одному показателю
  for (i in 1:Vars.name_list.cnt){
    Vars[[i]]  <- readWorksheet(excel.book, sheet = i, header=TRUE)
    Vars[[i]][is.na(Vars[[i]])] <- 0
  } 
  
  result <- list( names = Vars.name_list, count = Vars.name_list.cnt, vars = Vars)
  return(result)
}

# Загрузка исходных данных: Все банки (полная выборка)
Vars.couple <- Import.excel(excel.directory, excel.file_name_good)
# Загрузка исходных данных: Действующие банки
#Vars.couple <-  Import.excel(excel.directory, excel.file_name_good)
Vars		       <- Vars.couple$vars
Vars.name_list     <- Vars.couple$names
Vars.name_list.cnt <- Vars.couple$count
Vars.bank_cnt 	   <- ncol(Vars[[1]]) # Количество банков соответствует количеству столбцов листа
Vars.ts_cnt   	   <- nrow(Vars[[1]]) # Размерность временного ряда - числу строк


#=(1)=========== АНАЛИЗ ВЗАИМОСВЯЗИ ПОКАЗАТЕЛЕЙ =============
# Зам.: Проводится для ИД по действующим банкам (с полными выборками показателей)

#=(1.1) Формирование комбинаций показателей =================
# Вектора данных, содержащий в каждом элементе: детерминант корр. матрицы, список показателей комбинации, корр. матрица
CombList <- list() 		# для всех комбинаций показателей
DetListOfMax <- list()  # для комбинаций с максимальным значением детерминананта среди комбинаций равной мощности

# Формирование комбинаций показателей в различном количестве, 
# построение матриц парных коэф. корреляции показателей и расчет детерминантов
for(Vars.comb.curr.cnt in Vars.comb.low.cnt:Vars.comb.high.cnt)
{
  Vars.comb.list <- combn(1:Vars.name_list.cnt, m = Vars.comb.curr.cnt) # Формирование комбинаций различных показателей в группах по "Vars.comb.curr.cnt" показателей
  Vars.comb.list.cnt <- length(Vars.comb.list[1,]) # Количество получившихся комбинаций

  CorTab.comb.list <- list()	# Массив матриц парных коэф. корреляции показателей для всех комбинаций
  CorTab.comb.det.list <- c()	# Вектор детерминантов матриц парных коэф. корреляции показателей для всех комбинаций
  ###CorTab99.list <- list()
  CurrCntCombList <- list()
  
  # Построение матриц парных коэф. корреляции показателей для каждой комбинации показателей
  for (t in 1:Vars.comb.list.cnt)
  {
    Vars.comb.curr <- Vars.comb.list[,t] # Вектор индексов показателей в текущей комбинации

    CorTab.comb.curr <- matrix(0, nrow = Vars.comb.curr.cnt, ncol = Vars.comb.curr.cnt) # Задание корр. матрицы
	###CorTab99 <- CorTab95 <- CorTab <- matrix(0, nrow = Vars.comb.curr.cnt, ncol = Vars.comb.curr.cnt) # Задаем корр.матрицу	 

    # Вычисление матриц парных коэф. корреляции по банкам и их обощение
	for (i in 1:Vars.bank_cnt)
	{ 
	  bankTab.temp <- matrix(0, nrow = Vars.ts_cnt, ncol = Vars.comb.curr.cnt)

	  for (j in 1:Vars.comb.curr.cnt){
		bankTab.temp[,j] <- Vars[[Vars.comb.curr[j]]][,i] 
	  }  

	  CorTab.bank.temp <- cor(bankTab.temp, use = "complete.obs")
	  ###CorTab99.temp <- cor.mtest(bankTab.temp, 0.99)

	  # Вычисление средних значений коэффициентов корреляции
	  CorTab.comb.curr <- CorTab.comb.curr + CorTab.bank.temp
	  ###CorTab99 <- CorTab99 + abs(CorTab99.temp)
	}

	# Актуализация среднего значения коэффициентов корреляции
	CorTab.comb.list[[t]] <- CorTab.comb.curr/Vars.bank_cnt
	###CorTab99.list[[t]] <- CorTab99/Vars.bank_cnt

	# Оценка мультиколлинеарности (взаимозависимости) показателей для текущей комбинации
	CorTab.comb.det.list[[t]] <- det(CorTab.comb.list[[t]])

	# Задание имен строк и столбцов матрицы в соответствии с рассматриваемой комбинацией
	Vars.name_list.curr <- c()
	for (k in Vars.comb.curr) {
	  Vars.name_list.curr <- cbind(Vars.name_list.curr, Vars.name_list[k])
	} 
	colnames(CorTab.comb.list[[t]]) <- Vars.name_list.curr
	rownames(CorTab.comb.list[[t]]) <- Vars.name_list.curr
	###colnames(CorTab99.list[[t]]) <- Vars.name_list.curr
	###rownames(CorTab99.list[[t]]) <- Vars.name_list.curr

	# Заполнение информации оценки качества по текущей комбинации показателей
	CombList <- rbind(CombList, list(det = CorTab.comb.det.list[[t]], vars = Vars.comb.curr, cortab = CorTab.comb.list[[t]]))
	CurrCntCombList <- rbind(CurrCntCombList, list(det = CorTab.comb.det.list[[t]], vars = Vars.comb.curr, cortab = CorTab.comb.list[[t]]))
  }
  
  DetListOfMax <- rbind(DetListOfMax, c(CurrCntCombList[which.max(CurrCntCombList[,1]),]))
}

# Определение комбинации с максимальным детерминантом корр. матрицы
CombList.maxdet <- CombList[which.max(CombList[,1]),]
CombList.maxdet

#=(1.2) Визуализация ========================================
# Визуализация 1: Диаграмма детерминантов корр. матриц (по всем иследованным комбинациям показателей)
combnum <- detval <- c()
for(i in 1:length(CombList[,1])){
  detval <- c(detval, as.numeric(CombList[i,1]))
  combnum <- c(combnum, i)
}
barplot(detval, names.arg = combnum, xlab = "Комбинация показателей, №", ylab = "Отсутствие мультиколлинеарности")
print(CombList.maxdet)

# Визуализация 2: Матрицы парных коэф. показателей для различных комбинаций + Диаграмма детерминантов корр. матриц
# Зам.: Демонстрируется для комбинаций показателей при Vars.comb.low.cnt = Vars.comb.high.cnt = 8
par(mfrow = c(3, 4)) # Формирование шаблона под диаграммы в 3 строки * 4 столбца
for(t in 1:Vars.comb.list.cnt){
  corrplot(CorTab.comb.list[[t]],
			p.mat = CorTab.comb.list[[t]],
			sig.level=-1, insig = "p-value", 
			method="color", tl.cex = 1.0,
			cl.cex = 1.0,
			addrect = 10,
			cl.pos="b",
			main=t)
}
barplot(CorTab.comb.det.list, names.arg = 1:Vars.comb.list.cnt, xlab = "Комбинация показателей", ylab = "Отсутствие мультиколлинеарности")


#=(2)=========== МНОГОМЕРНАЯ КЛАСЕРИЗАЦИЯ ==============
# Динамическая многомерная кластеризация DTW+CORT (одновременно по множеству показателей)

#=(2.1) MDTW ===========================================
# Функция расчета интегрального расстояния DTW для 2-х объектов (банков)
mdtw.func <- function(in.i, in.j, in.Vars.ts_cnt, in.Vars.comb.curr, in.Vars)
{
  #print(c(ts_cnt = in.Vars.ts_cnt, i = in.i, j = in.j))
  if(in.i >= in.j) {
    return(0)
  }else{
    # Матрица покоординатных расстояний для двух объектов
    M.cost <- matrix(0, in.Vars.ts_cnt, in.Vars.ts_cnt)
    
    for(k in 1:in.Vars.ts_cnt) {
      for(n in 1:in.Vars.ts_cnt) {
        for(ind in in.Vars.comb.curr){
          M.cost[k,n] <- M.cost[k,n] + abs(in.Vars[[ind]][k,in.i] - in.Vars[[ind]][n,in.j])
        }
      }                             
    }
    # Интегральное расстояние DTW для двух объектов
    mdtw.temp <- dtw(M.cost)

    return(mdtw.temp$distance)
  }
}

# Функция построения сводной матрицы расстояний DTW для всех объектов
MDTW.main_func <- function(in.cl, in.Vars.bank_cnt, in.Vars.ts_cnt, in.Vars.comb.curr, in.Vars)
{
  D.mdtw <- matrix(0, in.Vars.bank_cnt, in.Vars.bank_cnt)
  
  for(bank_ind in 1:in.Vars.bank_cnt) {
    # Задание параметров кластера
    varlist <- c("in.Vars.bank_cnt","j","bank_ind","in.Vars.ts_cnt","in.Vars.comb.curr","in.Vars","mdtw.func")
    environment(mdtw.func) <- .GlobalEnv
    
    clusterExport(in.cl, varlist = varlist, envir = environment()) 
    clusterEvalQ(in.cl, library("dtw"))
    
    # Получение вектора значений интегрального расстояния DTW для 2-х объектов (банков)
    mdtw.func.row <- parSapply(in.cl, 1:in.Vars.bank_cnt, function(j) mdtw.func(bank_ind, j, in.Vars.ts_cnt, in.Vars.comb.curr, in.Vars))
    
    # Копирование элементов списка в строку матрицы
    for(col_ind in 1:in.Vars.bank_cnt){
      D.mdtw[bank_ind, col_ind] <- mdtw.func.row[[col_ind]] 
    }
	
	# Оценка прогресса выполнения алгоритма
    print(paste("MDTW Progress:",round(bank_ind/in.Vars.bank_cnt*100),"%"))
  }
  
  # Матрица D.tdw симметричная
  for(x in 1:(in.Vars.bank_cnt-1)){
    for(y in x:in.Vars.bank_cnt){
      D.mdtw[y,x] <- D.mdtw[x,y]
    }
  }
  return(D.mdtw)
}

# Обрамление параллельного расчета MDTW для комбинаций показателей с максимальным детерминантом
MDTW <- function(in.DetListOfMax, in.Vars.bank_cnt, in.Vars.ts_cnt, in.Vars, in.cores.cnt)
{
  # Создание кластера из n ядер
  cl <- makeCluster(in.cores.cnt) # getOption("cl.cores", 16)

  MDTW.matrix.list <- list()

  # Проход по всем наборам параметрам с максимальным коэф.мультиколлинеарности среди сочетаний равной мощности
  for(DetListOfMax.curr.cnt in 1:length(in.DetListOfMax[,1])){
    # Выбор одного из наборов параметров
    DetListOfMax.curr <- in.DetListOfMax[DetListOfMax.curr.cnt,]

    # Построение сводной матрицы расстояний MDTW для всех объектов по заданному набору параметров
    MDTW.curr <- MDTW.main_func(cl, in.Vars.bank_cnt, in.Vars.ts_cnt, DetListOfMax.curr$vars, in.Vars)
    MDTW.matrix.list <- rbind(MDTW.matrix.list, list(vars = DetListOfMax.curr, D.mdtw = MDTW.curr))

    print(paste("MDTW DetListOfMax Progress:",round(DetListOfMax.curr.cnt/length(in.DetListOfMax[,1])*100),"%"))
  }

  # Удаление кластера
  stopCluster(cl)

  return(MDTW.matrix.list)
}

# Выполнение параллельного расчета MDTW
time.set <- system.time(
  MDTW <- MDTW(DetListOfMax, Vars.bank_cnt, Vars.ts_cnt, Vars, cores.cnt)
)
print(time.set)

#=(2.2) MCORT ============================================
# Функция расчета расстояния CORT для 2-х объектов (банков)
mcort.func <- function(in.i, in.j, in.DetList.max.vars, in.Vars)
{ 
  if(in.i>=in.j) {
    return(1)
  } else {
    multi <- norm <- 0
	
    for(par_ind in in.DetList.max.vars){
      multi <- multi + sum( diff(ts(in.Vars[[par_ind]][,in.i],start=1)) * diff(ts(in.Vars[[par_ind]][,in.j],start=1)) )
      norm  <- norm  + sqrt(sum(diff(ts(in.Vars[[par_ind]][,in.i],start=1))^2)) * sqrt(sum(diff(ts(in.Vars[[par_ind]][,in.j],start=1))^2))
    }
  }
  return(multi/norm)
}     

# Функция построения сводной матрицы расстояний CORT для всех объектов  
MCORT.main_func <- function(in.cl, in.Vars.bank_cnt, in.DetList.max.vars, in.Vars)
{     
  D.mcort <- matrix(0, in.Vars.bank_cnt, in.Vars.bank_cnt)

  for(bank_ind in 1:in.Vars.bank_cnt) {
    # Задание параметров кластера
    varlist <- c( "bank_ind", "j", "in.Vars.bank_cnt", "in.DetList.max.vars", "in.Vars", "mcort.func" )
    environment(mcort.func) <- .GlobalEnv

    clusterExport( in.cl, 
                   varlist = varlist,
                   envir = environment() ) 
    #clusterEvalQ( in.cl, library("dtw") )
    # Получение вектора значений CORT 
    mcort.func.row <- parSapply(in.cl, 1:in.Vars.bank_cnt, function(j) mcort.func(bank_ind, j, in.DetList.max.vars, in.Vars))
    
    # Копирование элементов списка в строку матрицы
    for(col_ind in 1:in.Vars.bank_cnt){
      D.mcort[bank_ind,col_ind] <- mcort.func.row[[col_ind]] 
    }

    # Оценка прогресса выполнения алгоритма
    print(paste("MCORT Progress:",round(bank_ind/in.Vars.bank_cnt*100),"%"))
  }

  # Матрица D.CORT симметричная
  for(x in 1:(in.Vars.bank_cnt-1)){
    for(y in x:in.Vars.bank_cnt){
      D.mcort[y,x] <- D.mcort[x,y]
    }
  } 
  return(D.mcort)
}

# Обрамление параллельного расчета CORT для комбинаций показателей с максимальным детерминантом
MCORT <- function(in.DetListOfMax, in.Vars.bank_cnt, in.Vars, in.cores.cnt)
{
  # Создание кластера из n ядер
  cl <- makeCluster(in.cores.cnt)

  MCORT.matrix.list <- list()

  # Проход по всем наборам параметрам с максимальным детерминантом среди комбинаций равной мощности
  for(DetListOfMax.curr.cnt in 1:length(in.DetListOfMax[,1])){
    DetListOfMax.curr <- in.DetListOfMax[DetListOfMax.curr.cnt,]
    
    # Построение сводной матрицы расстояний CORT для всех объектов по заданному набору параметров
    MCORT.curr <- MCORT.main_func(cl, in.Vars.bank_cnt, DetListOfMax.curr$vars, in.Vars)

	# Построение списка записей вида: детерминант определенного набора параметров, соответствующая этому набору матрица CORT
	MCORT.matrix.list <- rbind(MCORT.matrix.list, list(vars = DetListOfMax.curr, D.mcort = MCORT.curr))

    # Оценка прогресса выполнения алгоритма
    print(paste("MCORT DetListOfMax Progress:",round(DetListOfMax.curr.cnt/length(DetListOfMax[,1])*100),"%"))
  }

  # Удаление кластера
  stopCluster(cl)

  return(MCORT.matrix.list)
}

# Выполнение параллельного расчета MCORT
time.set <- system.time(
  # Список матриц MCORT для "лучших" комбинаций показателей
  MCORT <- MCORT(DetListOfMax, Vars.bank_cnt, Vars, cores.cnt)
)
print(time.set)

#=(2.3) DTW+MCORT ========================================
# Вычисление итоговой меры расстояния
k = 4.5 # Входной параметр: степень чувсвительности к поведению

MDTWCORT <- list()
for(DetListOfMax.curr.cnt in 1:length(DetListOfMax[,1])){
  DetListOfMax.curr <- DetListOfMax[DetListOfMax.curr.cnt,]

  D.mdtw <- MDTW[DetListOfMax.curr.cnt,2]
  D.cort <- MCORT[DetListOfMax.curr.cnt,2]
  D.mdtwcort <- as.dist((2/(1+exp(k*D.cort[[1]])))*D.mdtw[[1]])

  MDTWCORT <- rbind(MDTWCORT, list(vars = DetListOfMax.curr, D.mdtwcort = D.mdtwcort))
}

#=(2.4) Иерархическая кластеризация =========================
# Входные параметры
clustmethod <- "ward.D2" # Расстояние между объектами (вариант: "average") 
clustnum <- 50			# Число желаемых кластеров

HC <- list()
for(DetListOfMax.curr.cnt in 1:length(DetListOfMax[,1])){
  hc.mdtwcort <- hclust(MDTWCORT[DetListOfMax.curr.cnt,]$D.mdtwcort, clustmethod)
  plot(hc.mdtwcort)
  clusters.mdtwcort <- cutree(hc.mdtwcort, k = clustnum)
  silh <- silhouette(clust.mdtwcort, MDTWCORT[DetListOfMax.curr.cnt,]$D.mdtwcort)
  
  HC <- rbind(HC, list(vars = DetListOfMax[1,]$vars, hc = hc.mdtwcort, clusters = clusters.mdtwcort, silh = silh))
}

#=(2.5) Визуализация ========================================
HC[1,]$clusters
par(mfrow = c(1, 1)) # Формирование шаблона под диаграммы в 1 строку * 1 столбец
plot(HC[1,]$hc)

# 1 - Вывод дендограммы с выделением кластеров
ColorDendrogram(HC[1,]$hc, y = HC[1,]$clusters, labels = rownames(C), main = "Дендрограмма", branchlength = 80, cex = 2.0) # cex - размер шрифта заголовка

# 2.1 - Вывод дендограммы без подписей объектов, выровненной, без выделения кластеров
ggdendrogram(HC[1,]$hc)
ggdendrogram(HC[1,]$hc, rotate = FALSE, size = 1, theme_dendro = TRUE, color = "black")

# 2.2 - Вывод дендограммы без подписей объектов, выровненной, с выделением кластеров заданным цветом
col.down = c("aquamarine","cornflowerblue","bisque","blueviolet","brown","cadetblue","coral",
             "cyan","darkblue","dodgerblue","darkorange","darkturquoise","darkviolet","deeppink","deepskyblue",
             "firebrick","floralwhite","forestgreen","gold","goldenrod","gray","green","greenyellow","honeydew")
			 "hotpink","indianred","indigo","lightgreen","lightpink","lightsalmon","lightseagreen","lightskyblue",
		     "mediumvioletred","midnightblue","mintcream","mistyrose","orange","orangered","orchid","palegoldenrod",
			 "palegreen","paleturquoise","palevioletred","purple","red","rosybrown","royalblue","seashell",
			 "springgreen","yellow","thistle") # Цвета для 52 ветвей
op = par(bg = "ghostwhite") # Цвет фона
A2Rplot(hcMV, k = clustnum, boxes = FALSE, col.up = "gray50", col.down)      



#============= ЭКСПОРТ РЕЗУЛЬТАТОВ КЛАСТЕРИЗАЦИИ =============
# Выполняется после расчетов матриц расстояний D.mdtw, D.mcort
writeWorksheetToFile(paste("D.mdtw",toString(DetListOfMax[1,]$vars),".xlsx"), data=as.matrix(MDTW[1,]$D.mdtw), sheet=paste("MDtw",toString(DetListOfMax[1,]$vars)), startRow=1, startCol = 1)
writeWorksheetToFile(paste("D.mcort",toString(DetListOfMax[1,]$vars),".xlsx"), data=as.matrix(MCORT[1,]$D.mcort), sheet=paste("MCort",toString(DetListOfMax[1,]$vars)), startRow=1, startCol = 1)

# Выполняется после иерархической кластеризации
writeWorksheetToFile(paste("clusters",toString(DetListOfMax[1,]$vars),".xlsx"), data=as.matrix(HC[1,]$clusters), sheet=paste("clusters",toString(DetListOfMax[1,]$vars)), startRow=1, startCol = 1)


#=(3)======== ГЕНЕРАЦИЯ ЛУЧШЕЙ МОДЕЛИ КЛАСТЕРИЗАЦИИ ==========
#================== ИМПОРТ МАТРИЦ РАССТОЯНИЙ =================
D.mdtw7 <- readWorksheet(loadWorkbook("D.mdtw 3, 4, 5, 6, 7, 8, 9 .xlsx", create = FALSE), sheet = 1, header = TRUE)
D.mcort7 <- readWorksheet(loadWorkbook("D.mcort 3, 4, 5, 6, 7, 8, 9 .xlsx", create = FALSE), sheet = 1, header = TRUE)
D.mdtw8 <- readWorksheet(loadWorkbook("D.mdtw 2, 3, 4, 5, 6, 7, 8, 9 .xlsx", create = FALSE), sheet = 1, header = TRUE)
D.mcort8 <- readWorksheet(loadWorkbook("D.mcort 2, 3, 4, 5, 6, 7, 8, 9 .xlsx", create = FALSE), sheet = 1, header = TRUE)

# Варианты параметров модели кластеризации (всего 108 вариантов)
varcomb <- list()
varcomb <- rbind(varcomb, list(vars = "3 4 5 6 7 8 9", D.mdtw = D.mdtw7, D.mcort = D.mcort7))
varcomb <- rbind(varcomb, list(vars = "2 3 4 5 6 7 8 9", D.mdtw = D.mdtw8, D.mcort = D.mcort8))
k <- c(1.5, 3, 4.5)
clustmethod <- c("average", "ward.D2")
clustnum <- c (30, 35, 40, 45, 50, 55, 60, 65, 70)

# Построение различных моделей кластеризации M DTW+CORT с оценкой их качества
CM <- list()
for(i in 1:length(varcomb[,1])){ 
  for(j in 1:length(k)){ 
    for(m in 1:length(clustmethod)){  
      for(l in 1:length(clustnum)){
		# Вычисление итоговой меры расстояния с учетом: комбинаций показателей(varcomb[i,]) и коэф. k
		D.mdtwcort <- as.dist((2/(1+exp(k[j]*varcomb[i,]$D.mcort)))*varcomb[i,]$D.mdtw)
		# Вычисление иерархического кластера с учетом: метода clustmethod
		hc.mdtwcort <- hclust(D.mdtwcort, clustmethod[m])
		# Построение кластеров объектов с учетом: количества кластеров clustnum
		clusters.mdtwcort <- cutree(hc.mdtwcort, k = clustnum[l])
		
		# Оценка качества кластеризации (силуэта кластеров)
		avg.silh <- summary(silhouette(clusters.mdtwcort, D.mdtwcort, do.clus.stat = TRUE))$avg.width
		
		CM <- rbind(CM, list(vars = toString(varcomb[i,]$vars), k = k[j], clustmethod = clustmethod[m], clustnum = clustnum[l],
							  D.mdtw = varcomb[i,]$D.mdtw, D.mcort = varcomb[i,]$D.mcort, D.mdtwcort = D.mdtwcort,
							  hc = hc.mdtwcort, clusters = clusters.mdtwcort, silh = avg.silh))
		#print(paste(toString(varcomb[i,]$vars), k[j], clustmethod[m], clustnum[l], silh))
	  }
	}
  }
}

# Вывод данных для анализа
CMtoExport <- CM[,which(!names(CM[1,]) %in% c("D.mdtw","D.mcort","D.mdtwcort","hc","clusters"))]
writeWorksheetToFile(paste("CMtoExport.xlsx"), data=CMtoExport, sheet="CMtoExport", startRow=1, startCol = 1)

# Определение наиболее качественной модели кластеризации
BestCMFull <- CM[which.max(CM[,10]),]	  # по максимальному значению силуэта (10-й элемент вектора характеристик модели)
BestCM	   <- CMtoExport[which.max(CM[,10]),]
print(BestCM)

# Вывод состава кластера для анализа
BestClusters <- CM[which.max(CM[,10]),]$clusters
writeWorksheetToFile(paste("BestCl",toString(BestCM$vars),BestCM$k,BestCM$clustmethod,BestCM$clustnum,".xlsx"), data=BestClusters, sheet="BestClusters", startRow=1, startCol = 1)


#=(*)======== Вывод результатов для загружаемых файлов ==========
## 1 ########################
clustmethod <- "ward.D2" # Расстояние между объектами (вариант: "average") 
clustnum <- 70			# Число желаемых кластеров
hc.mdtwcort <- hclust(BestCM$D.mdtwcort, clustmethod)
plot(hc.mdtwcort)

## 2 ########################   
clusters.mdtwcort <- cutree(hc.mdtwcort, k = clustnum)
ColorDendrogram(hc.mdtwcort, y = clusters.mdtwcort, labels = rownames(C), main = "Дендрограмма", branchlength = 80, cex = 2.0) # cex - размер шрифта заголовка

## 3 ########################   
ggdendrogram(hc.mdtwcort, rotate = FALSE, size = 1, theme_dendro = TRUE, color = "black")

## 4 ######################## 
col.down = c("aquamarine","cornflowerblue","bisque","blueviolet","brown","cadetblue","coral",
             "cyan","darkblue","dodgerblue","darkorange","darkturquoise","darkviolet","deeppink","deepskyblue",
             "firebrick","floralwhite","forestgreen","gold","goldenrod","gray","green","greenyellow","honeydew",
			 "hotpink","indianred","lightgreen","lightpink","lightsalmon","lightseagreen","lightskyblue",
		     "mediumvioletred","midnightblue","mintcream","mistyrose","orange","orangered","orchid","palegoldenrod",
			 "palegreen","paleturquoise","palevioletred","purple","red","rosybrown","royalblue","seashell",
			 "springgreen","yellow","thistle",
             "darkviolet","deepskyblue","gold","greenyellow","indianred","lightgreen","lightseagreen","mediumvioletred","midnightblue","forestgreen","brown","cadetblue","coral","cyan","darkblue","palevioletred","purple","red","darkorange","darkorchid","darkred","darksalmon","darkseagreen","darkslateblue","darkslategray","darkturquoise","darkviolet",
             "firebrick","floralwhite","forestgreen","gold","goldenrod","gray","green","greenyellow","honeydew",
			 )# Цвета для 52 ветвей
op = par(bg = "ghostwhite") # Цвет фона
A2Rplot(hc.mdtwcort, k = clustnum, boxes = FALSE, col.up = "gray50", col.down)      

#============================================================
