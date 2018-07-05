rm(list=ls())

setwd("")

if(!"readxl"%in%installed.packages()[,1]) install.packages("readxl")
if(!"tidyr"%in%installed.packages()[,1]) install.packages("tidyr")
if(!"openxlsx"%in%installed.packages()[,1]) install.packages("openxlsx")

library("readxl")
data_clients = read_xlsx("stats_store_data.xlsx",sheet = "clients")
data_cases = read_xlsx("stats_store_data.xlsx",sheet = "cases",col_types = "text")

#Seperating the cases in the clients sheet, so that every case has his own row
library("tidyr")
data_clients = separate_rows(data_clients,Case,sep=",")
colnames(data_clients)[6] = "CaseID"

#Merging the sheets by CaseID, maybe it can be merged by NameClient as well, you should try
data = merge(x = data_cases, y = data_clients, by = "CaseID", all = TRUE)

#The Meeting column was a real mess, here are the dates stored of every meeting. 
#I made a clear column with only single date-values. If there were more meeting-dates,
#I only saved the first one, because only that one is interesting for us I believe.
library("openxlsx")
for(i in 1:nrow(data)){
  if(grepl(",",data$Meeting[i])){
    data$Meeting[i] = as.character(as.Date(as.character(gsub("(.*),.*", "\\1",data$Meeting[i])),format="%d/%m/%Y"))
  }else{
    data$Meeting[i] = as.character(convertToDate((as.integer(data$Meeting[i])+1462)))
  }
}

#created multiple columns with seperate topics
library("stringr")
topics = c()
for(i in 1:nrow(data)){
  topics[i] = str_count(data$IsAbout[i],",")
}
about = str_split_fixed(data$IsAbout, ",",max(topics,na.rm = T)+1)
about = as.data.frame(about)
for(i in 1:ncol(about)){
  colnames(about)[i] = paste("IsAbout",i)
}
data = cbind(data[,c(1,2)],about,data[,4:11])
data = data[,c(1,11,3:10,12:15)]
colnames(data)[2] = "NameClient"

#creating the columns with topics and frequencies of those topics
topics = names(sort(table(as.matrix(about)),decreasing = T))
topicsfreq = unname(sort(table(as.matrix(about)),decreasing = T))
topicscomb = cbind(topics,topicsfreq)
topicscomb = as.data.frame(topicscomb)
topicscomb = topicscomb[topicscomb$topics != "",]
rownames(topicscomb) = NULL
empty = nrow(data)-nrow(topicscomb)
empty1 = as.data.frame(matrix(NA,nrow=empty,ncol = 2))
topicscomb = as.data.frame(rbind(as.matrix(topicscomb),as.matrix(empty1)))
data$topics = topicscomb$topics
data$topicsfreq = topicscomb$topicsfreq
data$topics = as.character(data$topics)
data$topicsfreq = as.numeric(data$topicsfreq)

# creating the coded department variables
# should just work bolted on at the end
department = as.matrix(data$Group)
unique = unique(department)

# 1 = BC; 2 = Clinical; 3 = Developmental; 4 = Social; 5 = WO; 6 = PML; 7 = Other.
classify = function(x) {
  if (x == unique[1]) {
    x = 1
  } else if (x == unique[2]) {
    x = 2
  } else if (x == unique[3]) {
    x = 3
  } else if (x == unique[4]|x == unique[10]) {
    x = 4
  } else if (x == unique[5]) {
    x = 5
  } else if (x == unique[6] | x == unique[8] | x == unique[9]) {
    x = 8
  } else if (x == unique[7]) {
    x = 6
  } else 
    x = 8
}

test.matrix = as.matrix(data$Group)
test = as.matrix(apply(test.matrix, 1, classify), ncol = 1)
data$group.classify = test

classify2 = function(x) {
  if (x == 1 | x == 2 | x == 3 | x == 4 | x == 5 | x == 6) {
    x = "Psychology"
  } else
    x = "Other"
}

# 0 = not psychology; 1 = psychology
test.matrix2 = as.matrix(data$group.classify)
test2 = as.matrix(apply(test.matrix2, 1, classify2), ncol = 1)
data$psychology = test2

#creating the background columns
table.client <- sort(table(data$TypeClient), decreasing = TRUE)
name.client.freq <- c(names(table.client)[1:7], "Other")
num.clients <- as.vector(c(table.client[1:7], sum(table.client[8:length(table.client)])))
background = cbind(name.client.freq,num.clients)
empty2 = nrow(data)-nrow(background)
empty3 = matrix(NA,ncol=2,nrow=empty2)
background = rbind(background,empty3)
data$background = background[,1]
data$backgroundfreq = background[,2]

#creating column of percentage of clients who have multiple cases
come.back <- table(data$NameClient) > 1
perc = length(come.back[come.back==TRUE]) / length(unique(data$NameClient)) * 100
data$perc.multiplecases = c(perc,rep(NA,nrow(data)-1))

# Combining certain topics to reduce 'other'

colnames = grep("IsAbout", names(data), value = TRUE)
max.col = length(colnames) 
max.row = nrow(data)

isabout = as.matrix(data[,colnames])

combine = function(search, replace, data, ncol = max.col, nrow = max.row) {
  my.matrix = matrix(NA, nrow, ncol)
  for(i in 1:ncol) {
    my.matrix[,i] = replace(data[,i], data[,i] == search, replace)
  }
  return(my.matrix)
}


# Combining all regressions - (GLM would be too vague I think)
test = isabout
index = grep("egression", test)

isabout = combine("Mixed regression (random & fixed effects)", "Regression", isabout)
isabout = combine("Logistic regression", "Regression", isabout)
isabout = combine("Logistic regression", "Regression", isabout)
isabout = combine("Ordinal regression", "Regression", isabout)
isabout = combine("extended generalized regression models", "Regression", isabout)
isabout = combine("autoregression", "Regression", isabout)

# Combinings ANOVAS & MANOVAS as "ANOVA/MANOVA"
test = isabout
index = grep("NOVA", test)

isabout = combine("RM ANOVA", "ANOVA/MANOVA", isabout)
isabout = combine("MANOVA", "ANOVA/MANOVA", isabout)
isabout = combine("ANOVA", "ANOVA/MANOVA", isabout)

# Combining: "Factor analysis" | "Structural Equation Modeling (SEM)" | "SEM" | "PCA & cluster analysis" | "PCA"
# as "Psychometrics"
isabout = combine("Factor analysis", "Psychometrics", isabout)
isabout = combine("Structural Equation Modeling (SEM)", "Psychometrics", isabout)
isabout = combine("SEM", "Psychometrics", isabout)
isabout = combine("PCA & cluster analysis", "Psychometrics", isabout)
isabout = combine("PCA", "Psychometrics", isabout)

# Combining ANCOVAS & MANCOVAS & Repeated Measures ANCOVA as "ANCOVA/MANCOVA" 
# "ANCOVA" | "RM ANCOVA"
test = isabout
index = grep("COVA", test)

isabout = combine("RM ANCOVA", "ANCOVA/MANCOVA", isabout)
isabout = combine("ANCOVA", "ANCOVA/MANCOVA", isabout)
isabout = combine("MANCOVA", "ANCOVA/MANCOVA", isabout)

data[,colnames] = isabout

# Add 2 columns
# Creates 2 columns in the data that returns the top 11 subjects + other and their respective numbers

sorted = sort(table(as.matrix(data[,3:(max.col+2)]))) 
sorted = sorted[1:(length(sorted)-1)]

current.11 = tail(sorted, 11)
remaining = sorted[1:(length(sorted)-11)] 

Other = sum(remaining)

current.11 = as.matrix(current.11, ncol = 1)
pie = rbind(current.11, Other)

Names = row.names(pie); row.names(pie) = NULL
Numbers = pie[,1]

pie2 = as.data.frame(matrix(ncol = 2, nrow = max.row))
for(i in 1:length(Names)) {
  pie2[i,1] = Names[i]
  pie2[i,2] = Numbers[i]
}

colnames(pie2) = c("Topics", "Number of inquiries")
data = cbind(data, pie2)

data$Meeting = format(as.Date(data$Meeting),"%Y-%m-%d")

#creating columns with frequency meeting dates in year and month
years = as.data.frame(table(substr(data$Meeting,1,4)))
colnames(years) = c("year","freq")
months = as.data.frame(table(substr(data$Meeting,1,7)))
colnames(months) = c("month","freq")
empty4 = nrow(data)-nrow(years)
empty5 = matrix(NA,ncol=2,nrow=empty4)
years1 = rbind(as.matrix(years),empty5)
data$years = years1[,1]
data$yearsfreq = years1[,2]
empty6 = nrow(data)-nrow(months)
empty7 = matrix(NA,ncol=2,nrow=empty6)
months1 = rbind(as.matrix(months),empty7)
data$months = months1[,1]
data$monthsfreq = months1[,2]

months.total = as.data.frame(table(substr(data$Meeting,6,7)))
months.total = cbind(months.total,rep(NA,nrow(months.total)))
colnames(months.total) = c("monthnum", "freq","month.total")
months.total$month.total = month.name[months.total$monthnum]
months.total = months.total[,c(3,2)]
empty8 = nrow(data)-nrow(months.total)
empty9 = matrix(NA,ncol=2,nrow=empty8)
months.total1 = rbind(as.matrix(months.total),empty9)
data$montotal = months.total1[,1]

#creating average cases per month minus first and last month (because they are maybe incomplete and therefore may cause a biased statistic)
first.month = sort(data$months)[1]
first.date = min(data$Meeting[!grepl(first.month,data$Meeting)])
last.month = sort(data$months,decreasing = T)[1]
last.date = max(data$Meeting[!grepl(last.month,data$Meeting)])
full.years = (as.numeric(substr(last.date,1,4))-as.numeric(substr(first.date,1,4)))-1
full.months = 13 - as.numeric(substr(first.date,6,7)) + as.numeric(substr(last.date,6,7))

if((full.months-12)==12){
  full.years = full.years + 2
  full.months = full.months - 24
}else if((full.months-12)>=0 & (full.months-12)<12){
  full.years = full.years + 1
  full.months = full.months - 12
}else{
  full.years = full.years
  full.months = full.months
}

start.month = as.numeric(substr(first.date,6,7))
month.freq = rep(0,12)
month.freq = month.freq + full.years
extra.months1 = start.month:(start.month+full.months-1)
extra.months2 = extra.months1[extra.months1>12]-12
extra.months = c(extra.months2,extra.months1[extra.months1<=12])
month.freq[extra.months] = month.freq[extra.months] + 1
monthlymeeting.freq = as.data.frame(table(substr(data$Meeting[(substr(data$Meeting,1,7)!=first.month) & (substr(data$Meeting,1,7)!=last.month)],6,7)))
data$monavg = c(monthlymeeting.freq$Freq/month.freq,rep(NA,nrow(data)-12))

#changing else to other in group column
data$Group[data$Group == "else"] = "Other"

#creating the prognosis
if(!"lubridate"%in%installed.packages()[,1]) install.packages("lubridate")
library("lubridate")
lastyear = max(substr(data$Meeting,1,4))
diyear = yday(max(data$Meeting[substr(data$Meeting,1,4)==lastyear]))
mlyear = length(data$Meeting[substr(data$Meeting,1,4)==lastyear])
prevyears = unique(substr(data$Meeting,1,4))
prevyearsl = length(unique(substr(data$Meeting,1,4)))-1
prevyears = prevyears[2:prevyearsl]
mp.year = c()
prevmeet = c()
for(i in 1:(prevyearsl-1)){
  mp.year[i] = length(data$Meeting[substr(data$Meeting,1,4)==prevyears[i]])
  prevmeet[i] = sum(yday(data$Meeting[substr(data$Meeting,1,4)==prevyears[i]])<=diyear)
}
mmp.year = mean(mp.year)
m.prevmeet = mean(prevmeet)
prognosis = round(mmp.year*(mlyear/m.prevmeet)) - mlyear
data$Prognosis = c(rep(0,prevyearsl),prognosis,rep(NA,(nrow(data)-(prevyearsl+1))))
# adjusting column name for a graph in dashboard
colnames(data)[25] <- "Number of clients"
#Changing HowKnow names
data$HowKnow[data$HowKnow=="else"] = "Other"
data$HowKnow[is.na(data$HowKnow)] = "Unknown"

#creating columns with year and month
data$Year = substr(data$Meeting,1,4)
data$MonthsFilter = data$montotal[as.integer(substr(data$Meeting,6,7))]

#creating cumulative column of monthly frequencies
data$cummonthsfreq = cumsum(data$monthsfreq)

#creating column with all the months
ye1 = substr(first.month,1,4)
ye2 = substr(last.month,1,4)
yes = as.numeric(ye1):as.numeric(ye2)
yes = as.character(yes)
mos = c("-01","-02","-03","-04","-05","-06","-07","-08","-09","-10","-11","-12")
yesmos = matrix(NA,nrow=length(mos),ncol=length(yes))
mo1 = as.numeric(substr(first.month,6,7))
for(i in 1:length(yes)){
  for(j in 1:length(mos)){
    yesmos[j,i] = paste0(yes[i],mos[j])
    if(paste0(yes[i],mos[j])==last.month){
      break
    }
  }
}
yesmos = as.vector(na.omit(as.vector(yesmos)))
yesmos = yesmos[which(yesmos==first.month):length(yesmos)]

#adding the number of meetings beside that column
ym = matrix(NA,ncol=3,nrow = length(yesmos))
ym[,1] = yesmos
for(i in 1:length(data$months)){
  for(j in 1:length(ym[,1])){
  if(data$months[i]==ym[j,1]){
    ym[j,2] = data$months[i]
    ym[j,3] = data$monthsfreq[i]
  }else{}
  }
}
ym[is.na(ym)]=0

#adding the new columns
data$allmonths = c(ym[,1],rep(NA,nrow(data)-length(ym[,1])))
data$allmonthscumfreq = c(cumsum(as.numeric(ym[,3])),rep(NA,nrow(data)-length(ym[,3])))

#remove all NA's
data[is.na(data)] = ""

write.csv(data,file="Dataset_Dashboard.csv")
