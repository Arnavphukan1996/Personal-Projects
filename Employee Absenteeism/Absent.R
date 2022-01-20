rm(list=ls())
options(warn = -1)
X = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", 
      "C50", "dummy",
      "e1071", "Information",
      "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees',"mlr",
      "gridExtra","outliers","partitions","class")
lapply(X, require, character.only = TRUE)
rm(X)
library(boot)
library(rpart)
library(caret)
library(mltools)
library(carData)
library(car)
library(caTools)
library(ANOVA.TFNs)
library(ANOVAreplication)
library(car)
library(class)
library(scales)
library(psych)
library(gplots) 
library(readxl)
library(corrplot)
library(rpart)
library(DMwR)
library(rpart.plot)
library(tree)
library(randomForest)
library(caTools)

setwd("C:/Users/Arnav/Desktop/Projects")
getwd()
df = read_excel("Abs.xls",sheet = 1)
df = df

colnames(df)
str(df)
dim(df)
length(unique(df$Son))
length(unique(df$Pet))
length(unique(df$`dfeeism time in hours`))

#renaming the variables
names(df)[2]= "Reason.for.absence"
names(df)[3]= "Month.of.absence"
names(df)[4]= "Day.of.the.week"
names(df)[6]= "Transportation.expense"
names(df)[7]= "Distance.from.residence.to.work"
names(df)[8]= "Service.time" 
names(df)[10]= "Workload.average.perday"
names(df)[11]= "Hit.target"
names(df)[12]= "Disciplinary.failure"
names(df)[15]= "Social.drinker"
names(df)[16]= "Social.smoker" 
names(df)[20]= "Body.mass.index"
names(df)[21]= "dfeeism.time.in.hours"

 
# Checking missing values
sum(is.na(df))
missing_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
row.names(missing_val) = NULL
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
missing_val = missing_val[,c(2,1)]

# Missing value plots
ggplot(data = missing_val[1:8,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
ggtitle("Missing data percentage") + theme_bw()

ggplot(data = missing_val[9:15,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage") + theme_bw()

ggplot(data = missing_val[16:21,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage") + theme_bw()

#actual 31
# mean 
#median 25 
#df[2,20]
#df[2,20] = NA

# Freezing mean and median for missing value analysis

df$Reason.for.absence[is.na(df$Reason.for.absence)] = median(df$Reason.for.absence,na.rm = T)
df$Month.of.absence[is.na(df$Month.of.absence)] = median(df$Month.of.absence,na.rm = T)
df$Transportation.expense[is.na(df$Transportation.expense)] = median(df$Transportation.expense,na.rm = T)
df$Distance.from.residence.to.work[is.na(df$Distance.from.residence.to.work)] = median(df$Distance.from.residence.to.work,na.rm = T)
df$Service.time[is.na(df$Service.time)] = median(df$Service.time,na.rm = T)
df$Age[is.na(df$Age)] = median(df$Age,na.rm = T)
df$Workload.average.perday[is.na(df$Workload.average.perday)] = median(df$Workload.average.perday,na.rm = T)
df$Hit.target[is.na(df$Hit.target)] = median(df$Hit.target,na.rm = T)
df$Disciplinary.failure[is.na(df$Disciplinary.failure)] = median(df$Disciplinary.failure,na.rm = T)
df$Education[is.na(df$Education)] = median(df$Education,na.rm = T)
df$Son[is.na(df$Son)] = median(df$Son,na.rm = T)
df$Social.drinker[is.na(df$Social.drinker)] = median(df$Social.drinker,na.rm = T)
df$Social.smoker[is.na(df$Social.smoker)] = median(df$Social.smoker,na.rm = T)
df$Pet[is.na(df$Pet)] = median(df$Pet,na.rm = T)
df$Weight[is.na(df$Weight)] = median(df$Weight,na.rm = T)
df$Height[is.na(df$Height)] = median(df$Height,na.rm = T)
df$Body.mass.index[is.na(df$Body.mass.index)] = mean(df$Body.mass.index,na.rm = T)
df$dfeeism.time.in.hours[is.na(df$dfeeism.time.in.hours)] = median(df$dfeeism.time.in.hours,na.rm = T)

sum(is.na(df))

#changing numeric to categorical

df$ID = as.factor(df$ID)
df$Reason.for.absence = as.factor(df$Reason.for.absence)
df$Month.of.absence = as.factor(df$Month.of.absence)
df$Day.of.the.week = as.factor(df$Day.of.the.week)
df$Seasons= as.factor(df$Seasons)
df$Disciplinary.failure = as.factor(df$Disciplinary.failure)
df$Education = as.factor(df$Education)
df$Social.drinker= as.factor(df$Social.drinker)
df$Social.smoker = as.factor(df$Social.smoker)

str(df)

# outlier analysis 

numeric_index = sapply(df,is.numeric)
numeric_data = df[,numeric_index]
numeric_data= as.data.frame(numeric_data)
cnames = colnames(numeric_data)

# boxplot for numeric data
for (i in 1:length(cnames))
   {
     assign(paste0("gn",i), ggplot(aes_string(y = (cnames[i]), x = "dfeeism.time.in.hours"), data = subset(df))+ 
              stat_boxplot(geom = "errorbar", width = 0.5) +
             geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                           outlier.size=1, notch=FALSE) +
              theme(legend.position="bottom")+
              labs(y=cnames[i],x="dfeeism.time.in.hours")+
              ggtitle(paste("Box plot of dfeeism for",cnames[i])))
}

  # ## Plotting plots together
 gridExtra::grid.arrange(gn1,gn2,ncol=2)
 gridExtra::grid.arrange(gn3,gn4,ncol=2)
 gridExtra::grid.arrange(gn5,gn6,ncol=2)
 gridExtra::grid.arrange(gn7,gn8,ncol=2)
 gridExtra::grid.arrange(gn9,gn10,ncol=2)
 gridExtra::grid.arrange(gn11,gn12,ncol=2)

 
 
val = df$Transportation.expense[df$Transportation.expense %in% boxplot.stats(df$Transportation.expense)$out]
df$Transportation.expense[(df$Transportation.expense %in% val)] = NA
df$Transportation.expense[is.na(df$Transportation.expense)] = median(df$Transportation.expense, na.rm = T)
summary(df$Transportation.expense)

val1 = df$Distance.from.residence.to.work[df$Distance.from.residence.to.work
                                               %in% boxplot.stats(df$Distance.from.residence.to.work)$out]
df$Distance.from.residence.to.work[(df$Distance.from.residence.to.work %in% val1)] = NA
df$Distance.from.residence.to.work[is.na(df$Distance.from.residence.to.work)] = median(df$Distance.from.residence.to.work, na.rm = T)
summary(df$Distance.from.residence.to.work)
        
val2 = df$Service.time[df$Service.time %in% boxplot.stats(df$Service.time)$out]
df$Service.time[(df$Service.time %in% val2)] = NA
df$Service.time[is.na(df$Service.time)] = median(df$Service.time, na.rm = T)
summary(df$Service.time)

val3 = df$Age[df$Age %in% boxplot.stats(df$Age)$out]
df$Age[(df$Age %in% val3)] = NA
df$Age[is.na(df$Age)] = median(df$Age, na.rm = T)
summary(df$Age)

val4 = df$Workload.average.perday[df$Workload.average.perday %in% boxplot.stats(df$Workload.average.perday)$out]
df$Workload.average.perday[(df$Workload.average.perday %in% val4)] = NA
df$Workload.average.perday[is.na(df$Workload.average.perday)] = median(df$Workload.average.perday, na.rm = T)
summary(df$Workload.average.perday)

val5 = df$Hit.target[df$Hit.target %in% boxplot.stats(df$Hit.target)$out]
df$Hit.target[(df$Hit.target %in% val5)] = NA
df$Hit.target[is.na(df$Hit.target)] = median(df$Hit.target, na.rm = T)
summary(df$Hit.target)

val6 = df$Weight[df$Weight %in% boxplot.stats(df$Weight)$out]
df$Weight[(df$Weight %in% val6)] = NA
df$Weight[is.na(df$Weight)] = median(df$Weight, na.rm = T)
summary(df$Weight)

val7 = df$Height[df$Height %in% boxplot.stats(df$Height)$out]
df$Height[(df$Height %in% val7)] = NA
df$Height[is.na(df$Height)] = median(df$Height, na.rm = T)
summary(df$Height)

val8 = df$Body.mass.index[df$Body.mass.index %in% boxplot.stats(df$Body.mass.index)$out]
df$Body.mass.index[(df$Body.mass.index %in% val8)] = NA
df$Body.mass.index[is.na(df$Body.mass.index)] = median(df$Body.mass.index, na.rm = T)
summary(df$Body.mass.index)

val9 = df$Pet[df$Pet %in% boxplot.stats(df$Pet)$out]
df$Pet[(df$Pet %in% val9)] = NA
df$Pet[is.na(df$Pet)] = median(df$Pet, na.rm = T)
summary(df$Pet)

val10 = df$Son[df$Son %in% boxplot.stats(df$Son)$out]
df$Son[(df$Son %in% val10)] = NA
df$Son[is.na(df$Son)] = median(df$Son, na.rm = T)
summary(df$Son)

# correlation for continuous variables
corrgram(df[,numeric_index], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

# anova for categorical data
str(df)
result = aov(formula=dfeeism.time.in.hours~ID, data = df)
summary(result)
result1 = aov(formula=dfeeism.time.in.hours~Reason.for.absence, data = df)
summary(result1)
result2 = aov(formula=dfeeism.time.in.hours~Month.of.absence, data = df)
summary(result2)
result3 = aov(formula=dfeeism.time.in.hours~Day.of.the.week, data = df)
summary(result3)
result4 = aov(formula=dfeeism.time.in.hours~Seasons, data = df)
summary(result4)
result5 = aov(formula=dfeeism.time.in.hours~Disciplinary.failure, data = df)
summary(result5)
result6 = aov(formula=dfeeism.time.in.hours~Education, data = df)
summary(result6)
result7 = aov(formula=dfeeism.time.in.hours~Social.smoker, data = df)
summary(result7)
result8 = aov(formula=dfeeism.time.in.hours~Social.drinker, data = df)
summary(result8)

# important variables
df_final = subset(df,select=-c(ID,Seasons,Disciplinary.failure,Pet,Age,Education,Son,Social.smoker,Body.mass.index,Height,Hit.target,Social.drinker))
str(df_final)

#checking data distributions
qqnorm(df_final$Transportation.expense)
hist(df_final$Transportation.expense)
qqnorm(df_final$Distance.from.residence.to.work)
hist(df_final$Distance.from.residence.to.work)
qqnorm(df_final$Service.time)
hist(df_final$Service.time)
qqnorm(df_final$Workload.average.perday)
hist(df_final$Workload.average.perday)
qqnorm(df_final$Weight)
hist(df_final$Weight)


# Feature Scaling

cnames1 = c("Transportation.expense","Distance.from.residence.to.work","Service.time","Weight",
            "Workload.average.perday","dfeeism.time.in.hours")
for (i in cnames1) 
  {
  print(i)
  df_final[,i] = (df_final[,i]-min(df_final[,i]))/(max(df_final[,i]-min(df_final[,i])))
} 
range(df_final$Transportation.expense)
range(df_final$Distance.from.residence.to.work)
range(df_final$Service.time)
range(df_final$Workload.average.perday)
range(df_final$Weight)

# histogram after normalisation

qqnorm(df_final$Transportation.expense)
hist(df_final$Transportation.expense)
qqnorm(df_final$Distance.from.residence.to.work)
hist(df_final$Distance.from.residence.to.work)
qqnorm(df_final$Service.time)
hist(df_final$Service.time)
qqnorm(df_final$Workload.average.perday)
hist(df_final$Workload.average.perday)
qqnorm(df_final$Weight)
hist(df_final$Weight)

                                            ## ML Algorithm ##

df_final = as.data.frame(df_final)
train_ind = sample(1:nrow(df_final),0.8*nrow(df_final))                                                                                      
train = df_final[train_ind,]
test = df_final[-train_ind,]

 # Decision trees

set.seed(1234)
fit = rpart(dfeeism.time.in.hours~.,data = train, method = 'anova')
summary(fit)
prediction_dt = predict(fit,test[,-9])
actual = test[,9]
predicted1 = data.frame(prediction_dt)
rmse(preds = prediction_dt, actuals = actual, weights = 1, na.rm = FALSE)
# error rate = 11.620 #accuracy 88.38

# Random forest

set.seed(1234)
rf_mod = randomForest(dfeeism.time.in.hours~.,train,importance = TRUE,ntree = 100)
rf_pred = predict(rf_mod,test[,-9])
actual = test[,9]
predicted1 = data.frame(rf_pred)
rmse(preds = rf_pred, actuals = actual, weights = 1, na.rm = FALSE)
#error rate = 11.105 #accuracy 88.895

 #for ntree = 200

 set.seed(1234)
 rf_mod1 = randomForest(dfeeism.time.in.hours~.,train,importance = TRUE,ntree = 200)
 rf_pred1 = predict(rf_mod1,test[,-9])
 actual = test[,9]
 predicted2 = data.frame(rf_pred1)
 rmse(preds = rf_pred1, actuals = actual, weights = 1, na.rm = FALSE) 
 # error rate =11.093 #accuracy 88.907
 
 # for ntree = 300
 
 set.seed(1234)
 rf_mod2 = randomForest(dfeeism.time.in.hours~.,train,importance = TRUE,ntree = 300)
 rf_pred2 = predict(rf_mod2,test[,-9])
 actual = test[,9]
 predicted3 = data.frame(rf_pred2)
 rmse(preds = rf_pred2, actuals = actual, weights = 1, na.rm = FALSE)
 # error rate = 11.147  #accuracy 88.853
 
 # for ntree = 500
 
 set.seed(1234)
 rf_mod3 = randomForest(dfeeism.time.in.hours~.,train,importance = TRUE,ntree = 500)
 rf_pred3 = predict(rf_mod3,test[,-9])
 actual = test[,9]
 predicted4 = data.frame(rf_pred3)
 rmse(preds = rf_pred3, actuals = actual, weights = 1, na.rm = FALSE)
 # error rate = 11.172 #accuracy = 88.828
 
#LINEAR REGRESSION
#vif
 
mymodel = lm(dfeeism.time.in.hours~.,data= df_final)
vif(mymodel)

# LR Model

set.seed(1234)
df = as.data.frame(df)
droplevels(df_final$Reason.for.absence)
train_in1 = sample(1:nrow(df),0.8*nrow(df_final))
train1 = df_final[train_in1,]
test1 = df_final[-train_in1,]
lr_model = lm(dfeeism.time.in.hours~.,data = train1)
summary(lr_model)
predictions_lr = predict(lr_model,test1[,1:8])
predicted_lr = data.frame(predictions_lr)
actual1 = test1[,9]
rmse(preds = predictions_lr, actuals = actual, weights = 1, na.rm = FALSE)
 # error rate = 11.052  #accuracy 88.948

#---------------------------------------part-2#----------------------------------------------#

# creating subset
   
loss = subset(df,select = c(Month.of.absence,Service.time,dfeeism.time.in.hours,Workload.average.perday))  

# Workloss/month = (df time * workload)/service time    mathematical formula
loss["month.loss"]=with(loss,((loss[,"Workload.average.perday"]*loss[,"dfeeism.time.in.hours"])/loss[,"Service.time"]))
for (i in 9) {
  emp = loss[which(loss["Month.of.absence"]==i),]
  print(sum(emp$month.loss))
}

print(emp$month.loss) 





