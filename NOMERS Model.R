rm(list = ls())
library("readxl")
library(tidyverse)
library(zoo)
setwd(getwd())
FileName <- 'Model Inputs.xlsx'
YearStart <- 2021
#Assigning Variables
model_inputs <- read_excel(FileName, sheet = 'Main')
MeritIncreases <- model_inputs[,6]
for(i in 1:nrow(model_inputs)){
  if(!is.na(model_inputs[i,2])){
    assign(as.character(model_inputs[i,2]),as.double(model_inputs[i,3]))
  }
}

#NPV Function for later
NPV = function(rate, cashflows) {
  for(i in 1:length(cashflows)){
    if(i == 1){
      NPV <- cashflows[i]/((1+rate)^(i))
    } else {
      NPV <- NPV + cashflows[i]/((1+rate)^(i))
    }
  }
  
  return(NPV)
}

#These rates dont change so they're outside the function
#Mortality Rates
#2033 is the Age for the MP-2017 rates
MaleMortality <- read_excel(FileName, sheet = 'MP-2020_Male') %>% select(Age,'2036')
FemaleMortality <- read_excel(FileName, sheet = 'MP-2020_Female') %>% select(Age,'2036')
SurvivalRates <- read_excel(FileName, sheet = 'Mortality Rates')

#Expand grid for ages 25-120 and years 2009 to 2019
SurvivalMale <- expand_grid(25:120,2009:2129)
colnames(SurvivalMale) <- c('Age','Years')
SurvivalMale$Value <- 0
#Join these tables to make the calculations easier
SurvivalMale <- left_join(SurvivalMale,SurvivalRates,by = 'Age')
#MPValue2 is the cumulative product of the MP-2017 value for year 2033. We use it later so make life easy and calculate now
SurvivalMale <- left_join(SurvivalMale,MaleMortality,by = 'Age') %>% group_by(Age) %>% mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)))
SurvivalMale$Value <- ifelse(SurvivalMale$Age == 120, 1, 
                             ifelse((SurvivalMale$Age <= 57) & (SurvivalMale$Years == 2010),SurvivalMale$`Pub-2010 Employee Male Teachers`, 
                                    ifelse((SurvivalMale$Age > 57) & (SurvivalMale$Years == 2010),SurvivalMale$`Pub-2010 Non-disabled Male Teachers`*ScaleMultiple,
                                           ifelse((SurvivalMale$Age <= 57) & (SurvivalMale$Years > 2010),SurvivalMale$`Pub-2010 Employee Male Teachers`*SurvivalMale$MPValue2,
                                                  #This one is tricky. So this value is consistently 0.1. I basically made it dependent on the year rather than a cumulative sum since it is difficult to reset the sum for each age
                                                  ifelse((SurvivalMale$Age > 57) & (SurvivalMale$Years > 2010),SurvivalMale$`Pub-2010 Non-disabled Male Teachers`*ScaleMultiple*SurvivalMale$MPValue2,0)))))

#filter out the necessary variables
SurvivalMale <- SurvivalMale %>% select('Age','Years','Value')

#Expand grid for ages 25-120 and years 2009 to 2019
SurvivalFemale <- expand_grid(25:120,2009:2129)
colnames(SurvivalFemale) <- c('Age','Years')
SurvivalFemale$Value <- 0
#Join these tables to make the calculations easier
SurvivalFemale <- left_join(SurvivalFemale,SurvivalRates,by = 'Age')
#MPValue2 is the cumulative product of the MP-2017 value for year 2033. We use it later so make life easy and calculate now
SurvivalFemale <- left_join(SurvivalFemale,FemaleMortality,by = 'Age') %>% group_by(Age) %>% mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)))
SurvivalFemale$Value <- ifelse(SurvivalFemale$Age == 120, 1, 
                               ifelse((SurvivalFemale$Age < NormalRetAgeII) & (SurvivalFemale$Years == 2010),SurvivalFemale$`Pub-2010 Employee Female Teachers`, 
                                      ifelse((SurvivalFemale$Age >= NormalRetAgeII) & (SurvivalFemale$Years == 2010),SurvivalFemale$`Pub-2010 Non-disabled Female Teachers`,
                                             ifelse((SurvivalFemale$Age < NormalRetAgeII) & (SurvivalFemale$Years > 2010),SurvivalFemale$`Pub-2010 Employee Female Teachers`*SurvivalFemale$MPValue2,
                                                    #This one is tricky. So this value is consistently 0.1. I basically made it dependent on the year rather than a cumulative sum since it is difficult to reset the sum for each age
                                                    ifelse((SurvivalFemale$Age >= NormalRetAgeII) & (SurvivalFemale$Years > 2010),SurvivalFemale$`Pub-2010 Non-disabled Female Teachers`*SurvivalFemale$MPValue2,0)))))

#filter out the necessary variables
SurvivalFemale <- SurvivalFemale %>% select('Age','Years','Value')

GetNormalCost <- function(HiringAge, StartingSalary){
  #Change the sequence for the Age and YOS depending on the hiring age
  Age <- seq(HiringAge,120)
  YOS <- seq(0,(95-(HiringAge-25)))
  #Merit Increases need to have the same length as Age when the hiring age changes
  MeritIncreases <- MeritIncreases[1:length(Age),]
  TotalSalaryGrowth <- as.matrix(MeritIncreases) + salary_growth
  
  #Salary increases and other
  SalaryData <- tibble(Age,YOS) %>%
    mutate(Salary = StartingSalary*cumprod(1+lag(TotalSalaryGrowth,default = 0)),
           IRSSalaryCap = pmin(Salary,IRSCompLimit),
           FinalAvgSalary = ifelse(YOS >= Vesting, rollmean(lag(Salary), k = FinAvgSalaryYears, fill = 0, align = "right"), 0),
           EEContrib = EE_Contrib*Salary, ERContrib = ER_Contrib*Salary,
           ERVested = pmin(pmax(0+EnhER5*(YOS>=5)+EnhER610*(YOS-5)),1))
  
  #Because Credit Interest != inflation, you cant use NPV formulae for DB Balance
  for(i in 1:nrow(SalaryData)){
    if(SalaryData$YOS[i] == 0){
      SalaryData$DBEEBalance[i] <- 0
      SalaryData$DBERBalance[i] <- 0
      SalaryData$CumulativeWage[i] <- 0
    } else {
      SalaryData$DBEEBalance[i] <- SalaryData$DBEEBalance[i-1]*(1+Interest) + SalaryData$EEContrib[i-1]
      SalaryData$DBERBalance[i] <- SalaryData$DBERBalance[i-1]*(1+Interest) + SalaryData$ERContrib[i-1]
      SalaryData$CumulativeWage[i] <- SalaryData$CumulativeWage[i-1]*(1+ARR) + SalaryData$Salary[i-1]
    }
  }
  
  #Adjusted Mortality Rates
  #The adjusted values follow the same difference between Year and Hiring Age. So if you start in 2020 and are 25, then 26 at 2021, etc.
  #The difference is always 1995. After this remove all unneccessary values
  AdjustedMale <- SurvivalMale %>% mutate(Value = ifelse((Years - Age) != (YearStart - HiringAge), 0, Value)) %>% filter(Value > 0, Years >= YearStart)
  AdjustedFemale <- SurvivalFemale %>% mutate(Value = ifelse((Years - Age) != (YearStart - HiringAge), 0, Value)) %>% filter(Value > 0, Years >= YearStart)
  
  #Rename columns for merging and taking average
  colnames(AdjustedMale) <- c('Age','Years','AdjMale')
  colnames(AdjustedFemale) <- c('Age','Years','AdjFemale')
  AdjustedValues <- left_join(AdjustedMale,AdjustedFemale) %>% mutate(AdjValue = (AdjMale+AdjFemale)/2)
  
  #Survival Probability and Annuity Factor
  AnnFactorData <- AdjustedValues %>% select(Age,AdjValue)
  AnnFactorData$Prob <- cumprod(1 - lag(AnnFactorData$AdjValue, default = 0))
  AnnFactorData$DiscProb <- AnnFactorData$Prob/(1+ARR)^(AnnFactorData$Age - HiringAge)
  AnnFactorData$surv_DR_COLA <- AnnFactorData$DiscProb*(1+COLA)^(AnnFactorData$Age - HiringAge)
  AnnFactorData$AnnuityFactor <- rev(cumsum(rev(AnnFactorData$surv_DR_COLA)))/AnnFactorData$surv_DR_COLA
  
  #Replacement Rates
  AFNormalRetAgeII <- AnnFactorData$AnnuityFactor[AnnFactorData$Age == NormalRetAgeII]
  SurvProbNormalRetAgeII <- AnnFactorData$Prob[AnnFactorData$Age  == NormalRetAgeII]
  ReplacementRates_Inputs <- read_excel(FileName, sheet = 'Replacement Rates')
  ReplacementRates <- expand_grid(HiringAge:120,5:60)
  colnames(ReplacementRates) <- c('Age','YOS')
  ReplacementRates <- left_join(ReplacementRates,ReplacementRates_Inputs, by = 'Age') 
  ReplacementRates <-  left_join(ReplacementRates,AnnFactorData %>% select(Age,Prob,AnnuityFactor), by = 'Age') %>%
    mutate(AF = AnnuityFactor, SurvProb = Prob,
           #Replacement rate depending on the different age conditions
           RepRate = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |      
                              (Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                              (Age + YOS) >= ReduceRetAge, 1,
                            ifelse(Age < NormalRetAgeII & YOS >= NormalYOSII,
                                   AFNormalRetAgeII / (1+ARR)^(NormalRetAgeII - Age)*SurvProbNormalRetAgeII / SurvProb / AF,
                                   ifelse((Age + YOS) >= ReduceRetAge, Factor, 0))))
  #Rename this column so we can join it to the benefit table later
  colnames(ReplacementRates)[1] <- 'Retirement Age'
  
  #Benefits, Annuity Factor and Present Value for ages 45-120
  BenefitsTable <- expand_grid(HiringAge:120,45:120)
  colnames(BenefitsTable) <- c('Age','Retirement Age')
  BenefitsTable <- left_join(BenefitsTable,SalaryData)
  BenefitsTable <- left_join(BenefitsTable,ReplacementRates) %>%
    mutate(GradedMult = BenMult1*pmin(YOS,25) + BenMult2*pmax(pmax(YOS,25)-25,0),
           RepRateMult = ifelse(YOS < Vesting,0, 
                                ifelse((GradMult == 0) | (YOS == 0),BenMult2*RepRate*YOS, RepRate*GradedMult)),
           AF = ifelse(Age >= `Retirement Age`,AnnuityFactor[Age == `Retirement Age`],AnnuityFactor),
           SurvProb = ifelse(Age >= `Retirement Age`,Prob[Age == `Retirement Age`],Prob),
           AnnFactorAdj = ifelse(YOS >= Vesting , (AF / ((1+ARR)^(`Retirement Age` - Age)))*(SurvProb / Prob),0),
           PensionBenefit = pmin(pmax(ifelse(YOS >= Vesting, RepRateMult*FinalAvgSalary, 0),MinRetBen),MaxRetBen),
           PresentValue = ifelse(Age > `Retirement Age`,0,PensionBenefit*AnnFactorAdj))
  
  #The max benefit is done outside the table because it will be merged with Salary data
  OptimumBenefit <- BenefitsTable %>% group_by(Age) %>% summarise(MaxBenefit = max(PresentValue))
  RetentionRates <- read_excel(FileName, sheet = 'Retention Rates')
  SalaryData <- left_join(SalaryData,OptimumBenefit) 
  SalaryData <- left_join(SalaryData,RetentionRates) %>%
    mutate(PenWealth = ifelse(YOS<Vesting,(DBERBalance*ERVested)+DBEEBalance,pmax((DBERBalance*ERVested)+DBEEBalance,MaxBenefit)),
           PVPenWealth = PenWealth/(1+ARR)^(Age-HiringAge),
           #Fiter out the NAs because when you change the Hiring age, there are NAs in separation probability,
           #or in Pension wealth
           PVCumWage = CumulativeWage/(1+ARR)^(Age-HiringAge)) %>% filter(!is.na(PVPenWealth),!is.na(SepProb))
  
  #Calc and return Normal Cost
  ProbCumWage <- sum(SalaryData$SepProb*SalaryData$PVCumWage)
  #Do the If else to avoid diving by 0.
  NormalCost <- ifelse(ProbCumWage > 0, sum(SalaryData$SepProb*SalaryData$PVPenWealth) / ProbCumWage, 0)
  return(NormalCost)
}

SalaryHeadcountData <- read_excel(FileName, sheet = 'Salary and Headcount')
#This part requires a for loop since GetNormalCost cant be vectorized.
for(i in 1:nrow(SalaryHeadcountData)){
  SalaryHeadcountData$NormalCost[i] <- GetNormalCost(SalaryHeadcountData$`Hiring Age`[i], SalaryHeadcountData$`Starting Salary`[i])
}
#Calc the weighted average Normal Cost
NormalCostFinal <- sum(SalaryHeadcountData$NormalCost*SalaryHeadcountData$`Average Salary`*SalaryHeadcountData$Headcount) /
  sum(SalaryHeadcountData$`Average Salary`*SalaryHeadcountData$Headcount)
print(NormalCostFinal)
#
