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

#Function for determining retirement eligibility
IsRetirementEligible <- function(Age, YOS){
  Check = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |      
                 (Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                 (Age + YOS) >= ReduceRetAge, TRUE, FALSE)
  return(Check)
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
SurvivalMale <- left_join(SurvivalMale,MaleMortality,by = 'Age') %>% group_by(Age) %>% 
  #MPValue2 is the cumulative product of the MP-2020 value for year 2036. We use it later so make life easy and calculate now
  mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)),
         Value = ifelse(Age == 120, 1,
                        ifelse(Age < NormalRetAgeII & Years == 2010, `Pub-2010 Employee Male Teachers`,
                               ifelse(Age >= NormalRetAgeII & Years == 2010, `Pub-2010 Non-disabled Male Teachers` * ScaleMultiple,
                                      ifelse(Age < NormalRetAgeII & Years > 2010, `Pub-2010 Employee Male Teachers` * MPValue2,
                                             ifelse(Age >= NormalRetAgeII & Years > 2010, `Pub-2010 Non-disabled Male Teachers` * ScaleMultiple * MPValue2, 0))))))
#filter out the necessary variables
SurvivalMale <- SurvivalMale %>% select('Age','Years','Value') %>% ungroup()

#Expand grid for ages 25-120 and years 2009 to 2019
SurvivalFemale <- expand_grid(25:120,2009:2129)
colnames(SurvivalFemale) <- c('Age','Years')
SurvivalFemale$Value <- 0
#Join these tables to make the calculations easier
SurvivalFemale <- left_join(SurvivalFemale,SurvivalRates,by = 'Age')
SurvivalFemale <- left_join(SurvivalFemale,FemaleMortality,by = 'Age') %>% group_by(Age) %>% 
  #MPValue2 is the cumulative product of the MP-2020 value for year 2036. We use it later so make life easy and calculate now
  mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)),
         Value = ifelse(Age == 120, 1,
                        ifelse(Age < NormalRetAgeII & Years == 2010, `Pub-2010 Employee Female Teachers`,
                               ifelse(Age >= NormalRetAgeII & Years == 2010, `Pub-2010 Non-disabled Female Teachers`,
                                      ifelse(Age < NormalRetAgeII & Years > 2010, `Pub-2010 Employee Female Teachers` * MPValue2,
                                             ifelse(Age >= NormalRetAgeII & Years > 2010, `Pub-2010 Non-disabled Female Teachers` * MPValue2, 0))))))
#filter out the necessary variables
SurvivalFemale <- SurvivalFemale %>% select('Age','Years','Value') %>% ungroup()

#Separation Rates
SeparationRates <- expand_grid(25:80,0:55)
colnames(SeparationRates) <- c('Age','YOS')
SeparationRates$SepProb <- SeparationRates$SepProbMale <- SeparationRates$SepProbFemale <- SeparationRates$Cond <- 0

TerminationRateAfter5 <- read_excel(FileName, sheet = 'Termination Rates after 5')
TerminationRateBefore5 <- read_excel(FileName, sheet = 'Termination Rates before 5')
RetirementRates <- read_excel(FileName, sheet = 'Retirement Rates')

SeparationRates <- left_join(SeparationRates,TerminationRateAfter5, by = 'Age')
SeparationRates <- left_join(SeparationRates,TerminationRateBefore5,by = 'YOS')
SeparationRates <- left_join(SeparationRates,RetirementRates,by = 'Age')
#Remove NAs
SeparationRates[is.na(SeparationRates)] <- 0

#If you're retirement eligible, use the retirement rates, then checks YOS < 5 and use the regular termination rates
SeparationRates <- SeparationRates %>% mutate(SepProbMale = ifelse(IsRetirementEligible(Age,YOS),`Retirement Rate`,ifelse(YOS < 5, `Termination before 5 Male`, `Termination after 5 Male`)),
                                              SepProbFemale = ifelse(IsRetirementEligible(Age,YOS),`Retirement Rate`,ifelse(YOS < 5, `Termination before 5 Female`, `Termination after 5 Female`)),
                                              SepProb = (pmax(SepProbMale,`Termination before 5 Male`) + pmax(SepProbFemale,`Termination before 5 Female`)) / 2,
                                              Cond = IsRetirementEligible(Age,YOS))
#Filter out unecessary values
SeparationRates <- SeparationRates %>% select(Age,YOS,SepProb)

GetNormalCost <- function(HiringAge, StartingSalary){
  #Filter out unecessary values
  SeparationRates <- SeparationRates %>% filter(Age - YOS == HiringAge)
  
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
  AdjustedMale <- SurvivalMale %>% filter(Age - (Years - YearStart) == HiringAge, Years >= YearStart)
  AdjustedFemale <- SurvivalFemale %>% filter(Age - (Years - YearStart) == HiringAge, Years >= YearStart)
  
  #Rename columns for merging and taking average
  colnames(AdjustedMale) <- c('Age','Years','AdjMale')
  colnames(AdjustedFemale) <- c('Age','Years','AdjFemale')
  AdjustedValues <- left_join(AdjustedMale,AdjustedFemale) %>% mutate(AdjValue = (AdjMale+AdjFemale)/2)
  
  #Survival Probability and Annuity Factor
  AnnFactorData <- AdjustedValues %>% select(Age,AdjValue) %>%
    mutate(Prob = cumprod(1 - lag(AdjValue, default = 0)),
           DiscProb = Prob/(1+ARR)^(Age - HiringAge),
           surv_DR_COLA = DiscProb * (1+COLA)^(Age-HiringAge),
           AnnuityFactor = rev(cumsum(rev(surv_DR_COLA)))/surv_DR_COLA)
  
  #Replacement Rates
  AFNormalRetAgeII <- AnnFactorData$AnnuityFactor[AnnFactorData$Age == NormalRetAgeII]
  SurvProbNormalRetAgeII <- AnnFactorData$Prob[AnnFactorData$Age  == NormalRetAgeII]
  RetentionRates_Inputs <- read_excel(FileName, sheet = 'Retention Rates')
  RetentionRates <- expand_grid(HiringAge:120,5:60)
  colnames(RetentionRates) <- c('Age','YOS')
  RetentionRates <- left_join(RetentionRates,RetentionRates_Inputs, by = 'Age') 
  RetentionRates <-  left_join(RetentionRates,AnnFactorData %>% select(Age,Prob,AnnuityFactor), by = 'Age') %>%
    mutate(AF = AnnuityFactor, SurvProb = Prob,
           #Replacement rate depending on the different age conditions
           RepRate = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |      
                              (Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                              (Age + YOS) >= ReduceRetAge, 1,
                            #Early Retirement rates
                            ifelse((Age < NormalRetAgeII & YOS >= NormalYOSII) |
                                   (Age < NormalRetAgeII & Age >= EarlyRetAge & YOS >= EarlyRetYOS),
                                   AFNormalRetAgeII / (1+ARR)^(NormalRetAgeII - Age)*SurvProbNormalRetAgeII / SurvProb / AF,
                                   ifelse((Age + YOS) >= ReduceRetAge, Factor, 0))))
  #Rename this column so we can join it to the benefit table later
  colnames(RetentionRates)[1] <- 'Retirement Age'
  RetentionRates <- RetentionRates %>% select(`Retirement Age`, YOS, RepRate)
  
  #Benefits, Annuity Factor and Present Value for ages 45-120
  BenefitsTable <- expand_grid(HiringAge:120,45:120)
  colnames(BenefitsTable) <- c('Age','Retirement Age')
  BenefitsTable <- left_join(BenefitsTable,SalaryData, by = "Age") %>% 
                   left_join(RetentionRates, by = c("Retirement Age", "YOS")) %>% 
                   left_join(AnnFactorData %>%  select(Age, Prob, AnnuityFactor), by = c("Retirement Age" = "Age")) %>%
                   #Prob and AF at retirement
                   rename(Prob_Ret = Prob, AF_Ret = AnnuityFactor) %>% 
                   #Rejoin the table to get the regular AF and Prob
                   left_join(AnnFactorData %>% select(Age, Prob), by = c("Age"))
  
  BenefitsTable <- BenefitsTable %>% 
    mutate(GradedMult = BenMult1*pmin(YOS,25) + BenMult2*pmax(pmax(YOS,25)-25,0),
           RepRateMult = case_when(GradMult == 0 ~ RepRate*BenMult2*YOS, TRUE ~ RepRate*GradedMult),
           AnnFactorAdj = (AF_Ret / ((1+ARR)^(`Retirement Age` - Age)))*(Prob_Ret / Prob),
           PensionBenefit = RepRateMult*FinalAvgSalary,
           PresentValue = ifelse(Age > `Retirement Age`,0,PensionBenefit*AnnFactorAdj))
  
  #The max benefit is done outside the table because it will be merged with Salary data
  OptimumBenefit <- BenefitsTable %>% group_by(Age) %>% summarise(MaxBenefit = max(PresentValue))
  SalaryData <- left_join(SalaryData,OptimumBenefit) 
  SalaryData <- left_join(SalaryData,SeparationRates,by = c('Age','YOS')) %>%
    mutate(PenWealth = ifelse(YOS < Vesting,(DBERBalance*ERVested)+DBEEBalance,pmax((DBERBalance*ERVested)+DBEEBalance,MaxBenefit)),
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
NormalCostFinal <- sum(SalaryHeadcountData$`Average Salary`*SalaryHeadcountData$`Headcount Total`*SalaryHeadcountData$NormalCost) /
                   sum(SalaryHeadcountData$`Average Salary`*SalaryHeadcountData$`Headcount Total`)
print(NormalCostFinal)

