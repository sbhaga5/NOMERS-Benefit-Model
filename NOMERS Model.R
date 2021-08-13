rm(list = ls())
library("readxl")
library(tidyverse)
library(zoo)
setwd(getwd())

FileName <- 'Model Inputs.xlsx'
YearStart <- 2021
Age <- 20:120
YOS <- 0:100
RetirementAge <- 20:120
Years <- 2010:2121    #(why 2121? Because 120 - 20 + 2021 = 2121)
#Assigning individual  Variables
model_inputs <- read_excel(FileName, sheet = 'Main')

for(i in 1:nrow(model_inputs)){
  if(!is.na(model_inputs[i,2])){
    assign(as.character(model_inputs[i,2]),as.double(model_inputs[i,3]))
  }
}

#Import key data tables
SurvivalRates <- read_excel(FileName, sheet = 'Mortality Rates')
MaleMP <- read_excel(FileName, sheet = 'MP-2020_Male')
FemaleMP <- read_excel(FileName, sheet = 'MP-2020_Female')
SalaryGrowth <- read_excel(FileName, sheet = "Salary Growth")
SalaryEntry <- read_excel(FileName, sheet = "Salary and Headcount") %>% 
  select(entry_age, start_sal, count_start)
TerminationRateAfter5 <- read_excel(FileName, sheet = 'Termination Rates after 5')
TerminationRateBefore5 <- read_excel(FileName, sheet = 'Termination Rates before 5')
RetirementRates <- read_excel(FileName, sheet = 'Retirement Rates')


#Function for determining retirement eligibility (including normal retirement, unreduced early retirement, and reduced early retirement)
IsRetirementEligible <- function(Age, YOS){
  Check = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |
                   YOS >= 30 |
                   (Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                   (Age + YOS) >= ReduceRetAge |
                   (Age >= 60 & YOS >= 10), TRUE, FALSE)
  return(Check)
}

#These rates dont change so they're outside the function
#Transform base mortality rates and mortality improvement rates
MaleMP <- MaleMP %>% 
  pivot_longer(-Age, names_to = "Years", values_to = "MP_male") %>% 
  mutate(Years = as.numeric(Years))

MaleMP_ultimate <- MaleMP %>% 
  filter(Years == max(Years)) %>% 
  rename(MP_ultimate_male = MP_male) %>% 
  select(-Years)

FemaleMP <- FemaleMP %>% 
  pivot_longer(-Age, names_to = "Years", values_to = "MP_female") %>% 
  mutate(Years = as.numeric(Years))

FemaleMP_ultimate <- FemaleMP %>% 
  filter(Years == max(Years)) %>% 
  rename(MP_ultimate_female = MP_female) %>% 
  select(-Years)

##Mortality calculations
#Expand grid for ages 20-120 and years 2010 to 2121 (why 2121? Because 120 - 20 + 2021 = 2121)
MortalityTable <- expand_grid(Age, Years)

#Join base mortality table with mortality improvement table and calculate the final mortality rates
MortalityTable <- MortalityTable %>% 
  left_join(SurvivalRates, by = "Age") %>% 
  left_join(MaleMP, by = c("Age", "Years")) %>% 
  left_join(FemaleMP, by = c("Age", "Years")) %>% 
  left_join(MaleMP_ultimate, by = "Age") %>% 
  left_join(FemaleMP_ultimate, by = "Age") %>% 
  mutate(MaleMP_final = ifelse(Years <= max(MaleMP$Years), MP_male, MP_ultimate_male),
         FemaleMP_final = ifelse(Years <= max(FemaleMP$Years), MP_female, MP_ultimate_female),
         entry_age = Age - (Years - YearStart),
         YOS = Age - entry_age) %>% 
  group_by(Age) %>% 
  #MPcumprod is the cumulative product of (1 - MP rates), starting from 2011. We use it later so make life easy and calculate now
  mutate(MPcumprod_male = cumprod(1 - MaleMP_final)/(1 - MaleMP_final[Years == 2010]),
         MPcumprod_female = cumprod(1 - FemaleMP_final)/(1 - FemaleMP_final[Years == 2010]),
         mort_male = ifelse(IsRetirementEligible(Age, YOS) == F, PubG_2010_employee_male * MPcumprod_male,
                            PubG_2010_healthy_retiree_male * MPcumprod_male),
         mort_female = ifelse(IsRetirementEligible(Age, YOS) == F, PubG_2010_employee_female * MPcumprod_female,
                              PubG_2010_healthy_retiree_female * MPcumprod_female),
         mort = (mort_male + mort_female)/2) %>% 
  filter(Years >= 2021,
         entry_age >= 20) %>% 
  ungroup()

#filter out the necessary variables
MortalityTable <- MortalityTable %>% select(Age, Years, entry_age, mort) %>% 
  arrange(entry_age) 

#Separation Rates
SeparationRates <- expand_grid(Age, YOS) %>% 
  mutate(entry_age = Age - YOS) %>% 
  filter(entry_age %in% SalaryEntry$entry_age) %>% 
  arrange(entry_age, Age) %>% 
  left_join(TerminationRateAfter5, by = "Age") %>%
  left_join(TerminationRateBefore5, by = "YOS") %>% 
  left_join(RetirementRates, by = "Age") 

#If you're retirement eligible, use the retirement rates, then checks YOS < 5 and use the regular termination rates
SeparationRates <- SeparationRates %>% 
  mutate(retirement_cond = IsRetirementEligible(Age,YOS),
         SepRateMale = ifelse(retirement_cond == T, retirement_rate,
                              ifelse(YOS < 5, term_before_5_male, term_after_5_male)),
         SepRateFemale = ifelse(retirement_cond == T,retirement_rate,
                                ifelse(YOS < 5, term_before_5_female, term_after_5_female)),
         SepRate = (SepRateMale + SepRateFemale) / 2) %>% 
  group_by(entry_age) %>% 
  mutate(RemainingProb = cumprod(1 - lag(SepRate, default = 0)),
         SepProb = lag(RemainingProb, default = 1) - RemainingProb) %>% 
  ungroup()

#Filter out unecessary values
SeparationRates <- SeparationRates %>% select(Age,YOS,SepProb)

#Create a long-form table of Age and YOS and merge with salary data
SalaryData <- expand_grid(Age, YOS) %>% 
  mutate(entry_age = Age - YOS) %>%    #Add entry age
  filter(entry_age %in% SalaryEntry$entry_age) %>% 
  arrange(entry_age) %>% 
  left_join(SalaryEntry, by = "entry_age") %>% 
  left_join(SalaryGrowth, by = "Age")

#Custom function to calculate cumulative future values
cumFV <- function(interest, cashflow){
  cumvalue <- double(length = length(cashflow))
  for (i in 2:length(cumvalue)) {
    cumvalue[i] <- cumvalue[i - 1]*(1 + interest) + cashflow[i - 1]
  }
  return(cumvalue)
}

#Calculate FAS and cumulative EE contributions
SalaryData <- SalaryData %>% 
  group_by(entry_age) %>% 
  mutate(Salary = start_sal*cumprod(1+lag(salary_increase,default = 0)),
         # IRSSalaryCap = pmin(Salary,IRSCompLimit),
         FinalAvgSalary = rollmean(lag(Salary), k = FinAvgSalaryYears, fill = NA, align = "right"),
         EEContrib = EE_Contrib*Salary,
         DBEEBalance = cumFV(Interest, EEContrib),
         CumulativeWage = cumFV(ARR, Salary)) %>% 
  ungroup()

#Survival Probability and Annuity Factor
AnnFactorData <- MortalityTable %>% 
  select(Age, entry_age, mort) %>%
  group_by(entry_age) %>% 
  mutate(surv = cumprod(1 - lag(mort, default = 0)),
         surv_DR = surv/(1+ARR)^(Age - entry_age),
         surv_DR_COLA = surv_DR * (1+COLA)^(Age - entry_age),
         AnnuityFactor = rev(cumsum(rev(surv_DR_COLA)))/surv_DR_COLA) %>% 
  ungroup()

#Reduced Factor
  #Unreduced retirement benefit (for those hired after 2018) when:
  #Age 65 and 5 YOS
  #30 YOS
  #Age 62 with 20 YOS
  #Age + YOS >= 80
  #Reduced retirement benefit when:
  #Age 60 and 10 YOS: reduced by 3% per year prior to age 62
ReducedFactor <- expand_grid(20:120,0:100)
colnames(ReducedFactor) <- c('RetirementAge','YOS')
ReducedFactor <- ReducedFactor %>% 
  mutate(RF = ifelse(RetirementAge >= 65 & YOS >=5 |
                       YOS >= 30 |
                       RetirementAge >= 62 & YOS >= 20 |
                       RetirementAge + YOS >= 80, 1,
                     ifelse(RetirementAge >= 60 & YOS >= 10, pmin(1 - 0.03*(62 - RetirementAge),1), 0)))

# ReducedFactor_test <- ReducedFactor %>% pivot_wider(names_from = YOS, values_from = RF)

#Benefits, Annuity Factor and Present Value 
BenefitsTable <- expand_grid(Age, YOS, RetirementAge) %>% 
  mutate(entry_age = Age - YOS) %>% 
  filter(entry_age %in% SalaryEntry$entry_age) %>% 
  arrange(entry_age, Age, RetirementAge) %>% 
  left_join(SalaryData, by = c("Age", "YOS", "entry_age")) %>% 
  left_join(ReducedFactor, by = c("RetirementAge", "YOS")) %>% 
  left_join(AnnFactorData %>%  select(Age, entry_age, surv_DR, AnnuityFactor), by = c("RetirementAge" = "Age", "entry_age")) %>%
  #Rename surv_DR and AF to make clear that these variables are at retirement
  rename(surv_DR_ret = surv_DR, AF_Ret = AnnuityFactor) %>% 
  #Rejoin the table to get the surv_DR for the termination age
  left_join(AnnFactorData %>% select(Age, entry_age, surv_DR), by = c("Age", "entry_age")) %>% 
  mutate(GradedMult = BenMult1*pmin(YOS,25) + BenMult2*pmax(pmax(YOS,25)-25,0),
         ReducedFactMult = case_when(GradMult == 0 ~ RF*BenMult1*YOS, 
                                     TRUE ~ RF*GradedMult),
         AnnFactorAdj = AF_Ret * surv_DR_ret / surv_DR,
         MinBenefit = ifelse(YOS >= 10, 3600, 0),      #Minimum retirement benefit of $3,600 per year for any member with at least 10 YOS
         PensionBenefit = pmax(ReducedFactMult * FinalAvgSalary, MinBenefit),
         PresentValue = ifelse(Age > RetirementAge, 0, PensionBenefit*AnnFactorAdj))

#The max benefit is done outside the table because it will be merged with Salary data
OptimumBenefit <- BenefitsTable %>% 
  group_by(entry_age, Age) %>% 
  summarise(MaxBenefit = max(PresentValue)) %>%
  mutate(MaxBenefit = ifelse(is.na(MaxBenefit), 0, MaxBenefit)) %>% 
  ungroup()

#Combine optimal benefit with employee balance and calculate the PV of future benefits and salaries 
SalaryData <- SalaryData %>% 
  left_join(OptimumBenefit, by = c("Age", "entry_age")) %>% 
  left_join(SeparationRates, by = c("Age", "YOS")) %>%
  mutate(PenWealth = pmax(DBEEBalance,MaxBenefit),
         PVPenWealth = PenWealth/(1 + ARR)^YOS * SepProb,
         PVCumWage = CumulativeWage/(1 + ARR)^YOS * SepProb) 


#Calculate normal cost rate for each entry age
NormalCost <- SalaryData %>% 
  group_by(entry_age) %>% 
  summarise(normal_cost = sum(PVPenWealth)/sum(PVCumWage)) %>% 
  ungroup()

#Calculate the aggregate normal cost
NC_aggregate <- sum(NormalCost$normal_cost * SalaryEntry$start_sal * SalaryEntry$count_start)/
  sum(SalaryEntry$start_sal * SalaryEntry$count_start)





