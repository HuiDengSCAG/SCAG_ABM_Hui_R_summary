library("dplyr")
library("xlsx")
library(data.table)
library(tidyverse)
library(readxl)


setwd("D:/Data1/ABM_Hui/")
Model_path <- "D:/Data1/ABM_Hui/20R16BY_test_rE_1pct/"
Summary_path <- "D:/Data1/ABM_Hui/Output_Summary_Reports/"
Target_path <- "I:/Staff/HuiDeng/R project/R_ABM_outputs/SCAG_Targets/"

popOutPath <- paste0(Model_path, "Outputs/popsyn/")
ctrampOutPath <- paste0(Model_path, "Outputs/ctramp/")

options(java.parameters = "-Xmx220g")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_workArrangement.R")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_DLCarownership.R")


source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_TripPurpMode.R")


source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_PrimaryDest.R")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_CoordinatedDailyActivityPattern.R")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_MActivity_TSkeleton.R")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_mandatoryPrelimTod.R")

source("I:/Staff/HuiDeng/R project/R_ABM_outputs/r_code/R_SchoolEscort.R")