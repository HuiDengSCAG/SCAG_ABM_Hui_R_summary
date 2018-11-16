#library("dplyr")
#library("xlsx")
#library(data.table)
#library(tidyverse)
#library(readxl)


#setwd("D:/Data1/ABM_Hui/")
#Model_path <- "D:/Data1/ABM_Hui/20R16BY_test_rD5/"
#Summary_path <- "D:/Data1/ABM_Hui/Output_Summary_Reports/"
#Target_path <- "I:/Staff/HuiDeng/R project/R_ABM_outputs/SCAG_Targets/"

#popOutPath <- paste0(Model_path, "Outputs/popsyn_warmstart/")
#ctrampOutPath <- paste0(Model_path, "Outputs/ctramp_warmstart/")

### Model Files #####

hMod <- fread(paste0(popOutPath,"hhs.csv"))
names(hMod) <- gsub('.{2}$', '', names(hMod)) # Remove last two characters from all column names


pMod <- fread(paste0(ctrampOutPath,"derivedPersonTypes.csv"))
names(pMod) <- gsub('.{2}$', '', names(pMod)) # Remove last two characters from all column names

WorkArrangment_Mod <- fread(paste0(ctrampOutPath,"workArrangement.csv"))
names(WorkArrangment_Mod) <- gsub('.{2}$', '', names(WorkArrangment_Mod)) # Remove last two characters from all column names

WorkArrangment_Mod <- left_join(WorkArrangment_Mod, hMod[, c("hhid", "HHINC")], by=c("hhid"="hhid"))

WorkArrangment_Mod <- mutate(WorkArrangment_Mod, Inc=cut(WorkArrangment_Mod$HHINC, c(0, 35000,75000, 150000, Inf),
                                       labels= c("1=Inc0-35","2=Inc35-75","3=Inc75-150", "4=Inc150+")))

WorkArrangment_Mod$workDuration <- factor(WorkArrangment_Mod$workDuration, labels=c("1=0-20","2=20-34", "3=35+"))
WorkArrangment_Mod$workPlaceType <- factor(WorkArrangment_Mod$workPlaceType, labels=c("1=Fixed","2=Home", "3=Variable"))
WorkArrangment_Mod$numJobs <- factor(WorkArrangment_Mod$numJobs, labels=c("1=Single","2=Multiple"))



WorkArrangment_Summary <- aggregate(pid~Inc+workDuration+workPlaceType+numJobs, data = WorkArrangment_Mod, FUN = length)

names(WorkArrangment_Summary)[names(WorkArrangment_Summary)=="pid"] <- "Number of persons" 

#write.csv(WorkArrangment_Summary, file = paste0(ctrampOutPath, "summaries/csv/WorkArrangment_Summary.csv"), row.names = F)

#==============load xlsx and define dataframe format========================
Target_WorkArrangment_Flex <- loadWorkbook(paste0(Target_path, 
                                                  "= VT01 03 20180713 Work Arrangement and Schedule Flexibility.xlsx"),
                                           password=NULL)
removeSheet(Target_WorkArrangment_Flex, sheetName="Data")
removeSheet(Target_WorkArrangment_Flex, sheetName="Question")
sheets <- getSheets(Target_WorkArrangment_Flex)
sheet_num_data <- createSheet(Target_WorkArrangment_Flex, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_WorkArrangment_Flex, sheetName = "pct_Data")

#functions

xlsx.AddTable <- function(dataframe, wb, sheet, startRow, startColumn, CellColor){
  # Styles for the data table row/column names
  TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) + Fill(foregroundColor=CellColor) +
    Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
           pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 
  TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) + Fill(foregroundColor=CellColor) +
    Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
           pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 
  CELL_STYLE <- CellStyle(wb)+ Font(wb,isBold=FALSE) + Fill(foregroundColor=CellColor) + 
    Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
           pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 
  TABLE_CELL_STYLE <- rep(list(CELL_STYLE), dim(dataframe)[2])
  names(TABLE_CELL_STYLE) <- seq(1, dim(dataframe)[2], by=1)
  
  #  browser()  
  addDataFrame(dataframe, sheet, startRow=startRow, startColumn=startColumn,
               colStyle = TABLE_CELL_STYLE,
               colnamesStyle = TABLE_COLNAMES_STYLE,
               rownamesStyle = TABLE_ROWNAMES_STYLE 
  )
}

xlsx.AddCell <- function(wb, sheet, rowIndex, colIndex, Cell_value, Cell_style){
  if(Cell_style=="Title_1"){
    Cell_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=16, color="blue", isBold=TRUE, underline=1)
  } else {
    if(Cell_style=="Title_2"){
      Cell_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=12, color="blue", isBold=TRUE)
    } else { # for tabel cells
      if(Cell_style=="T_Cell_1"){
        Cell_STYLE <- CellStyle(wb)+ Font(wb,isBold=TRUE) + Fill(foregroundColor="darkseagreen1") + 
          Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
                 pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 
      } else{
        Cell_STYLE <- CellStyle(wb)+ Font(wb,isBold=FALSE) + Fill(foregroundColor="darkseagreen1") + 
          Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
                 pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 
      }
    }
  }
  TABLE_CELL_STYLE <- rep(list(Cell_STYLE), 1)
  names(TABLE_CELL_STYLE) <- seq(1, 1, by=1)
  
  addDataFrame(as.data.frame(Cell_value), sheet = sheet,
               startRow=rowIndex, startColumn=colIndex,
               row.names = F, col.names = F,
               colStyle = TABLE_CELL_STYLE
               )
}

#==============WorkArrangment_Inc_by_Duration===============================
Temp <- reshape(aggregate(`Number of persons`~Inc+workDuration, data = WorkArrangment_Summary, FUN = sum), 
                idvar = "Inc", timevar = "workDuration", direction = "wide")
rownames(Temp) <- Temp$Inc
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- c("0-20", "20-35", "35+", "Total")
rownames(Temp) <- c("1=Inc0-35","2=Inc35-75","3=Inc75-150", "4=Inc150+", "All")

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=1, colIndex=1, "Work Arrangment and Schedule", "Title_1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=2, colIndex=1, "Work Duration by Household Income", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=3, colIndex=1, "Income\\Hours", "T_Cell_1")

#addDataFrame(as.data.frame("Income\\Hours"), sheet_num_data,startRow=3, startColumn=1,row.names = F, col.names = F)

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=1, colIndex=1, "Work Arrangment and Schedule", "Title_1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=2, colIndex=1, "Work Duration by Household Income", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=3, colIndex=1, "Income\\Hours", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\Hours"), sheet_pct_data,startRow=3, startColumn=1,row.names = F, col.names = F)

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.1 WK Arrangement`, 
              startRow=5, startColumn=7, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.1 WK Arrangement`, rowIndex=5, colIndex=7, "Income\\Hours", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\Hours"), sheet = sheets$`2.1.1 WK Arrangement`,startRow=5, startColumn=7,row.names = F, col.names = F)

#===========WorkArrangment_Inc_by_Duration==============#

#==============WorkArrangment_Inc_by_workPlaceType===============================
Temp <- reshape(aggregate(`Number of persons`~Inc+workPlaceType, data = WorkArrangment_Summary, FUN = sum), 
                idvar = "Inc", timevar = "workPlaceType", direction = "wide")
rownames(Temp) <- Temp$Inc
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- gsub("Number of persons.", "", names(Temp))
rownames(Temp) <- c("1=Inc0-35","2=Inc35-75","3=Inc75-150", "4=Inc150+", "All")

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=10, colIndex=1, "Work PlaceType by Household Income", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=11, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=11, colIndex=1, "Income\\WLOC", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\WLOC"), sheet_num_data,startRow=11, startColumn=1,row.names = F, col.names = F)

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=10, colIndex=1, "Work PlaceType by Household Income", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=11, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=11, colIndex=1, "Income\\WLOC", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\WLOC"), sheet_pct_data,startRow=11, startColumn=1,row.names = F, col.names = F)

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.1 WK Arrangement`, 
              startRow=16, startColumn=7, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.1 WK Arrangement`, rowIndex=16, colIndex=7, "Income\\WLOC", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\WLOC"), sheet = sheets$`2.1.1 WK Arrangement`,startRow=16, startColumn=7,row.names = F, col.names = F)

#===========WorkArrangment_Inc_by_workPlaceType==============#

#==============WorkArrangment_Inc_by_number of jobs===============================
Temp <- reshape(aggregate(`Number of persons`~Inc+numJobs, data = WorkArrangment_Summary, FUN = sum), 
                idvar = "Inc", timevar = "numJobs", direction = "wide")
rownames(Temp) <- Temp$Inc
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- gsub("Number of persons.", "", names(Temp))
rownames(Temp) <- c("1=Inc0-35","2=Inc35-75","3=Inc75-150", "4=Inc150+", "All")

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=18, colIndex=1, 
             Cell_value = "Number of Jobs by Household Income", Cell_style = "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=19, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=19, colIndex=1, "Income\\Jobs", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\Jobs"), sheet_num_data,startRow=19, startColumn=1,row.names = F, col.names = F)

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=18, colIndex=1, "Number of Jobs by Household Income", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=19, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=19, colIndex=1, "Income\\Jobs", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\Jobs"), sheet_pct_data,startRow=19, startColumn=1,row.names = F, col.names = F)

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.1 WK Arrangement`, 
              startRow=28, startColumn=7, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.1 WK Arrangement`, rowIndex=28, colIndex=7, "Income\\Jobs", "T_Cell_1")
#addDataFrame(as.data.frame("Income\\Jobs"), sheet = sheets$`2.1.1 WK Arrangement`,startRow=28, startColumn=7,row.names = F, col.names = F)

#===========WorkArrangment_Inc_by_number of jobs==============#




#==========work flexibility by worker Age ====================
WorkFlex_Mod <- fread(paste0(ctrampOutPath,"usualWorkScheduleFlexibility.csv"))
names(WorkFlex_Mod) <- gsub('.{2}$', '', names(WorkFlex_Mod)) # Remove last two characters from all column names

WorkFlex_Mod <- left_join(WorkFlex_Mod, pMod[, c("pid", "hhid", "AGE")], by=c("hhid"="hhid", "pid"="pid"))
WorkFlex_Mod <- mutate(WorkFlex_Mod, Age=cut(WorkFlex_Mod$AGE, c(-Inf, 15, 29, 44, 64, Inf), 
                                              labels=c("0-15", "16-29", "30-44", "45-64", "65+")))
WorkFlex_Mod$wdaysCatF <- factor(WorkFlex_Mod$workNumDays, labels=c("1-4 days","5 Days", "6+ days"))
WorkFlex_Mod$flexibilityF <- factor(WorkFlex_Mod$workFlexibility, labels=c("1=None","2=Some", "3=High"))
WorkFlex_Mod$compWeekF <- factor(WorkFlex_Mod$compressedWeek, labels=c("2=No","1=Yes"))

WorkFlex_Summary <- aggregate(pid~Age+wdaysCatF+flexibilityF+compWeekF, data = WorkFlex_Mod, FUN = length)

names(WorkFlex_Summary)[names(WorkFlex_Summary)=="pid"] <- "Number of persons" 

#write.csv(WorkFlex_Summary, file = paste0(ctrampOutPath, "summaries/csv/WorkFelxibility_Summary.csv"), row.names = F)



#==============WorkFelxibility_Age_by_working days===============================
Temp <- reshape(aggregate(`Number of persons`~Age+wdaysCatF, data = WorkFlex_Summary, FUN = sum), 
                idvar = "Age", timevar = "wdaysCatF", direction = "wide")
rownames(Temp) <- Temp$Age
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- gsub("Number of persons.", "", names(Temp))
rownames(Temp) <- c("16-29", "30-44", "45-64", "65+", "All")

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=26, colIndex=1, "Work Flexibility", "Title_1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=27, colIndex=1, "Working Days by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=28, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=28, colIndex=1, "Age\\Working Days", "T_Cell_1")

xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=26, colIndex=1, "Work Flexibility", "Title_1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=27, colIndex=1, "Working Days by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=28, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=28, colIndex=1, "Age\\Working Days", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.3 Work Scehdule Flex`, 
              startRow=5, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.3 Work Scehdule Flex`, rowIndex=5, colIndex=9, "Age\\Working Days", "T_Cell_1")
#===========WorkFelxibility_Age_by_working days==============#

#==============WorkFelxibility_Age_by_Flexibility Index===============================
Temp <- reshape(aggregate(`Number of persons`~Age+flexibilityF, data = WorkFlex_Summary, FUN = sum), 
                idvar = "Age", timevar = "flexibilityF", direction = "wide")
rownames(Temp) <- Temp$Age
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- gsub("Number of persons.", "", names(Temp))
rownames(Temp) <- c("16-29", "30-44", "45-64", "65+", "All")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=35, colIndex=1, "Work Flexibility by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=36, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=36, colIndex=1, "Age\\Flexibility", "T_Cell_1")


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=35, colIndex=1, "Work Flexibility by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=36, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=36, colIndex=1, "Age\\Flexibility", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.3 Work Scehdule Flex`, 
              startRow=14, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.3 Work Scehdule Flex`, rowIndex=14, colIndex=9, "Age\\Flexibility", "T_Cell_1")

#===========WorkFelxibility_Age_by_WorkFelxibility Index==============#


#==============WorkFelxibility_Age_by_Compressed Week===============================
Temp <- reshape(aggregate(`Number of persons`~Age+compWeekF, data = WorkFlex_Summary, FUN = sum), 
                idvar = "Age", timevar = "compWeekF", direction = "wide")
rownames(Temp) <- Temp$Age
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

names(Temp) <- gsub("Number of persons.", "", names(Temp))
rownames(Temp) <- c("16-29", "30-44", "45-64", "65+", "All")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=43, colIndex=1, "Work Flexibility by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=44, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=44, colIndex=1, "Age\\Flexibility", "T_Cell_1")


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=43, colIndex=1, "Work Flexibility by Age", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=44, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=44, colIndex=1, "Age\\Flexibility", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheets$`2.1.3 Work Scehdule Flex`, 
              startRow=24, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.3 Work Scehdule Flex`, rowIndex=24, colIndex=9, "Age\\Flexibility", "T_Cell_1")

#==============WorkFelxibility_Age_by_Compressed Week===============================

#============== Percent of Days X Work Flexibility X Compressed Week===============================
Temp1 <- reshape(aggregate(`Number of persons`~wdaysCatF+ flexibilityF, 
                          data = WorkFlex_Summary[which(WorkFlex_Summary$compWeekF=="1=Yes"), ], FUN = sum), 
                idvar = "wdaysCatF", timevar = "flexibilityF", direction = "wide")
names(Temp1) <- gsub("Number of persons.", "", names(Temp1))
names(Temp1)[-1] <- paste0(names(Temp1)[-1], "_Compr=Yes")
Temp2 <- reshape(aggregate(`Number of persons`~wdaysCatF+ flexibilityF, 
                           data = WorkFlex_Summary[which(WorkFlex_Summary$compWeekF=="2=No"), ], FUN = sum), 
                 idvar = "wdaysCatF", timevar = "flexibilityF", direction = "wide")
names(Temp2) <- gsub("Number of persons.", "", names(Temp2))
names(Temp2)[-1] <- paste0(names(Temp2)[-1], "_Compr=No")
Temp <- left_join(Temp1,Temp2,by="wdaysCatF")
rownames(Temp) <- Temp$wdaysCatF
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))

rownames(Temp) <- c("1-4 days","5 Days", "6+ days", "All")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=51, colIndex=1, "Percent of Days X Work Flexibility X Compressed Week", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_WorkArrangment_Flex, sheet = sheet_num_data, 
              startRow=52, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_num_data, rowIndex=52, colIndex=1, "Wday\\Flex", "T_Cell_1")


xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=51, colIndex=1, "Percent of Days X Work Flexibility X Compressed Week", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_WorkArrangment_Flex, sheet = sheet_pct_data, 
              startRow=52, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheet_pct_data, rowIndex=52, colIndex=1, "Wday\\Flex", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct[-4,-7]), Target_WorkArrangment_Flex, sheet = sheets$`2.1.3 Work Scehdule Flex`, 
              startRow=35, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_WorkArrangment_Flex, sheets$`2.1.3 Work Scehdule Flex`, rowIndex=35, colIndex=9, "Wday\\Flex", "T_Cell_1")

#============== Percent of Days X Work Flexibility X Compressed Week===============================
saveWorkbook(Target_WorkArrangment_Flex, file = paste0(Model_path, "Outputs/user/==VT01 03 20180713 Work Arrangement and Schedule Flexibility.xlsx"))

rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", 
                           "hMod", "pMod")])
