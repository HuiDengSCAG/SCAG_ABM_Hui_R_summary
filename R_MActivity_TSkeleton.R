

if(file.exists(paste0(ctrampOutPath,"/workerTourBreak_cleaned.csv"))){
  WorkerTourBreak <- data.frame(fread(paste0(ctrampOutPath,"/workerTourBreak_cleaned.csv")))
} else{
  WorkerTourBreak <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/workerTourBreak.csv"))
}

PType_map <-  data.frame(personType=c(1:8),
                         PBptype= c("P1FTW","P2PTW", "P3Col", "P4NWadult", "P5NWold", 
                                    "P6_16_18", "P7_6_15", "P8_0_5"),stringsAsFactors = F)

WorkerTourBreak <- WorkerTourBreak %>%
  left_join(pMod[,c("pid", "hhid", "personType")], by=c("pid"="pid", "hhid"="hhid")) %>%
  filter(personType %in% c(1,2)) %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  mutate(MandPattern=paste0(workerTourBreakTour1Types,
                            ifelse(workerTourBreakTour2Types=='','',paste0('-',workerTourBreakTour2Types))  )) %>%
  mutate(MandPattern= gsub('-',' ',gsub('[0-9]','',MandPattern))) %>%
  mutate(Num_W=str_count(MandPattern, "W")) %>%
  mutate(Num_B=str_count(MandPattern, "B")) %>%
  mutate(Num_S=str_count(MandPattern, "S")) %>%
  mutate(Num_WBS = Num_W+Num_B+Num_S)

PType_count <- pMod %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  group_by(PBptype) %>%
  summarise(NUM_ptype=n())

if(file.exists(paste0(ctrampOutPath,"/studentActivityFrequencyOrder_cleaned.csv"))){
  StudentActivityFOrder <- data.frame(fread(paste0(ctrampOutPath,"/studentActivityFrequencyOrder_cleaned.csv")))
} else{
  StudentActivityFOrder <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/studentActivityFrequencyOrder.csv"))
}

StudentActivityFOrder <- StudentActivityFOrder %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  mutate(MandPattern=paste0(studentActivityTour1Types, 
                            ifelse(studentActivityTour2Types=="", "", paste0("-", studentActivityTour2Types)))) %>%
  mutate(Num_S=str_count(MandPattern, "S")) %>%
  mutate(Num_W=str_count(MandPattern, "W")) %>%
  mutate(Num_SW=Num_S+Num_W)



#==============load xlsx and define dataframe format========================
Target_MA_TS <- loadWorkbook(paste0(Target_path, 
                                   "= VT09 20180502 Mandatory activity frequency and tour skeletons (5.2.2).xlsx"),
                            password=NULL)

sheet_num_data <- createSheet(Target_MA_TS, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_MA_TS, sheetName = "pct_Data")
sheets <- getSheets(Target_MA_TS)
removeSheet(Target_MA_TS, sheetName=sheets[[2]])
removeSheet(Target_MA_TS, sheetName=sheets[[3]])
removeSheet(Target_MA_TS, sheetName=sheets[[5]])
sheets <- getSheets(Target_MA_TS)

#=============summary for Worker work Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_W, data = WorkerTourBreak, FUN = length),
                idvar = "PBptype", timevar = "Num_W", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -ncol(Temp_pct)]


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=1, colIndex=1, "Mandatory Activity Frequency and Tour Skeletons Summary", "Title_1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=2, colIndex=1, "Worker Work Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=3, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=1, colIndex=1, "Mandatory Activity Frequency and Tour Skeletons Summary", "Title_1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=2, colIndex=1, "Worker Work Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=3, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=10, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=10, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Worker work Mandatory Activity==============================#

#=============summary for Worker Business Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_B, data = WorkerTourBreak, FUN = length),
                idvar = "PBptype", timevar = "Num_B", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -ncol(Temp_pct)]


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=7, colIndex=1, "Worker Business Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=8, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=8, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=7, colIndex=1, "Worker Business Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=8, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=8, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=16, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=16, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Worker Business Mandatory Activity==============================#


#=============summary for Worker Work Business Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_WBS, data = WorkerTourBreak, FUN = length),
                idvar = "PBptype", timevar = "Num_WBS", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -ncol(Temp_pct)]
Temp0 <- cbind(Temp0, Avg_MActivity=c(mean(WorkerTourBreak[which(WorkerTourBreak$PBptype=="P1FTW"),]$Num_WBS), 
                        (mean(WorkerTourBreak[which(WorkerTourBreak$PBptype=="P2PTW"),]$Num_WBS))))


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=12, colIndex=1, "Worker Mandatory Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=13, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=13, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=12, colIndex=1, "Worker Mandatory Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=13, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=13, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=22, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=22, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Worker work Business Mandatory Activity==============================#

#=============summary for Worker School Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_S, data = WorkerTourBreak, FUN = length),
                idvar = "PBptype", timevar = "Num_S", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -ncol(Temp_pct)]


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=17, colIndex=1, "Worker School Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=18, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=18, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=17, colIndex=1, "Worker School Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=18, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=18, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=28, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=28, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Worker School Mandatory Activity==============================#

#=============summary for Student School Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_S, data = StudentActivityFOrder, FUN = length),
                idvar = "PBptype", timevar = "Num_S", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -c(1,ncol(Temp_pct))]


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=22, colIndex=1, "Student School Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=23, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=23, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=22, colIndex=1, "Student School Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=23, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=23, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=37, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=37, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Student School Mandatory Activity==============================#

#=============summary for Student School Mandatory Activity==============================
Temp <- reshape(aggregate(pid~PBptype+Num_W, data = StudentActivityFOrder, FUN = length),
                idvar = "PBptype", timevar = "Num_W", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, All=rowSums(Temp))

Temp_pct <- Temp/Temp$All

Temp0 <- Temp_pct[, -ncol(Temp_pct)]


xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=29, colIndex=1, "Student Work Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=30, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=30, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=29, colIndex=1, "Student Work Activity", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=30, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=30, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.1 to 3`, 
              startRow=45, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.1 to 3`, rowIndex=45, colIndex=9, "Person Type", "T_Cell_1")

#=============summary for Student School Mandatory Activity==============================#


#=============summary for Mandatory Tour by person type==============================
if(file.exists(paste0(ctrampOutPath,"/mandatoryPrelimTod_cleaned.csv"))){
  MandActivityTOD <- data.frame(fread(paste0(ctrampOutPath,"/mandatoryPrelimTod_cleaned.csv")))
} else{
  MandActivityTOD <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/mandatoryPrelimTod.csv"))
}
MandActivityTOD <- mutate(MandActivityTOD, MActivity=ifelse(mandPrelimActPurps2=="[]", 1, 2))

MandActivity <- pMod[,c("pid", "hhid", "personType")] %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  left_join(MandActivityTOD, by=c("pid"="pid", "hhid"="hhid")) %>%
  mutate(MActivity=ifelse(is.na(MActivity), 0, MActivity))

Temp <- reshape(aggregate(pid~PBptype+MActivity, data = MandActivity, FUN = length),
                idvar = "PBptype", timevar = "MActivity", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- Temp/Temp$Total

Temp0 <- Temp[, -c(1, ncol(Temp))]
Temp0 <- Temp0/rowSums(Temp0)
Temp0[is.na(Temp0)] <- 0

xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=36, colIndex=1, "Mandatory Tour Frequency By Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MA_TS, sheet = sheet_num_data, 
              startRow=37, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_num_data, rowIndex=37, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=36, colIndex=1, "Mandatory Tour Frequency By Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MA_TS, sheet = sheet_pct_data, 
              startRow=37, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheet_pct_data, rowIndex=37, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MA_TS, sheet = sheets$`5.2.2.5`, 
              startRow=17, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MA_TS, sheets$`5.2.2.5`, rowIndex=17, colIndex=1, "CT RAMP Person Type", "T_Cell_1")

#=============summary for Mandatory Tour by person type==============================#
saveWorkbook(Target_MA_TS, file = paste0(Model_path, "Outputs/user/== VT09 20180502 Mandatory activity frequency and tour skeletons (5.2.2).xlsx"))


rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])
