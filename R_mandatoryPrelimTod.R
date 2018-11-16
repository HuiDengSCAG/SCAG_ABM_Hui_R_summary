if(file.exists(paste0(ctrampOutPath,"/mandatoryPrelimTod_cleaned.csv"))){
  MandActivityTOD <- data.frame(fread(paste0(ctrampOutPath,"/mandatoryPrelimTod_cleaned.csv")))
} else{
  MandActivityTOD <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/mandatoryPrelimTod.csv"))
}

MandActivityTOD1 <- MandActivityTOD[, 1:18]
MandActivityTOD2 <- cbind(MandActivityTOD[, c(1:3)], MandActivityTOD[, c(19:ncol(MandActivityTOD))])
MandActivityTOD2 <- MandActivityTOD2[which(MandActivityTOD2$mandatoryPrelimTodChosenAlt2!=0), ]

names(MandActivityTOD1) <- gsub("1", "", names(MandActivityTOD1))
names(MandActivityTOD2) <- gsub("2", "", names(MandActivityTOD2))

MandActivityTOD <- rbind(MandActivityTOD1, MandActivityTOD2)

PType_map <-  data.frame(personType=c(1:8),
                         PBptype= c("P1FTW","P2PTW", "P3Col", "P4NWadult", "P5NWold", 
                                    "P6_16_18", "P7_6_15", "P8_0_5"),stringsAsFactors = F)

MandActivityTOD <- MandActivityTOD %>%
  left_join(pMod[,c("pid", "hhid", "personType")], by=c("pid"="pid", "hhid"="hhid")) %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  mutate(MPPA_Arr=as.integer(mandPrelimPrimaryActArrive/4+3)%%24+1) %>%
  mutate(MPPA_Dep=as.integer(mandPrelimPrimaryActDepart/4+3)%%24+1) %>%
  mutate(MPPA_Dur=as.integer(mandPrelimDurationTour/4)+1) 

#==============load xlsx and define dataframe format========================
Target_MandTOD <- loadWorkbook(paste0(Target_path, 
                                    "= VT10 20180502 Preliminary mandatory activity schedule (5.2.3).xlsx"),
                             password=NULL)

sheet_num_data <- createSheet(Target_MandTOD, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_MandTOD, sheetName = "pct_Data")
sheets <- getSheets(Target_MandTOD)
removeSheet(Target_MandTOD, sheetName=sheets[[3]])
removeSheet(Target_MandTOD, sheetName=sheets[[4]])
removeSheet(Target_MandTOD, sheetName=sheets[[5]])
removeSheet(Target_MandTOD, sheetName=sheets[[6]])
sheets <- getSheets(Target_MandTOD)

#==========Preliminary mandatory activity schedule Summary==============#
#======== mandatory activity's Arrival time at mandatory location============
Temp <- reshape(aggregate(pid~MPPA_Arr+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Arr", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Arr
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[, -ncol(Temp_pct)]
Temp0["4", ] <- colSums(Temp0[c("2", "3", "4"), ])
Temp0 <- Temp0[-which(rownames(Temp0) %in% c("2", "3")), ]


xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=1, colIndex=1, "Preliminary mandatory activity schedule Summary", "Title_1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=2, colIndex=1, "mandatory activity's Arrival time", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=3, colIndex=1, "Hour", "T_Cell_1")


xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=1, colIndex=1, "Preliminary mandatory activity schedule Summary", "Title_1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=2, colIndex=1, "mandatory activity's Arrival time", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=3, colIndex=1, "Hour", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time`, 
              startRow=8, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time`, rowIndex=8, colIndex=13, "Hour", "T_Cell_1")

#======== mandatory activity's Arrival time at mandatory location============#

#======== mandatory activity's Departure time at mandatory location============
Temp <- reshape(aggregate(pid~MPPA_Dep+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Dep", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Dep
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[-nrow(Temp_pct), -ncol(Temp_pct)]
Temp0["4", ] <- colSums(Temp0[which(rownames(Temp0) %in% c("2", "3", "4")), ])
Temp0 <- Temp0[-which(rownames(Temp0) %in% c("2", "3")), ]
Temp0 <- Temp0[order(as.integer(rownames(Temp0))), ]

xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=30, colIndex=1, "mandatory activity's Arrival time", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=31, colIndex=1, "Hour", "T_Cell_1")

xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=30, colIndex=1, "mandatory activity's Arrival time", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=31, colIndex=1, "Hour", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time`, 
              startRow=36, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time`, rowIndex=36, colIndex=13, "Hour", "T_Cell_1")
#======== mandatory activity's Departure time at mandatory location============#

#============== mandatory activity's Duration time =============================
Temp <- reshape(aggregate(pid~MPPA_Dur+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Dur", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Dur
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[-nrow(Temp_pct), -ncol(Temp_pct)]


xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=58, colIndex=1, "mandatory activity's Duration", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=59, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=59, colIndex=1, "Hour", "T_Cell_1")

xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=58, colIndex=1, "mandatory activity's Duration", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=59, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=59, colIndex=1, "Hour", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time`, 
              startRow=64, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time`, rowIndex=64, colIndex=13, "Hour", "T_Cell_1")

#============== mandatory activity's Duration time ============================#



MandActivityTOD <- MandActivityTOD %>%
  mutate(MPPA_Arr_TOD=cut(MPPA_Arr, breaks = c(0,6,9,15,19,21,24), 
                          labels = c("NT", "AM", "MD", "PM", "EV", "NT"),
                          include.lowest = T, right = F)) %>%
  mutate(MPPA_Dep_TOD=cut(MPPA_Dep, breaks = c(0,6,9,15,19,21,24), 
                          labels = c("NT", "AM", "MD", "PM", "EV", "NT"), 
                          include.lowest = T, right = F)) %>% 
  mutate(MPPA_Dur_Cat=cut(MPPA_Dur, breaks = c(0, 3, 6, 10, 12, 24),
                          include.lowest = F, right = T))

MandActivityTOD$MPPA_Arr_TOD <- factor(MandActivityTOD$MPPA_Arr_TOD, levels = c("AM", "MD", "PM", "EV", "NT"))
MandActivityTOD$MPPA_Dep_TOD <- factor(MandActivityTOD$MPPA_Dep_TOD, levels = c("AM", "MD", "PM", "EV", "NT"))

#======== mandatory activity's Arrival time at mandatory location TOD============
Temp <- reshape(aggregate(pid~MPPA_Arr_TOD+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Arr_TOD", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Arr
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[-nrow(Temp_pct), -ncol(Temp_pct)]



xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=86, colIndex=1, "Preliminary mandatory activity schedule Summary TOD", "Title_1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=87, colIndex=1, "mandatory activity's Arrival time TOD", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=88, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=88, colIndex=1, "TOD", "T_Cell_1")


xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=86, colIndex=1, "Preliminary mandatory activity schedule Summary TOD", "Title_1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=87, colIndex=1, "mandatory activity's Arrival time TOD", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=88, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=88, colIndex=1, "TOD", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time (2)`, 
              startRow=7, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time (2)`, rowIndex=7, colIndex=9, "TOD", "T_Cell_1")

#======== mandatory activity's Arrival time at mandatory location TOD============#

#======== mandatory activity's Departure time at mandatory location TOD============
Temp <- reshape(aggregate(pid~MPPA_Dep_TOD+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Dep_TOD", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Dep
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[-nrow(Temp_pct), -ncol(Temp_pct)]

xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=96, colIndex=1, "mandatory activity's Arrival time TOD", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=97, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=97, colIndex=1, "TOD", "T_Cell_1")

xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=96, colIndex=1, "mandatory activity's Arrival time TOD", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=97, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=97, colIndex=1, "TOD", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time (2)`, 
              startRow=18, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time (2)`, rowIndex=18, colIndex=9, "TOD", "T_Cell_1")
#======== mandatory activity's Departure time at mandatory location TOD============#

#============== mandatory activity's Duration time TOD =============================
Temp <- reshape(aggregate(pid~MPPA_Dur_Cat+PBptype, data = MandActivityTOD, FUN = length),
                idvar = "MPPA_Dur_Cat", timevar = "PBptype", direction = "wide")
names(Temp) <- gsub("pid.", "", names(Temp))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$MPPA_Dur
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)

Temp0 <- Temp_pct[-nrow(Temp_pct), -ncol(Temp_pct)]


xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=105, colIndex=1, "mandatory activity's Duration TOD", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_MandTOD, sheet = sheet_num_data, 
              startRow=106, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_num_data, rowIndex=106, colIndex=1, "TOD", "T_Cell_1")

xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=105, colIndex=1, "mandatory activity's Duration", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_MandTOD, sheet = sheet_pct_data, 
              startRow=106, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheet_pct_data, rowIndex=106, colIndex=1, "TOD", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_MandTOD, sheet = sheets$`5.2.2.4 Mand Time (2)`, 
              startRow=29, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_MandTOD, sheets$`5.2.2.4 Mand Time (2)`, rowIndex=29, colIndex=9, "TOD", "T_Cell_1")

#============== mandatory activity's Duration time TOD============================#



saveWorkbook(Target_MandTOD, file = paste0(Model_path, "Outputs/user/== VT10 20180502 Preliminary mandatory activity schedule (5.2.3).xlsx"))


rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])
