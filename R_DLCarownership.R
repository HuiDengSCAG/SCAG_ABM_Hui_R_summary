

### Model Files #####
DriversLicenseMod <- fread(paste0(ctrampOutPath,"driversLicense.csv"))
names(DriversLicenseMod) <- gsub('.{2}$', '', names(DriversLicenseMod)) # Remove last two characters from all column names

DriversLicenseMod <-  DriversLicenseMod %>%
  left_join(pMod[, c("pid", "hhid", "personType", "AGE", "GENDER")], by=c("pid"="pid", "hhid"="hhid")) %>%
  left_join(hMod[, c("hhid", "HCOUNTY")], by=c("hhid"="hhid"))

DriversLicenseMod <- mutate(DriversLicenseMod, Age=cut(DriversLicenseMod$AGE, c(-Inf, 20, 40, 55, 70, 80, Inf), 
                                             labels=c("16-20", "21-40", "41-55", "56-70", "71-80", "81-99")))

AutoOwnershipMod <- fread(paste0(ctrampOutPath,"autoOwnership.csv"))
names(AutoOwnershipMod) <- gsub('.{2}$', '', names(AutoOwnershipMod)) # Remove last two characters from all column names

HHWorker <- hMod %>%
  left_join(pMod[, c("pid", "hhid", "WORKER")], by=c("hhid"="hhid")) %>%
  mutate(Employed=ifelse(WORKER==1, 1, 0)) %>%
  group_by(hhid) %>%
  summarise(HHWorkers = sum(Employed)) %>%
  ungroup()


AutoOwnershipMod <- AutoOwnershipMod %>%
  left_join(hMod[,c("hhid", "HCOUNTY", "HHSIZE")], by=c("hhid"="hhid")) %>%
  left_join(HHWorker,by=c("hhid"="hhid")) %>%
  mutate(HHSIZECat = ifelse(HHSIZE>=4, "4+", HHSIZE)) %>%
  mutate(HHWorkersCat = ifelse(HHWorkers>=4, "4+", HHWorkers)) %>%
  mutate(NCars = ifelse(autoOwnershipChosenAlt-1>=4, "4+", autoOwnershipChosenAlt-1)) %>%
  mutate(NumCars = ifelse(autoOwnershipChosenAlt-1>=4, 4.5, autoOwnershipChosenAlt-1))
  



#==============load xlsx and define dataframe format========================
Target_DL_Car_Ownership <- loadWorkbook(paste0(Target_path, 
                                                  "= VT06 07 20180713 Driver's License and Household Vehicles.xlsx"),
                                           password=NULL)
sheets <- getSheets(Target_DL_Car_Ownership)
removeSheet(Target_DL_Car_Ownership, sheetName=sheets[[3]])
removeSheet(Target_DL_Car_Ownership, sheetName=sheets[[4]])
sheets <- getSheets(Target_DL_Car_Ownership)
sheet_num_data <- createSheet(Target_DL_Car_Ownership, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_DL_Car_Ownership, sheetName = "pct_Data")

#==============DrvierLicense by persontype===============================
Temp <- reshape(aggregate(pid~personType+driversLicenseChosenAlt, data = DriversLicenseMod, FUN = length),
                idvar = "personType", timevar = "driversLicenseChosenAlt", direction = "wide")
rownames(Temp) <- Temp$personType
Temp <- Temp[, c(-1)]

Temp$Total <- rowSums(Temp)

names(Temp) <- gsub("pid.1", "No_DL", names(Temp))
names(Temp) <- gsub("pid.2", "DriverLicense", names(Temp))

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=1, colIndex=1, "Driver License and Auto Ownership", "Title_1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=2, colIndex=1, "Driver License by person type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=3, colIndex=1, "personType\\Driver_License", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=1, colIndex=1, "Driver License and Auto Ownership", "Title_1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=2, colIndex=1, "Driver License by person type", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=3, colIndex=1, "personType\\Driver_License", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct$DriverLicense), Target_DL_Car_Ownership, sheet = sheets$`3.1 license`, 
              startRow=5, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.1 license`, rowIndex=5, colIndex=9, "personType", "T_Cell_1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.1 license`, rowIndex=5, colIndex=10, "\\% License", "T_Cell_1")

#==============DrvierLicense by personType===============================#

#==============DrvierLicense by Age and Gender===============================
Temp <- reshape(aggregate(pid~Age+GENDER, 
                           data = DriversLicenseMod[which(DriversLicenseMod$driversLicenseChosenAlt==2),], FUN = length),
                idvar = "Age", timevar = "GENDER", direction = "wide")
rownames(Temp) <- Temp$Age
Temp <- Temp[, c(-1)]
names(Temp) <- c("Male", "Female")
Temp$Total <- rowSums(Temp)

Temp2 <- reshape(aggregate(pid~Age+GENDER, data = DriversLicenseMod, FUN = length),
                 idvar = "Age", timevar = "GENDER", direction = "wide")
rownames(Temp2) <- Temp2$Age
Temp2 <- Temp2[, c(-1)]
names(Temp2) <- c("Male", "Female")
Temp2$Total <- rowSums(Temp2)

Temp_pct <- Temp/Temp2

xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=11, colIndex=1, "Driver License by Age and Gender", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=12, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=12, colIndex=1, "Age\\Gender", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=11, colIndex=1, "Driver License by Age and Gender", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=12, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=12, colIndex=1, "Age\\Gender", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheets$`3.1 license`, 
              startRow=15, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.1 license`, rowIndex=15, colIndex=9, "Age\\Gender", "T_Cell_1")

#==============DrvierLicense by Age and Gender==============================#

#==============DrvierLicense by County===============================
Temp <- reshape(aggregate(pid~HCOUNTY+driversLicenseChosenAlt, 
                          data = DriversLicenseMod, FUN = length),
                idvar = "HCOUNTY", timevar = "driversLicenseChosenAlt", direction = "wide")
rownames(Temp) <- Temp$HCOUNTY
Temp <- Temp[, c(-1)]

Temp$Total <- rowSums(Temp)

names(Temp) <- gsub("pid.1", "No_DL", names(Temp))
names(Temp) <- gsub("pid.2", "DriverLicense", names(Temp))

Temp <- rbind(Temp, colSums(Temp))
rownames(Temp) <- c("IM", "LA", "OR", "RIV", "SBD", "VN", "SCAG")

Temp_pct <- Temp/Temp$Total
Temp_1 <- as.data.frame(Temp[, c("DriverLicense")])
rownames(Temp_1) <- rownames(Temp)
names(Temp_1) <- "DriverLicense"


xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=20, colIndex=1, "Driver License by County", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=21, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=21, colIndex=1, "County", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=20, colIndex=1, "Driver License by County", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=21, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=21, colIndex=1, "County", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_1), Target_DL_Car_Ownership, sheet = sheets$`3.1 license`, 
              startRow=25, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.1 license`, rowIndex=25, colIndex=9, "County\\DL", "T_Cell_1")

#==============DrvierLicense by County===================================#

#==============HH Number of vehicle by HHSIZE===============================
Temp <- reshape(aggregate(hhid~HHSIZECat+NCars, 
                          data = AutoOwnershipMod, FUN = length),
                idvar = "HHSIZECat", timevar = "NCars", direction = "wide")
rownames(Temp) <- Temp$HHSIZECat
Temp <- Temp[, c(-1)]

Temp$Total <- rowSums(Temp)

names(Temp) <- gsub("hhid.", "", names(Temp))

Temp <- rbind(Temp, colSums(Temp))
rownames(Temp) <- c("1", "2", "3", "4+", "All")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=30, colIndex=1, "Household Auto by HHSIZE", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=31, colIndex=1, "Hhsize\\Hhveh", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=30, colIndex=1, "Household Auto by HHSIZE", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=31, colIndex=1, "Hhsize\\Hhveh", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct[, -c(6)]), Target_DL_Car_Ownership, sheet = sheets$`3.2 HHVEH`, 
              startRow=6, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.2 HHVEH`, rowIndex=6, colIndex=9, "Hhsize\\Hhveh", "T_Cell_1")

#==============HH Number of vehicle by HHSIZE==============================#

#==============HH Number of vehicle by HH number of workers===============================
Temp <- reshape(aggregate(hhid~HHWorkersCat+NCars, 
                          data = AutoOwnershipMod, FUN = length),
                idvar = "HHWorkersCat", timevar = "NCars", direction = "wide")
rownames(Temp) <- Temp$HHWorkersCat
Temp <- Temp[, c(-1)]

Temp$Total <- rowSums(Temp)

names(Temp) <- gsub("hhid.", "", names(Temp))

Temp <- rbind(Temp, colSums(Temp))
rownames(Temp) <- c("0", "1", "2", "3", "4+", "All")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=38, colIndex=1, "Household Auto by HHworker", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=39, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=39, colIndex=1, "Hhworker\\Hhveh", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=38, colIndex=1, "Household Auto by HHworker", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=39, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=39, colIndex=1, "Hhworker\\Hhveh", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct[-5, -6]), Target_DL_Car_Ownership, sheet = sheets$`3.2 HHVEH`, 
              startRow=16, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.2 HHVEH`, rowIndex=16, colIndex=9, "Hhworker\\Hhveh", "T_Cell_1")

#==============HH Number of vehicle by HH number of workers==============================#

#==============HH County by HH number of vehicle ===============================
Temp <- reshape(aggregate(hhid~HCOUNTY+NCars, 
                          data = AutoOwnershipMod, FUN = length),
                idvar = "HCOUNTY", timevar = "NCars", direction = "wide")
rownames(Temp) <- Temp$HCOUNTY
Temp <- Temp[, c(-1)]

Temp$Total <- rowSums(Temp)

names(Temp) <- gsub("hhid.", "", names(Temp))

Temp <- rbind(Temp, colSums(Temp))
rownames(Temp) <- c("IM", "LA", "OR", "RIV", "SBD", "VN", "SCAG")

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=47, colIndex=1, "County by HH Auto", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=48, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=48, colIndex=1, "County\\Hhveh", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=47, colIndex=1, "County by HH Auto", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=48, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=48, colIndex=1, "County\\Hhveh", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct[-7, -6]), Target_DL_Car_Ownership, sheet = sheets$`3.2 HHVEH`, 
              startRow=26, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.2 HHVEH`, rowIndex=26, colIndex=9, "County\\Hhveh", "T_Cell_1")

#==============HH County by HH number of vehicle ================================#

#==============County by number of vehicles ===============================
Temp <- aggregate(NumCars~HCOUNTY, data = AutoOwnershipMod, FUN = sum)
Temp$HCOUNTY <- c("IM", "LA", "OR", "RIV", "SBD", "VN")
SCAG_NumCars <- sum(Temp$NumCars)
names(Temp) <- c("County", "Vehicles")
Temp_pct <- Temp
Temp <- rbind(Temp, c("SCAG", SCAG_NumCars))
Temp_pct$Vehicles <- Temp_pct$Vehicles/SCAG_NumCars
Temp_pct <- rbind(Temp_pct, c("SCAG", 1))


Temp0 <- data.frame(Temp[,-1])
names(Temp0) <- "HH Vehicles"
rownames(Temp0) <- c("IM", "LA", "OR", "RIV", "SBD", "VN", "SCAG_Total")

xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=57, colIndex=1, "Vehicles by County", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_DL_Car_Ownership, sheet = sheet_num_data, 
              startRow=58, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_num_data, rowIndex=58, colIndex=1, "County\\Hhveh", "T_Cell_1")


xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=57, colIndex=1, "Vehicles by County", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_DL_Car_Ownership, sheet = sheet_pct_data, 
              startRow=58, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheet_pct_data, rowIndex=58, colIndex=1, "County\\Hhveh", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_DL_Car_Ownership, sheet = sheets$`3.2 HHVEH`, 
              startRow=36, startColumn=9, CellColor = "darkseagreen1")
xlsx.AddCell(Target_DL_Car_Ownership, sheets$`3.2 HHVEH`, rowIndex=36, colIndex=9, "County", "T_Cell_1")

#==============County by number of vehicles ================================#

saveWorkbook(Target_DL_Car_Ownership, file = paste0(Model_path, "Outputs/user/== VT06 07 20180713 Driver's License and Household Vehicles.xlsx"))

rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", 
                           "hMod", "pMod")])