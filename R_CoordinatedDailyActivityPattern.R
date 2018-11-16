CDAP_person <- fread(paste0(ctrampOutPath,"coordinatedDailyActivityPattern_person.csv"))
names(CDAP_person) <- gsub('.{2}$', '', names(CDAP_person)) # Remove last two characters from all column names
PType_map <-  data.frame(personType=c(1:8),
                         PBptype= c("P1FTW","P2PTW", "P3Col", "P4NWadult", "P5NWold", 
                                    "P6_16_18", "P7_6_15", "P8_0_5"),stringsAsFactors = F)
CDAP_person <- left_join(CDAP_person, PType_map, by=c("personType"="personType"))


if(file.exists(paste0(ctrampOutPath,"/coordinatedDailyActivityPattern_hh_cleaned.csv"))){
  CDAP_hh <- data.frame(fread(paste0(ctrampOutPath,"/coordinatedDailyActivityPattern_hh_cleaned.csv")))
} else{
CDAP_hh <- cleanAndReadCT2_out(paste0(ctrampOutPath,"coordinatedDailyActivityPattern_hh.csv"))
}

CDAP_hh <- CDAP_hh %>%
  left_join(hMod[, c("hhid", "HHSIZE")], by=c("hhid"="hhid")) %>%
  mutate(HHSize=ifelse(HHSIZE>=4, "4+", HHSIZE)) 


#==============load xlsx and define dataframe format========================
Target_CDAP <- loadWorkbook(paste0(Target_path, 
                                          "= VT08 20180510 Coordinated Daily Activity Pattern (5.1).xlsx"),
                                   password=NULL)

sheet_num_data <- createSheet(Target_CDAP, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_CDAP, sheetName = "pct_Data")
sheets <- getSheets(Target_CDAP)
removeSheet(Target_CDAP, sheetName=sheets[[3]])
removeSheet(Target_CDAP, sheetName=sheets[[4]])
sheets <- getSheets(Target_CDAP)

#==============Share of CDAP pattern by person type==================== 
Temp <- reshape(aggregate(pid~personDaps+PBptype, data = CDAP_person, FUN = length), 
                idvar = "PBptype", timevar = "personDaps", direction = "wide")
names(Temp) <- c("PBptype", "Mandatory", "Non-Mandatory", "Home")
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp[is.na(Temp)] <- 0
Temp <- cbind(Temp, Population=rowSums(Temp))
Temp <- rbind(Temp, All=colSums(Temp))

Temp_pct <- Temp/Temp$Population

Temp0 <- Temp_pct
Temp0$Population <- Temp$Population

xlsx.AddCell(Target_CDAP, sheet_num_data, rowIndex=1, colIndex=1, "Coordinated Daily Activity Pattern Summary", "Title_1")
xlsx.AddCell(Target_CDAP, sheet_num_data, rowIndex=2, colIndex=1, "Share of CDAP pattern by person type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_CDAP, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheet_num_data, rowIndex=3, colIndex=1, "Person Type", "T_Cell_1")


xlsx.AddCell(Target_CDAP, sheet_pct_data, rowIndex=1, colIndex=1, "Coordinated Daily Activity Pattern Summary", "Title_1")
xlsx.AddCell(Target_CDAP, sheet_pct_data, rowIndex=2, colIndex=1, "Share of CDAP pattern by person type", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_CDAP, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheet_pct_data, rowIndex=3, colIndex=1, "Person Type", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_CDAP, sheet = sheets$`5.1.1`, 
              startRow=15, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheets$`5.1.1`, rowIndex=15, colIndex=1, "Person Type", "T_Cell_1")
#==============Share of CDAP pattern by person type==================== #

#============== Share of Households with Joint Travel by Household Siz=============
Temp <-  reshape(aggregate(hhid~HHSize+cdapJointAlt, data = CDAP_hh, FUN = length), 
                 idvar = "HHSize", timevar = "cdapJointAlt", direction = "wide")
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$HHSize
Temp <- Temp[,-1]
names(Temp) <- c("No_Joint_Travel", "Joint_Travel")
Temp <- rbind(Temp, All=colSums(Temp))
Temp <- cbind(Temp, Total=rowSums(Temp))

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_CDAP, sheet_num_data, rowIndex=14, colIndex=1, "Share of Household with Joint Travel by Household Size", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_CDAP, sheet = sheet_num_data, 
              startRow=15, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheet_num_data, rowIndex=15, colIndex=1, "HHSize", "T_Cell_1")

xlsx.AddCell(Target_CDAP, sheet_pct_data, rowIndex=14, colIndex=1, "Share of Household with Joint Travel by Household Size", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_CDAP, sheet = sheet_pct_data, 
              startRow=15, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheet_pct_data, rowIndex=15, colIndex=1, "HHSize", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_CDAP, sheet = sheets$`5.1.2`, 
              startRow=15, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_CDAP, sheets$`5.1.2`, rowIndex=15, colIndex=1, "HHSize", "T_Cell_1")
#============== Share of Households with Joint Travel by Household Siz=============#


saveWorkbook(Target_CDAP, file = paste0(Model_path, "Outputs/user/== VT08 20180510 Coordinated Daily Activity Pattern (5.1).xlsx"))



rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])