#==============load xlsx and define dataframe format========================
Target_PrimaryDest <- loadWorkbook(paste0(Target_path, 
                                        "= VD 20181023 Primary Destination.xlsx"),
                                 password=NULL)

sheet_num_data <- createSheet(Target_PrimaryDest, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_PrimaryDest, sheetName = "pct_Data")
sheets <- getSheets(Target_PrimaryDest)
removeSheet(Target_PrimaryDest, sheetName=sheets[[2]])
removeSheet(Target_PrimaryDest, sheetName=sheets[[3]])
sheets <- getSheets(Target_PrimaryDest)

#==============Distance to Primary Destination ==================== 
TourList_half <- TripList1 %>%
  mutate(tripTime=tripArriveMinute-tripDepartMinute) %>%
  filter(mcTourType != 4) %>%
  group_by(hhid, pid, persTourNum, tourPurpose, mcTourType, TourPurp, PBptype, tripDir) %>%
  summarise(Trips_in_tour=n(), TourLength = sum(tripDist), TourTime=sum(tripTime)) %>%
  ungroup %>%
  mutate(TourDist_Bin=cut(TourLength, c(0:300), labels = c(0:299)+0.5))

Temp <- TourList_half %>%
  group_by(TourPurp) %>%
  summarise(Trips_to_PrimaryDest=mean(Trips_in_tour), Distance_to_PrimaryDest = mean(TourLength), 
            TravelTime_to_PrimaryDest=mean(TourTime)) %>%
  data.frame

rownames(Temp) <- Temp$TourPurp
Temp <- Temp[,-1]
Temp <- rbind(Temp, All_purpose=c(mean(TourList_half$Trips_in_tour), mean(TourList_half$TourLength),
                                  mean(TourList_half$TourTime)))


xlsx.AddCell(Target_PrimaryDest, sheet_num_data, rowIndex=1, colIndex=1, "Primary Destination Summary", "Title_1")
xlsx.AddCell(Target_PrimaryDest, sheet_num_data, rowIndex=2, colIndex=1, "Average Trips Distance and Time to Primary Destination", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_PrimaryDest, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheet_num_data, rowIndex=3, colIndex=1, "F_TourPurpose", "T_Cell_1")


xlsx.AddCell(Target_PrimaryDest, sheet_pct_data, rowIndex=1, colIndex=1, "Primary Destination Summary", "Title_1")
xlsx.AddCell(Target_PrimaryDest, sheet_pct_data, rowIndex=2, colIndex=1, "Average Trips Distance and Time to Primary Destination", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_PrimaryDest, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheet_pct_data, rowIndex=3, colIndex=1, "F_TourPurpose", "T_Cell_1")


xlsx.AddTable(as.data.frame(Temp), Target_PrimaryDest, sheet = sheets$`Distance to Primary Destibnatio`, 
              startRow=22, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheets$`Distance to Primary Destibnatio`, rowIndex=22, colIndex=1, "Primary Purpose", "T_Cell_1")

#==============Distance to Primary Destination ====================#


#==============Distance to Primary Destination Distribution by Primary Purpose ================
PrimaryDest_TLD <- reshape(aggregate(persTourNum~TourDist_Bin+TourPurp, data = TourList_half, FUN = length), 
                    idvar = "TourDist_Bin", timevar = "TourPurp", direction = "wide")
colnames(PrimaryDest_TLD) <- gsub("persTourNum.", "", colnames(PrimaryDest_TLD))
Temp <- data.frame(levels(PrimaryDest_TLD$TourDist_Bin))
names(Temp) <- "Dist_PrimaryDest"
Temp$Dist_PrimaryDest <- as.character(Temp$Dist_PrimaryDest)
PrimaryDest_TLD$TourDist_Bin <- as.character(PrimaryDest_TLD$TourDist_Bin)
Temp <- left_join(Temp, PrimaryDest_TLD, by=c("Dist_PrimaryDest"="TourDist_Bin"))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$Dist_PrimaryDest
Temp <- Temp[,c(-1)]
Temp <- cbind(Temp, All_Purpose=rowSums(Temp))
Temp <- rbind(Temp, Total=colSums(Temp))

Temp_pct <- as.data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$Total
Temp_pct <- as.data.frame(t(Temp_pct))


xlsx.AddCell(Target_PrimaryDest, sheet_num_data, rowIndex=21, colIndex=1, "Distance to Primary Destination Distribution by purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_PrimaryDest, sheet = sheet_num_data, 
              startRow=22, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheet_num_data, rowIndex=22, colIndex=1, "Distance to Primary Destination", "T_Cell_1")

xlsx.AddCell(Target_PrimaryDest, sheet_pct_data, rowIndex=21, colIndex=1, "Distance to Primary Destination Distribution by purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_PrimaryDest, sheet = sheet_pct_data, 
              startRow=22, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheet_pct_data, rowIndex=22, colIndex=1, "Distance to Primary Destination", "T_Cell_1")


xlsx.AddTable(as.data.frame(Temp_pct[-c(76:nrow(Temp_pct)),]), Target_PrimaryDest, sheet = sheets$`Distance to Primary Destibnatio`, 
              startRow=4, startColumn=22, CellColor = "darkseagreen1")
xlsx.AddCell(Target_PrimaryDest, sheets$`Distance to Primary Destibnatio`, rowIndex=4, colIndex=22, "Length", "T_Cell_1")

#==============Distance to Primary Destination Distribution by Primary Purpose ================#
saveWorkbook(Target_PrimaryDest, file = paste0(Model_path, "Outputs/user/== VD 20181023 Primary Destination.xlsx"))



rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])


