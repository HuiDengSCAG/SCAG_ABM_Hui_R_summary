#==========begin summary of Trips with ABM purpose===========
cleanAndReadCT2_out <- function(fName){
  
  inpLines <- readLines(fName)
  headerLine <- paste0(gsub('.{2}$', '', unlist(str_split(inpLines[[1]],','))),collapse = ',') # Cleaning up the header row
  # adding a double quote before square bracket opening and and after a square bracket ending
  cleanted_fName <- gsub('.csv','_cleaned.csv',fName)
  writeLines(append(headerLine,unlist(lapply(2:length(inpLines), function(iLine) gsub("\\]",'\\]"',gsub("\\[",'"\\[',inpLines[[iLine]]))  ))),
             cleanted_fName)
  return(fread(cleanted_fName) %>% data.table)
  
}

if(file.exists(paste0(ctrampOutPath,"/output_disaggTripList_cleaned.csv"))){
  TripList <- data.frame(fread(paste0(ctrampOutPath,"/output_disaggTripList_cleaned.csv")))
} else{
  TripList <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/output_disaggTripList.csv"))
}


#refer to trip purpose and Mode map
TripPurp_Map <- read.csv(paste0(Target_path,"TripPurp.csv"), header=T, stringsAsFactors=FALSE)
names(TripPurp_Map) <- c("tripPurpOrig", "tripPurpDest", "TripPurp")
Mode_map <- data.frame(tripMode=c(1:14),
                       Tripmode= c("SOV","HOV2Dr","HOV3Dr","HOVPass",
                                   "Walk Bus","KNR Bus","PNR Bus","Walk Rail","KNR Rail","PNR Rail",
                                   "Walk","Bike","Taxi","School Bus"),stringsAsFactors = F)
PType_map <-  data.frame(personType=c(1:8),
                         PBptype= c("P1FTW","P2PTW", "P3Col", "P4NWadult", "P5NWold", 
                                    "P6_16_18", "P7_6_15", "P8_0_5"),stringsAsFactors = F)
TourPurpose_map <- data.frame(tourPurpose=c(1,15,2,3,411,42,5,6,61,62,7,71,72,73,8,9,
                                            1,15,2,3,411,42,5,6,61,62,7,71,72,73,8,9),
                              mcTourType_3=c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                                             0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),
                              TourPurp= c("01Work", "01Work","02Univ","03School","04PureEsc","05IndEsc",
                                          "15JT_shop","16JT_main","16JT_main","16JT_main","17JT_eat","17JT_eat",
                                          "17JT_eat","17JT_eat","18JT_visit","19JT_discr",
                                          "01Work","01Work","02Univ","03School","04PureEsc","05IndEsc",
                                          "10Ind_shop","111nd_main","111nd_main","111nd_main","121nd_eat","121nd_eat",
                                          "121nd_eat","121nd_eat","131nd_visi","141nd_disc"),stringsAsFactors = F)
  
TripList1 <- TripList %>%
  left_join(TripPurp_Map, by=c("tripPurpOrig"="tripPurpOrig", "tripPurpDest"="tripPurpDest")) %>%
  mutate(TripPurp=ifelse(is.na(TripPurp), "OBO", TripPurp)) %>%
  left_join(Mode_map, by=c("tripMode"="tripMode")) %>%
  left_join(pMod[,c("pid", "hhid", "personType")], by=c("pid"="pid", "hhid"="hhid")) %>%
  left_join(PType_map, by=c("personType"="personType")) %>%
  mutate(tripMinute=tripArriveMinute-tripDepartMinute) %>%
  mutate(TripDist_Bin=cut(tripDist, c(0:300), labels = c(0:299)+0.5)) %>%
  mutate(mcTourType_3=ifelse(mcTourType==3, 3, 0)) %>%
  left_join(TourPurpose_map, by=c("tourPurpose"="tourPurpose", "mcTourType_3"="mcTourType_3"))


#==============load xlsx and define dataframe format========================
Target_Trip_Tour <- loadWorkbook(paste0(Target_path, 
                                               "= VD 20181023 Trip Tour Length.xlsx"),
                                        password=NULL)

sheet_num_data <- createSheet(Target_Trip_Tour, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_Trip_Tour, sheetName = "pct_Data")
sheets <- getSheets(Target_Trip_Tour)
#removeSheet(Target_Trip_Tour, sheetName=sheets[[3]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[4]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[7]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[8]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[9]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[10]])
#removeSheet(Target_Trip_Tour, sheetName=sheets[[11]])
#sheets <- getSheets(Target_Trip_Tour)



#==============ABM Trips by Trip Purpose and Mode ===============================
Trip_Mode_Purp <- reshape(aggregate(tripId~Tripmode+TripPurp, data = TripList1, FUN = length), 
                          idvar = "TripPurp", timevar = "Tripmode", direction = "wide")
colnames(Trip_Mode_Purp) <- gsub("tripId.", "", colnames(Trip_Mode_Purp))

Trip_Mode_Purp[is.na(Trip_Mode_Purp)] <-   0

Temp <- Trip_Mode_Purp
rownames(Temp) <- Temp$TripPurp
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))
row.names(Temp)[nrow(Temp)] <- "All_Purpose"

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=1, colIndex=1, "Trip and Tour Summary", "Title_1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=2, colIndex=1, "Trip by ABM Purpose and Mode", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=3, colIndex=1, "ABM Purpose\\Mode", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=1, colIndex=1, "Trip and Tour Summary", "Title_1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=2, colIndex=1, "Trip by ABM Purpose and Mode", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=3, colIndex=1, "ABM Purpose\\Mode", "T_Cell_1")

#==============ABM Trips by Trip Purpose and Mode ===============================#

#==============ABM Trips and Avg length by Trip Purpose ===============================
Temp <- TripList1 %>%
  group_by(TripPurp) %>%
  summarise(Trips=n(), Avg_Length=mean(tripDist), Avg_Time=mean(tripMinute))
Temp <- as.data.frame(Temp)
rownames(Temp) <- Temp$TripPurp
Temp <- Temp[, c(-1)]
Temp <- rbind(Temp, c(nrow(TripList1), mean(TripList1$tripDist), mean(TripList1$tripMinute)))
row.names(Temp)[nrow(Temp)] <- "Sum"
Temp_pct <- Temp
Temp_pct$Trips <- Temp_pct$Trips/nrow(TripList1)

Temp0 <- Temp[,c("Trips", "Avg_Length")]
Temp0$Trips_pct <- Temp_pct$Trips
Temp0 <- Temp0[, c("Trips", "Trips_pct", "Avg_Length")]

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=16, colIndex=1, "Trips and Avg length by ABM Trip Purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=17, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=17, colIndex=1, "ABM TripPurpose", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=16, colIndex=1, "Trips and Avg length by ABM Trip Purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=17, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=17, colIndex=1, "ABM TripPurpose", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_Trip_Tour, sheet = sheets$`01 Trip & Length`, 
              startRow=4, startColumn=6, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`01 Trip & Length`, rowIndex=4, colIndex=6, "ABMPurp", "T_Cell_1")

#==============ABM Trips and Avg length by Trip Purpose ===============================#

#==============ABM Trips by person type ===============================
Temp <- TripList1 %>%
  group_by(personType) %>%
  summarise(Trips=n())
Temp0 <- pMod %>% group_by(personType) %>% summarise(Persons=n())

Temp <- Temp0 %>%
  left_join(Temp, by="personType") %>%
  left_join(PType_map, by="personType") 
Total_Persons <- sum(Temp$Persons)
Total_Trips <- sum(Temp$Trips)
Temp <- Temp %>%
  rbind(c(1, Total_Persons, Total_Trips, "Total")) %>%
  data.frame() 
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, c(-1,-4)]
Temp$Persons <- as.numeric(Temp$Persons)
Temp$Trips <- as.numeric(Temp$Trips)
Temp$Trip_per_Person <- Temp$Trips/Temp$Persons

Temp_pct <- Temp
Temp_pct$Persons=Temp_pct$Persons/Total_Persons
Temp_pct$Trips=Temp_pct$Trips/Total_Trips


xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=30, colIndex=1, "Trip by Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=31, colIndex=1, "PBpType", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=30, colIndex=1, "Trip by Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=31, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=31, colIndex=1, "PBpType", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheets$`01 Trip & Length`, 
              startRow=18, startColumn=6, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`01 Trip & Length`, rowIndex=18, colIndex=6, "PBpType", "T_Cell_1")

#==============ABM Trips by person type  ===============================#

#==============Trips by person type and Trip purpose ===============================
Trip_pType_Purp <- reshape(aggregate(tripId~PBptype+TripPurp, data = TripList1, FUN = length), 
                          idvar = "TripPurp", timevar = "PBptype", direction = "wide")
colnames(Trip_pType_Purp) <- gsub("tripId.", "", colnames(Trip_pType_Purp))

Trip_pType_Purp[is.na(Trip_pType_Purp)] <- 0

Temp <- Trip_pType_Purp
rownames(Temp) <- Temp$TripPurp
Temp <- Temp[, c(-1)]
Temp$Total <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))
row.names(Temp)[nrow(Temp)] <- "All_Purpose"

Temp_pct <- as.data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All_Purpose
Temp_pct <- as.data.frame(t(Temp_pct))

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=42, colIndex=1, "Trips by person type and Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=43, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=43, colIndex=1, "ABM Trip Purpose\\ PBptype", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=42, colIndex=1, "Trips by person type and Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=43, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=43, colIndex=1, "ABM Trip Purpose\\ PBptype", "T_Cell_1")
#==============Trips by person type and Trip purpose ===============================#

#==============Trip Length by person type and Trip purpose ===============================
TripLength_pType_Purp <- reshape(aggregate(tripDist~PBptype+TripPurp, data = TripList1, FUN = mean), 
                           idvar = "TripPurp", timevar = "PBptype", direction = "wide")
colnames(TripLength_pType_Purp) <- gsub("tripDist.", "", colnames(TripLength_pType_Purp))

#TripLength_pType_Purp[is.na(TripLength_pType_Purp)] <- 0



Temp <- TripLength_pType_Purp
rownames(Temp) <- Temp$TripPurp
Temp <- Temp[, c(-1)]

Temp0 <- aggregate(tripDist~TripPurp, data = TripList1, FUN = mean)
Temp$Total <- Temp0$tripDist

Temp0 <- aggregate(tripDist~PBptype, data = TripList1, FUN = mean)

Temp <- rbind(Temp, c(Temp0$tripDist, mean(TripList1$tripDist)))
row.names(Temp)[nrow(Temp)] <- "All_Purpose"


xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=56, colIndex=1, "Trip Length by person type and Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=57, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=57, colIndex=1, "ABM Trip Purpose\\ PBptype", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=56, colIndex=1, "Trip Length by person type and Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=57, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=57, colIndex=1, "ABM Trip Purpose\\ PBptype", "T_Cell_1")



xlsx.AddTable(as.data.frame(Temp[-11,-9]), Target_Trip_Tour, sheet = sheets$`01 Trip & Length`, 
              startRow=34, startColumn=11, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`01 Trip & Length`, rowIndex=34, colIndex=11, "PBpType", "T_Cell_1")

#==============Trip Length by person type and Trip purpose ===============================#


#==============Trip Length Distribution by Trip purpose ===============================
TLD_Purpose <- reshape(aggregate(tripId~TripDist_Bin+TripPurp, data = TripList1, FUN = length), 
                       idvar = "TripDist_Bin", timevar = "TripPurp", direction = "wide")
colnames(TLD_Purpose) <- gsub("tripId.", "", colnames(TLD_Purpose))
Temp <- data.frame(levels(TLD_Purpose$TripDist_Bin))
names(Temp) <- "TripLength"
Temp$TripLength <- as.character(Temp$TripLength)
TLD_Purpose$TripDist_Bin <- as.character(TLD_Purpose$TripDist_Bin)
Temp <- left_join(Temp, TLD_Purpose, by=c("TripLength"="TripDist_Bin"))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$TripLength
Temp <- Temp[,c(-1)]
Temp$ALL_Purpose <- rowSums(Temp)
Temp <- rbind(Temp, colSums(Temp))
row.names(Temp)[nrow(Temp)] <- "Total"

Temp_pct <- as.data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$Total
Temp_pct <- as.data.frame(t(Temp_pct))

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=139, colIndex=1, "Trip Length Distribution by Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=140, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=140, colIndex=1, "Trip Length\\ Trip Purpose", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=139, colIndex=1, "Trip Length Distribution by Trip purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=140, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=140, colIndex=1, "Trip Length\\ Trip Purpose", "T_Cell_1")


xlsx.AddTable(as.data.frame(Temp_pct[-c(51:301),-11]), Target_Trip_Tour, sheet = sheets$`02 Trip TLD`, 
              startRow=4, startColumn=12, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`02 Trip TLD`, rowIndex=4, colIndex=12, "Trip Length", "T_Cell_1")

#==============Trip Length Distribution by Trip purpose ==================================#


#==============Tour summary ==============================#
TourList <- TripList1 %>%
  mutate(tripTime=tripArriveMinute-tripDepartMinute) %>%
  filter(mcTourType != 4) %>%
  group_by(hhid, pid, persTourNum, tourPurpose, mcTourType, TourPurp, PBptype) %>%
  summarise(Trips_in_tour=n(), TourLength = sum(tripDist), TourTime=sum(tripTime)) %>%
  ungroup

#==============Tour by Tour purpose =============================
Tour_Purpose <- TourList %>%
  group_by(TourPurp)%>%
  summarise(Tour = n(), TourLength = mean(TourLength), TourTime = mean(TourTime))

Temp <- data.frame(Tour_Purpose)
rownames(Temp) <- Temp$TourPurp
Temp <- Temp[ ,c(-1)]
Temp <-rbind(Temp, c(sum(Tour_Purpose$Tour), mean(TourList$TourLength), mean(TourList$TourTime)))
row.names(Temp)[nrow(Temp)] <- "All_Purpose"

Temp_pct <- Temp
Temp_pct$Tour <- Temp_pct$Tour/sum(Tour_Purpose$Tour)

Temp0 <- Temp[, c("Tour", "TourLength")]
Temp0$Tour_pct <- Temp_pct$Tour
Temp0 <- Temp0[, c("Tour", "Tour_pct", "TourLength")]

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=70, colIndex=1, "Tour by Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=71, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=71, colIndex=1, "Tour Purpose", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=70, colIndex=1, "Tour by Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=71, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=71, colIndex=1, "Tour Purpose", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_Trip_Tour, sheet = sheets$`03 Tour & Length`, 
              startRow=4, startColumn=6, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`03 Tour & Length`, rowIndex=4, colIndex=6, "Tour Purpose", "T_Cell_1")

#==============Tour by Tour purpose =============================#


#==============Tour by Person Type =============================
Temp <- TourList %>%
  group_by(PBptype) %>%
  summarise(Tour = n(), TourLength = mean(TourLength), TourTime = mean(TourTime))

Temp0 <- pMod %>%
  left_join(PType_map, by="personType") %>%
  group_by(PBptype) %>% summarise(Persons=n()) %>%
  left_join(Temp, by="PBptype") %>%
  mutate(Tour_per_person= Tour/Persons) %>%
  data.frame

rownames(Temp0) <- Temp0$PBptype
Temp0 <- Temp0[,-1]
Temp <- rbind(Temp0, c(sum(Temp0$Persons), sum(Temp0$Tour), 
              mean(TourList$TourLength), mean(TourList$TourTime),
              sum(Temp0$Tour)/sum(Temp0$Persons)))

rownames(Temp)[nrow(Temp)] <- "Total"

Temp_pct <- Temp
Temp_pct$Persons <- Temp_pct$Persons/sum(Temp0$Persons)
Temp_pct$Tour <- Temp_pct$Tour/sum(Temp0$Tour)

Temp0 <- Temp[,c("Persons", "Tour", "Tour_per_person")]

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=89, colIndex=1, "Tour by PB Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=90, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=90, colIndex=1, "PBptype", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=89, colIndex=1, "Tour by PB Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=90, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=90, colIndex=1, "PBptype", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_Trip_Tour, sheet = sheets$`03 Tour & Length`, 
              startRow=28, startColumn=6, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`03 Tour & Length`, rowIndex=28, colIndex=6, "PBptype", "T_Cell_1")

#==============Tour by Person Type =============================#


#==============Tour by Person Type and tour purpose =============================
Temp <- reshape(aggregate(persTourNum~TourPurp+PBptype, data = TourList, FUN = length), 
                idvar = "TourPurp", timevar = "PBptype", direction = "wide")

rownames(Temp) <- Temp$TourPurp
Temp <- Temp[order(Temp$TourPurp), ]
Temp <- Temp[,-1]
Temp[is.na(Temp)] <- 0
colnames(Temp) <- gsub("persTourNum.", "", colnames(Temp))

Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, All_Purpose=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$All_Purpose
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)



xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=101, colIndex=1, "Tour by PB Person Type and Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=102, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=102, colIndex=1, "PBptype", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=101, colIndex=1, "Tour by PB Person Type and Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=102, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=102, colIndex=1, "PBptype", "T_Cell_1")
#==============Tour by Person Type and tour purpose ================================#

#==============Tour Length by Person Type and tour purpose =============================
Temp <- reshape(aggregate(TourLength~TourPurp+PBptype, data = TourList, FUN = mean), 
                idvar = "TourPurp", timevar = "PBptype", direction = "wide")

rownames(Temp) <- Temp$TourPurp
Temp <- Temp[order(Temp$TourPurp),-1]
colnames(Temp) <- gsub("TourLength.", "", colnames(Temp))

TourLength_Ptype <- aggregate(TourLength~PBptype, data = TourList, FUN = mean)
TourLength_tourp <- aggregate(TourLength~TourPurp, data = TourList, FUN = mean)
 
Temp <- cbind(Temp, Total=TourLength_tourp$TourLength)
Temp <- rbind(Temp, All_Purpose=c(TourLength_Ptype$TourLength, mean(TourList$TourLength)))

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=120, colIndex=1, "Tour Length by PB Person Type and Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=121, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=121, colIndex=1, "PBptype", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=120, colIndex=1, "Tour Length by PB Person Type", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=121, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=121, colIndex=1, "PBptype", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheets$`03 Tour & Length`, 
              startRow=44, startColumn=11, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`03 Tour & Length`, rowIndex=44, colIndex=11, "PBptype", "T_Cell_1")


#==============Tour Length by Person Type and tour purpose =============================#

#==============Tour Length Distribution by Tour Purpose =====================
TourList <- TourList %>%
  mutate(TourPurp2=ifelse(TourPurp %in% c("15JT_shop","16JT_main","17JT_eat","18JT_visit","19JT_discr"), "JonitNM
", ifelse(TourPurp %in% c("10Ind_shop", "111nd_main","121nd_eat","131nd_visi","141nd_disc"), "IndNM", TourPurp))) %>%
  mutate(TourDist_Bin=cut(TourLength, c(0:500), labels = c(0:499)+0.5))

Tour_TLD <- reshape(aggregate(persTourNum~TourDist_Bin+TourPurp2, data = TourList, FUN = length), 
                       idvar = "TourDist_Bin", timevar = "TourPurp2", direction = "wide")
colnames(Tour_TLD) <- gsub("persTourNum.", "", colnames(Tour_TLD))
Temp <- data.frame(levels(Tour_TLD$TourDist_Bin))
names(Temp) <- "TourLength"
Temp$TourLength <- as.character(Temp$TourLength)
Tour_TLD$TourDist_Bin <- as.character(Tour_TLD$TourDist_Bin)
Temp <- left_join(Temp, Tour_TLD, by=c("TourLength"="TourDist_Bin"))
Temp[is.na(Temp)] <- 0
rownames(Temp) <- Temp$TourLength
Temp <- Temp[,c(-1)]
Temp <- cbind(Temp, All_Purpose=rowSums(Temp))
Temp <- rbind(Temp, Total=colSums(Temp))

Temp_pct <- as.data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$Total
Temp_pct <- as.data.frame(t(Temp_pct))

xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=443, colIndex=1, "Tour Length Distribution by Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_Trip_Tour, sheet = sheet_num_data, 
              startRow=444, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_num_data, rowIndex=444, colIndex=1, "Tour Length\\ Tour Purpose", "T_Cell_1")

xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=443, colIndex=1, "Tour Length Distribution by Tour purpose", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_Trip_Tour, sheet = sheet_pct_data, 
              startRow=444, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheet_pct_data, rowIndex=444, colIndex=1, "Tour Length\\ Tour Purpose", "T_Cell_1")


xlsx.AddTable(as.data.frame(Temp_pct[-c(76:nrow(Temp_pct)),-ncol(Temp_pct)]), Target_Trip_Tour, sheet = sheets$`04 Tour TLD`, 
              startRow=4, startColumn=10, CellColor = "darkseagreen1")
xlsx.AddCell(Target_Trip_Tour, sheets$`04 Tour TLD`, rowIndex=4, colIndex=10, "Tour Length", "T_Cell_1")

#==============Tour Length Distribution by Tour Purpose =====================#

saveWorkbook(Target_Trip_Tour, file = paste0(Model_path, "Outputs/user/== VD 20181023 Length -Trip, Tour Length.xlsx"))

rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])