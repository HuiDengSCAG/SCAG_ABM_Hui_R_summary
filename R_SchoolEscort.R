#update the function for SchoolEscorting clean
cleanAndReadCT2_out <- function(fName){
  
  inpLines <- readLines(fName)
  headerLine <- paste0(gsub('.{2}$', '', unlist(str_split(inpLines[[1]],','))),collapse = ',') # Cleaning up the header row
  # adding a double quote before square bracket opening and and after a square bracket ending
  cleanted_fName <- gsub('.csv','_cleaned.csv',fName)
  writeLines(append(headerLine,unlist(lapply(2:length(inpLines), function(iLine) gsub("\\]",'\\]"',gsub("\\[",'"\\[',unlist(str_split(inpLines[[iLine]],',\\[\\['))[[1]]))))),
             cleanted_fName)
  return(read.csv(cleanted_fName) %>% data.table)
  
}

if(file.exists(paste0(ctrampOutPath,"/schoolEscorting_cleaned.csv"))){
  SchoolEscort <- data.frame(fread(paste0(ctrampOutPath,"/schoolEscorting_cleaned.csv")))
} else{
  SchoolEscort <- cleanAndReadCT2_out(paste0(ctrampOutPath,"/schoolEscorting.csv"))
}

PType_map <-  data.frame(personType=c(1:8),
                         PBptype= c("P1FTW","P2PTW", "P3Col", "P4NWadult", "P5NWold", 
                                    "P6_16_18", "P7_6_15", "P8_0_5"),stringsAsFactors = F)


cdapPerson <- read_csv(paste0(ctrampOutPath,"coordinatedDailyActivityPattern_person.csv"))
names(cdapPerson) <- gsub('.{2}$', '', names(cdapPerson)) # Remove last two characters from all column names
totChildrenWithMand <- cdapPerson %>% 
  filter(personType %in% c(6,7,8) & personDaps==1) %>% 
  left_join(PType_map, by=c("personType"="personType")) %>%
  group_by(PBptype) %>% 
  summarise(Total=n())



# escortBundleDirection: 1=Outbound, 2=Inbound
# escortBundleType: 1=Ride-sharing,2=Pure-escort

SchoolEscort <- SchoolEscort %>%
  mutate(recordID = 1:n()) %>%
  mutate(Direction=ifelse(escortBundleDirection==1,'Outbound','Inbound')) %>%
  mutate(Alternative=ifelse(escortBundleType==1,'Ride-sharing',
                            ifelse(escortBundleType==2,'Pure-escort','No-escort'))) %>%
  mutate(numChildrenBundle=str_count(escortBundleChildren, ",")+1,
         escortBundleChildPids=gsub("\\]",'',gsub("\\[",'',escortBundleChildPids))) %>%
  left_join(PType_map, by=c("personType"="personType"))

Student_SchoolEscort <- SchoolEscort %>%
  select(recordID,escortBundleChildPids,Direction,Alternative,numChildrenBundle) %>% 
  data.table

Student_SchoolEscort <- Student_SchoolEscort[rep(seq_len(nrow(Student_SchoolEscort)), Student_SchoolEscort$numChildrenBundle), ]

Student_SchoolEscort$n <- Student_SchoolEscort[,.(1:numChildrenBundle),by='recordID']$V1

Student_SchoolEscort <- Student_SchoolEscort %>% rowwise() %>% mutate(childPID = as.integer(unlist(str_split(escortBundleChildPids,','))[n]))

Student_SchoolEscort <- Student_SchoolEscort %>%
  left_join(pMod[,c("pid", "personType")], by=c("childPID"="pid")) %>%
  left_join(PType_map, by=c("personType"="personType"))


#==============load xlsx and define dataframe format========================
Target_SchoolEscort <- loadWorkbook(paste0(Target_path, 
                                      "= VT11 20180611  School escorting & schedule consolidation (5.2.4).xlsx"),
                               password=NULL)

sheet_num_data <- createSheet(Target_SchoolEscort, sheetName = "Number_Data")
sheet_pct_data <- createSheet(Target_SchoolEscort, sheetName = "pct_Data")
sheets <- getSheets(Target_SchoolEscort)
removeSheet(Target_SchoolEscort, sheetName=sheets[[4]])
removeSheet(Target_SchoolEscort, sheetName=sheets[[5]])
removeSheet(Target_SchoolEscort, sheetName=sheets[[6]])
sheets <- getSheets(Target_SchoolEscort)
#================School Escort Summary===================#


#================Student Outbound Escort Summary===================
Temp <- reshape(aggregate(recordID~Alternative+PBptype,
                  data = Student_SchoolEscort[which(Student_SchoolEscort$Direction=="Outbound"), ], FUN = length),
                idvar = "PBptype", timevar = "Alternative", direction = "wide")
names(Temp) <- gsub("recordID.", "", names(Temp))
Temp <- Temp %>%
  left_join(totChildrenWithMand, by=c("PBptype")) %>%
  mutate(`No-escort`=Total-`Ride-sharing`- `Pure-escort`)
  
Temp <- Temp[, c("PBptype", "Ride-sharing", "Pure-escort", "No-escort", "Total")]
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- rbind(Temp, Sum=colSums(Temp))

Temp_pct <- Temp/Temp$Total

xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=1, colIndex=1, "School Escort Summary", "Title_1")
xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=2, colIndex=1, "Student Outbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_SchoolEscort, sheet = sheet_num_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=3, colIndex=1, "Students", "T_Cell_1")


xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=1, colIndex=1, "School Escort Summary", "Title_1")
xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=2, colIndex=1, "Student Outbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheet_pct_data, 
              startRow=3, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=3, colIndex=1, "Students", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheets$`5.2.4`, 
              startRow=8, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheets$`5.2.4`, rowIndex=8, colIndex=13, "Students", "T_Cell_1")
#================Student Outbound Escort Summary===================#

#================Chauffer Outbound Escort Summary===================
Temp <- reshape(aggregate(numChildrenBundle~Alternative+PBptype,
                          data = SchoolEscort[which(SchoolEscort$Direction=="Outbound"), ], FUN = length),
                idvar = "PBptype", timevar = "Alternative", direction = "wide")
names(Temp) <- gsub("numChildrenBundle.", "", names(Temp))
Temp[is.na(Temp)] <- 0

Temp <- Temp[, c("PBptype", "Ride-sharing", "Pure-escort")]
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, Sum=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$Sum
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)
names(Temp_pct) <- names(Temp)

Temp0 <- Temp[c("Sum"),c("Ride-sharing", "Pure-escort")]/Temp[c("Sum"),c("Total")]
Temp0 <- rbind(Temp_pct[,c("Ride-sharing", "Pure-escort")], Sum2=Temp0)

xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=9, colIndex=1, "Chauffer Outbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_SchoolEscort, sheet = sheet_num_data, 
              startRow=10, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=10, colIndex=1, "Adults", "T_Cell_1")

xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=9, colIndex=1, "Chauffer Outbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheet_pct_data, 
              startRow=10, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=10, colIndex=1, "Adults", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_SchoolEscort, sheet = sheets$`5.2.4`, 
              startRow=15, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheets$`5.2.4`, rowIndex=15, colIndex=13, "Adults", "T_Cell_1")
#================Chauffer Outbound Escort Summary===================#

#================Student Inbound Escort Summary===================
Temp <- reshape(aggregate(recordID~Alternative+PBptype,
                          data = Student_SchoolEscort[which(Student_SchoolEscort$Direction=="Inbound"), ], FUN = length),
                idvar = "PBptype", timevar = "Alternative", direction = "wide")
names(Temp) <- gsub("recordID.", "", names(Temp))
Temp <- Temp %>%
  left_join(totChildrenWithMand, by=c("PBptype")) %>%
  mutate(`No-escort`=Total-`Ride-sharing`- `Pure-escort`)

Temp <- Temp[, c("PBptype", "Ride-sharing", "Pure-escort", "No-escort", "Total")]
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- rbind(Temp, Sum=colSums(Temp))

Temp_pct <- Temp/Temp$Total


xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=18, colIndex=1, "Student Inbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_SchoolEscort, sheet = sheet_num_data, 
              startRow=19, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=19, colIndex=1, "Students", "T_Cell_1")


xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=18, colIndex=1, "Student Inbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheet_pct_data, 
              startRow=19, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=19, colIndex=1, "Students", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheets$`5.2.4`, 
              startRow=27, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheets$`5.2.4`, rowIndex=27, colIndex=13, "Students", "T_Cell_1")
#================Student Inbound Escort Summary===================#

#================Chauffer Inbound Escort Summary===================
Temp <- reshape(aggregate(numChildrenBundle~Alternative+PBptype,
                          data = SchoolEscort[which(SchoolEscort$Direction=="Inbound"), ], FUN = length),
                idvar = "PBptype", timevar = "Alternative", direction = "wide")
names(Temp) <- gsub("numChildrenBundle.", "", names(Temp))
Temp[is.na(Temp)] <- 0

Temp <- Temp[, c("PBptype", "Ride-sharing", "Pure-escort")]
rownames(Temp) <- Temp$PBptype
Temp <- Temp[, -1]
Temp <- cbind(Temp, Total=rowSums(Temp))
Temp <- rbind(Temp, Sum=colSums(Temp))

Temp_pct <- data.frame(t(Temp))
Temp_pct <- Temp_pct/Temp_pct$Sum
Temp_pct <- data.frame(t(Temp_pct))
rownames(Temp_pct) <- rownames(Temp)
names(Temp_pct) <- names(Temp)

Temp0 <- Temp[c("Sum"),c("Ride-sharing", "Pure-escort")]/Temp[c("Sum"),c("Total")]
Temp0 <- rbind(Temp_pct[,c("Ride-sharing", "Pure-escort")], Sum2=Temp0)

xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=25, colIndex=1, "Chauffer Inbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp), Target_SchoolEscort, sheet = sheet_num_data, 
              startRow=26, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_num_data, rowIndex=26, colIndex=1, "Adults", "T_Cell_1")

xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=25, colIndex=1, "Chauffer Inbound Escort Summary", "Title_2")
xlsx.AddTable(as.data.frame(Temp_pct), Target_SchoolEscort, sheet = sheet_pct_data, 
              startRow=26, startColumn=1, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheet_pct_data, rowIndex=26, colIndex=1, "Adults", "T_Cell_1")

xlsx.AddTable(as.data.frame(Temp0), Target_SchoolEscort, sheet = sheets$`5.2.4`, 
              startRow=34, startColumn=13, CellColor = "darkseagreen1")
xlsx.AddCell(Target_SchoolEscort, sheets$`5.2.4`, rowIndex=34, colIndex=13, "Adults", "T_Cell_1")
#================Chauffer Inbound Escort Summary===================#


saveWorkbook(Target_SchoolEscort, file = paste0(Model_path, "Outputs/user/== VT11 20180611  School escorting & schedule consolidation (5.2.4).xlsx"))

rm(list=ls()[! ls() %in% c("ctrampOutPath","Model_path", "popOutPath", "Summary_path", "Target_path", 
                           "xlsx.AddCell", "xlsx.AddTable", "cleanAndReadCT2_out",
                           "hMod", "pMod", "TripList1")])