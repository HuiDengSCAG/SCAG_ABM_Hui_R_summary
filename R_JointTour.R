jointTourFrequency <- cleanAndReadCT2_out(paste0(ctrampOutPath,"jointTourFrequency.csv"))
jointTourFormation <- cleanAndReadCT2_out(paste0(ctrampOutPath,"jointTourFormation.csv"))
jointTourFormation <- jointTourFormation %>% filter(jointTourPurpose!=0)

jointTourFormation <- jointTourFormation %>%  
  mutate(jointTourActivityDists=gsub("\\]",'',gsub("\\[",'',jointTourActivityDists))) %>% 
  rowwise() %>% 
  mutate(numStops = length(unlist(str_split(jointTourActivityTazs,',')))-1,
    primActDistance = as.numeric(unlist(str_split(jointTourActivityDists,','))[jointTourPrimaryIndex+1])) %>% 
  mutate(
    numStopsOutbound = jointTourPrimaryIndex,
    numStopsInbound = numStops-jointTourPrimaryIndex,
    jointTodStartHr = binToHour(tourDepartHome),
    jointTodEndHr = binToHour(tourArriveHome),
    jointTodDurHr = ceiling((tourArriveHome-tourDepartHome+1)/4)
  ) %>% 
  data.table










Temp <- jointTourFrequency %>% filter(jointTourPurpose!=0) %>% group_by(hhid) %>% summarise(numJTours=n()) %>% 
  right_join(hMod %>% select(hhid,HHSIZE) %>% filter(HHSIZE>1),by='hhid') %>%
  mutate(numJTours=ifelse(is.na(numJTours),0,numJTours)) %>% 
  mutate(HHSIZE=ifelse(HHSIZE>5,5,HHSIZE)) %>% 
  group_by(HHSIZE,numJTours) %>% summarise(numHH=n()) %>%
  data.frame %>%
  reshape(idvar = "HHSIZE", timevar = "numJTours", direction = "wide")
  






jointTourPID <- jointTourFrequency %>% filter(jointTourPurpose!=0) %>% mutate(recordID=1:n()) %>% 
  select(recordID,hhid, partyType,jointTourPurpose,jointTourPartyPersTypes) %>% 
  mutate(numpersonJTour=unlist(lapply(jointTourPartyPersTypes,function(x) length(unlist(str_split(x,','))))),
         jointTourPartyPersTypes=gsub("\\]",'',gsub("\\[",'',jointTourPartyPersTypes))) %>% data.table

jointTourPID <- jointTourPID[,lapply(.SD,rep,numpersonJTour)]
jointTourPID$n <- jointTourPID[,.(1:numpersonJTour),by='recordID']$V1
jointTourPID <- jointTourPID %>% rowwise() %>% mutate(personType = as.integer(unlist(str_split(jointTourPartyPersTypes,','))[n])) %>% ungroup()
#jointTourPID <- jointTourPID %>% left_join(pMod %>% select(pid,personType,personTypeDetailed),by=c("childPID"="pid"))

m_summary2 <- jointTourPID %>% group_by(personType,partyType) %>% summarise(numTours = n()) %>%
  data.frame %>%
  reshape(idvar = "personType", timevar = "partyType", direction = "wide")