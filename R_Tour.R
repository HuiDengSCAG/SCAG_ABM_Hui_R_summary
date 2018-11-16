
TourPurp_Map <- read.csv(paste0(targetPath,"TourPurp.csv"), header=T, stringsAsFactors=FALSE)
trip_tour <- TripList %>%
  left_join(TourPurp_Map, by=c("mcTourType"="mcTourType", "tourPurpose"="tourPurpose")) %>%
  mutate(tripTime=tripArriveMinute-tripDepartMinute) %>%
  filter(mcTourType != 4) %>%
  group_by(hhid, pnum, persTourNum,tourPurpose) %>%
  summarise(TourLength = sum(tripDist), TourTime=sum(tripTime)) %>%
  ungroup%>%
  group_by(tourPurpose)%>%
  summarise(Tour = n(), TourLength = mean(TourLength), TourTime = mean(TourTime))





Tour_Purpose <- reshape(aggregate(tripId~mcTourType+tourPurpose, data = TripList, FUN = length), 
                           idvar = "mcTourType", timevar = "tourPurpose", direction = "wide")