suppressMessages({
  library(readxl)
  library(writexl)
  library(plyr)
  library(dplyr)
  library(data.table)
  library(zoo)
  library(shiny)
  library(shinydashboard)
  library(shinydashboardPlus)
  library(shinyWidgets)
  library(htmlwidgets)
  library(lubridate)
  library(tcltk)
  library(tidyverse)
  library(plotly)
  library(knitr)
  library(kableExtra)
  library(leaflet)
  library(grid)
  library(gridExtra)
  library(eeptools)
  library(ggQC)
  #library(zipcode)
  library(utils)
  library(scales)
  library(chron)
  library(bupaR)
  library(shiny)
  library(DT)
  library(DiagrammeR)
  library(shinyalert)
  library(edeaR)
  library(processmapR)
  library(processmonitR)
  library(processanimateR)
  library(tidyr)
  library(lubridate)
  library(RColorBrewer)
  library(DiagrammeR)
  library(ggplot2)
  library(leaflet)
  library(readr)
  library(highcharter)
  library(ggforce) # for 'geom_arc_bar'
  library(packcircles) # for packed circle graph
  library(viridis)
  library(ggiraph)
  library(treemapify)
  library(treemap)
  library(broom)
  library(extrafont)
  library(tis) # for US holidays
  library(vroom)
  library(sjmisc)
  library(tools)
  library(here)
  library(shinyBS)
  library(shinyscreenshot)
  library(fasttime)
  library(shinycssloaders)
  library(feather)
  library(zipcodeR)
  library(formattable)
  library(shinyjs)
  library(janitor)
  library(patchwork)
  library(pryr)
  library(DBI)
  library(odbc)
})

memory.limit(size = 8000000)

process_data <- function(access_data){
  ### (3) Pre-process data ----------------------------------------------------------------------------------
  # SCheduling Data Pre-processing
  data.raw <- access_data # Assign scheduling Data

  # Dummy columns until they are added to Clarity table: SEX, FPA
  data.raw$SEX <- "Male"
  data.raw$VITALS_TAKEN_TM <- ""
  data.raw$Provider_Leave_DTTM <- ""
  
  # Data fields incldued for analysis 
  original.cols <- c("DEP_RPT_GRP_SEVENTEEN","DEPT_SPECIALTY_NAME","DEPARTMENT_NAME","PROV_NAME_WID",
                     "MRN","PAT_NAME","ZIP_CODE","SEX","BIRTH_DATE","FINCLASS",
                     "APPT_MADE_DTTM","APPT_DTTM","PRC_NAME","APPT_LENGTH","DERIVED_STATUS_DESC",
                     "APPT_CANC_DTTM", "CANCEL_REASON_NAME",
                     "SIGNIN_DTTM","PAGED_DTTM","CHECKIN_DTTM","ARVL_LIST_REMOVE_DTTM",
                     "ROOMED_DTTM","FIRST_ROOM_ASSIGN_DTTM","VITALS_TAKEN_TM",
                     "PHYS_ENTER_DTTM","Provider_Leave_DTTM",
                     "VISIT_END_DTTM","CHECKOUT_DTTM",
                     "TIME_IN_ROOM_MINUTES","CYCLE_TIME_MINUTES","VIS_NEW_TO_DEP_YN","LOS_NAME", "DEP_RPT_GRP_THIRTYONE", 
                     "APPT_ENTRY_USER_NAME_WID", "ACCESS_CENTER_SCHEDULED_YN", "VISIT_METHOD", "VISIT_PROV_STAFF_RESOURCE_C",
                     "PRIMARY_DX_CODE", "ENC_CLOSED_CHARGE_STATUS", "Y_ENC_COSIGN_TIME", "Y_ENC_CLOSE_TIME", "Y_ENC_OPEN_TIME", "NPI")
  
  # Subset raw data 
  data.subset <- data.raw[original.cols]
  
  # Rename data fields (columns) 
  new.cols <- c("Campus","Campus.Specialty","Department","Provider",
                "MRN","Patient.Name","Zip.Code","Sex","Birth.Date","Coverage",
                "Appt.Made.DTTM","Appt.DTTM","Appt.Type","Appt.Dur","Appt.Status",
                "Appt.Cancel.DTTM", "Cancel.Reason",
                "Signin.DTTM","Paged.DTTM","Checkin.DTTM","Arrival.remove.DTTM",
                "Roomin.DTTM","Room.assigned.DTTM","Vitals.DTTM",
                "Providerin_DTTM","Providerout_DTTM",
                "Visitend.DTTM","Checkout.DTTM",
                "Time.in.room","Cycle.time","New.PT","Class.PT","Cadence",
                "Appt.Source","Access.Center","Visit.Method","Resource",
                "PRIMARY_DX_CODE", "ENC_CLOSED_CHARGE_STATUS", "Y_ENC_COSIGN_TIME", "Y_ENC_CLOSE_TIME", "Y_ENC_OPEN_TIME", "NPI")
  
  colnames(data.subset) <- new.cols
  
  # Format Date and Time Columns
  dttm.cols <- c("Birth.Date","Appt.Made.DTTM","Appt.DTTM","Appt.Cancel.DTTM",
                 "Checkin.DTTM","Arrival.remove.DTTM","Roomin.DTTM","Room.assigned.DTTM",
                 "Vitals.DTTM","Providerin_DTTM","Providerout_DTTM",
                 "Visitend.DTTM","Checkout.DTTM")
  
  # Clean up department names (X_..._DEACTIVATED)
  data.subset <- data.subset %>%
    mutate(Department = ifelse(str_detect(Department, "DEACTIVATED"),
                               gsub('^.{2}|.{12}$', '', Department), 
                               ifelse(startsWith(Department,"X_"),
                                      gsub('^.{2}', '', Department), Department)))
  
  dttm <- function(x) {
    as.POSIXct(x,format="%Y-%m-%d %H:%M:%S",tz=Sys.timezone(),origin = "1970-01-01")
  }
  
  data.subset$Birth.Date <- dttm(data.subset$Birth.Date)
  data.subset$Appt.Made.DTTM <- dttm(data.subset$Appt.Made.DTTM)
  data.subset$Appt.DTTM <- dttm(data.subset$Appt.DTTM)
  data.subset$Appt.Cancel.DTTM <- dttm(data.subset$Appt.Cancel.DTTM)
  data.subset$Checkin.DTTM <- dttm(data.subset$Checkin.DTTM)
  data.subset$Arrival.remove.DTTM <- dttm(data.subset$Arrival.remove.DTTM)
  data.subset$Roomin.DTTM <- dttm(data.subset$Roomin.DTTM)
  data.subset$Room.assigned.DTTM <- dttm(data.subset$Room.assigned.DTTM)
  data.subset$Vitals.DTTM <- dttm(data.subset$Vitals.DTTM)
  data.subset$Providerin_DTTM <- dttm(data.subset$Providerin_DTTM)
  data.subset$Providerout_DTTM <- dttm(data.subset$Providerout_DTTM)
  data.subset$Visitend.DTTM <- dttm(data.subset$Visitend.DTTM)
  data.subset$Checkout.DTTM <- dttm(data.subset$Checkout.DTTM)
  
  # Remove Provider ID from Provider Name column
  data.subset$Provider <- trimws(gsub("\\[.*?\\]", "", data.subset$Provider))
  
  # New Patient Classification based on level of care ("LOS_NAME")
  data.subset$New.PT2 <- ifelse(is.na(data.subset$Class.PT), "",grepl("NEW", data.subset$Class.PT, fixed = TRUE))
  # New Patient Classification based on level of care ("LOS_NAME") and Visit New to Department (New.PT) TEMPORARY
  data.subset$New.PT3 <- ifelse(data.subset$New.PT2 == "", 
                                ifelse(data.subset$New.PT == "Y", TRUE, FALSE), data.subset$New.PT2)
  
  
  # Pre-process Appointment Source: access center, entry person, zocdoc, mychart, staywell
  data.subset$Appt.Source.New <- ifelse(data.subset$Access.Center == "Y", "Access Center","")
  data.subset$Appt.Source.New <- ifelse(data.subset$Appt.Source.New == "",
                                        ifelse(grepl("ZOCDOC", data.subset$Appt.Source, fixed = TRUE)==TRUE, "Zocdoc",
                                               ifelse(grepl("MYCHART", data.subset$Appt.Source, fixed = TRUE)==TRUE, "MyChart",
                                                      ifelse(grepl("STAYWELL", data.subset$Appt.Source, fixed = TRUE)==TRUE, "StayWell","Other"))),data.subset$Appt.Source.New)
  
  
  # Notify and remove duplicates in data 
  # data.duplicates <- data.subset %>% duplicated()
  # data.duplicates <- length(data.duplicates[data.duplicates == TRUE]) ## Count of duplicated records
  # data.subset.new <- data.subset %>% distinct() ## New data set with duplicates removed
  data.subset.new <- data.subset
  
  # Create additional columns for analysis 
  data.subset.new$Appt.DateYear <- as.Date(data.subset.new$Appt.DTTM, format="%Y-%m-%d") ## Create date-year column
  data.subset.new$Appt.MonthYear <- format(as.Date(data.subset.new$Appt.DTTM, format="%m/%d/%Y"), "%Y-%m") ## Create month - year column
  data.subset.new$Appt.Date <- format(as.Date(data.subset.new$Appt.DTTM, format="%m/%d/%Y"), "%m-%d") ## Create date column
  data.subset.new$Appt.Year <- format(as.Date(data.subset.new$Appt.DTTM, format="%m/%d/%Y"), "%Y") ## Create year column
  data.subset.new$Appt.Month <- format(as.Date(data.subset.new$Appt.DTTM, format="%m/%d/%Y"), "%b") ## Create month colunm
  data.subset.new$Appt.Quarter <- quarters(as.Date(data.subset.new$Appt.DTTM)) ## Create quarter column 
  data.subset.new$Appt.Week <- floor_date(as.Date(data.subset.new$Appt.DateYear, "%Y-%m-%d"), unit="week", week_start = 1) # Create week column
  data.subset.new$Appt.Day <- format(as.Date(data.subset.new$Appt.DTTM, format="%m/%d/%Y"), "%a") ## Create day of week colunm
  data.subset.new$Time <- format(as.POSIXct(as.ITime(data.subset.new$Appt.DTTM, format = "%H:%M")), "%H:%M") ## Create Time column
  data.subset.new$Appt.TM.Hr <- format(strptime(as.ITime(floor_date(data.subset.new$Appt.DTTM, "hour")), "%H:%M:%S"),'%H:%M') ## Appt time rounded by hour 
  # data.subset.new$Appt.TM.30m <- format(strptime(as.ITime(round_date(data.subset.new$Appt.DTTM, "30 minutes")), "%H:%M:%S"),'%H:%M') ## Appt time rounded by 30-min
  data.subset.new$Checkin.Hr <- format(strptime(as.ITime(round_date(data.subset.new$Checkin.DTTM, "hour")), "%H:%M:%S"),'%H:%M') ## Checkin time rounded by hour 
  # data.subset.new$Checkin.30m <- format(strptime(as.ITime(round_date(data.subset.new$Checkin.DTTM, "30 minutes")), "%H:%M:%S"),'%H:%M') ## Checkin time rounded by 30-min
  # data.subset.new$Lead.Days <- as.numeric((difftime(as.Date(data.subset.new$data.subset.new$Appt.DTTM, format="%Y-%m-%d"), as.Date(data.subset.new$data.subset.new$Appt.Cancel.DTTM, format="%Y-%m-%d"),  units = "days"))) ## Lead days for appt cancellation 
  data.subset.new$Lead.Days <- as.Date(data.subset.new$Appt.DTTM, format="%Y-%m-%d")-as.Date(data.subset.new$Appt.Cancel.DTTM, format="%Y-%m-%d") ## Lead days for appt cancellation
  data.subset.new$Wait.Time <- as.Date(data.subset.new$Appt.DTTM, format="%Y-%m-%d")-as.Date(data.subset.new$Appt.Made.DTTM, format="%Y-%m-%d")
  data.subset.new$uniqueId <- paste(data.subset.new$Department,data.subset.new$Provider,data.subset.new$MRN,data.subset.new$Appt.DTTM) ## Unique ID 
  data.subset.new$cycleTime <- as.numeric(round(difftime(data.subset.new$Visitend.DTTM,data.subset.new$Checkin.DTTM,units="mins"),1)) ## Checkin to Visitend (min)
  data.subset.new$checkinToRoomin <- as.numeric(round(difftime(data.subset.new$Roomin.DTTM,data.subset.new$Checkin.DTTM,units="mins"),1)) ## Checkin to Roomin (min)
  data.subset.new$providerinToOut <- as.numeric(round(difftime(data.subset.new$Providerout_DTTM,data.subset.new$Providerin_DTTM,units="mins"),1)) ## Provider in to out (min)
  data.subset.new$visitEndToCheckout <- as.numeric(round(difftime(data.subset.new$Checkout.DTTM,data.subset.new$Visitend.DTTM,units="mins"),1)) ## Visitend to Checkout (min)
  data.subset.new$Resource <- ifelse(data.subset.new$Resource == 1, "Provider", "Resource")
  
  ## Identify US Holidays in Data 
  hld <- holidaysBetween(min(data.subset.new$Appt.DTTM, na.rm=TRUE), max(data.subset.new$Appt.DTTM, na.rm=TRUE))
  holid <- as.Date(as.character(hld), format = "%Y%m%d")
  names(holid) <- names(hld)
  holid <- as.data.frame(holid)
  holid <- cbind(names(hld), holid)
  rownames(holid) <- NULL
  colnames(holid) <- c("holiday","date")
  holid$holiday <- as.character(holid$holiday)
  holid$date <- as.character(holid$date)
  holid$date <- as.Date(holid$date, format="%Y-%m-%d")
  
  data.subset.new$holiday <- holid$holiday[match(data.subset.new$Appt.DateYear, holid$date)]
  
  # Pre-processed Scheduling Dataframe
  data.subset.new <- as.data.frame(data.subset.new)
  
  reuturn_list <- list(data.subset.new,holid)
  return(reuturn_list)
}



con <- dbConnect(odbc(), Driver = "Oracle",
                 Host = "msx01-scan.mountsinai.org",
                 Port = 1521,
                 SVC = "PRD_MSX_TAF.msnyuhealth.org",
                 UID = "villea04",
                 PWD = "villea04123$"
)


access_date_1 <- "2021-01-01"
access_date_2 <- "2021-08-31"

access_sql <- paste0("SELECT DEP_RPT_GRP_SEVENTEEN,DEPT_SPECIALTY_NAME,DEPARTMENT_NAME,PROV_NAME_WID,DEPARTMENT_ID,REFERRING_PROV_NAME_WID,
                     MRN,PAT_NAME,ZIP_CODE,BIRTH_DATE,FINCLASS,
                     APPT_MADE_DTTM,APPT_DTTM,PRC_NAME,APPT_LENGTH,DERIVED_STATUS_DESC,
                     APPT_CANC_DTTM, CANCEL_REASON_NAME,
                     SIGNIN_DTTM,PAGED_DTTM,CHECKIN_DTTM,ARVL_LIST_REMOVE_DTTM,
                     ROOMED_DTTM,FIRST_ROOM_ASSIGN_DTTM,
                     PHYS_ENTER_DTTM,VISIT_END_DTTM,CHECKOUT_DTTM,
                     TIME_IN_ROOM_MINUTES,CYCLE_TIME_MINUTES,VIS_NEW_TO_DEP_YN,LOS_NAME,LOS_CODE,
                     DEP_RPT_GRP_THIRTYONE,APPT_ENTRY_USER_NAME_WID, 
                     ACCESS_CENTER_SCHEDULED_YN, VISIT_METHOD, VISIT_PROV_STAFF_RESOURCE_C,
                     PRIMARY_DX_CODE,ENC_CLOSED_CHARGE_STATUS,Y_ENC_COSIGN_TIME,Y_ENC_CLOSE_TIME,Y_ENC_OPEN_TIME, NPI
FROM CRREPORT_REP.MV_DM_PATIENT_ACCESS
WHERE CONTACT_DATE BETWEEN TO_DATE('", access_date_1,  "00:00:00', 'YYYY-MM-DD HH24:MI:SS')
				AND TO_DATE('", access_date_2, "23:59:59', 'YYYY-MM-DD HH24:MI:SS')")



access_raw <- dbGetQuery(con, access_sql)


process_data_run <- process_data(access_raw)
data.subset.new <- process_data_run[[1]]
holid <- process_data_run[[2]]



#Create Historical
max_date <- data.subset.new %>% filter(Appt.Status %in% c("Arrived"))
max_date <- max(max_date$Appt.DateYear) ## Or Today's Date
historical.data <- data.subset.new %>% filter(Appt.DateYear<= max_date) ## Filter out historical data only


#Utilization Data
max_date_all <- max(historical.data$Appt.DateYear) - 365
all.data <- historical.data #%>% filter(Appt.DTTM >= max_date_all) ## All data: Arrived, No Show, Canceled, Bumped, Rescheduled
arrived.data <- all.data %>% filter(Appt.Status %in% c("Arrived")) ## Arrived data: Arrived
canceled.bumped.rescheduled.data <- all.data %>% filter(Appt.Status %in% c("Canceled","Bumped","Rescheduled")) ## Canceled data: canceled appointments only
sameDay <- canceled.bumped.rescheduled.data %>% filter(Lead.Days == 0) # Same day canceled, rescheduled, bumped appts
noShow.data <- all.data %>% filter(Appt.Status %in% c("No Show")) ## Arrived + No Show data: Arrived and No Show
noShow.data <- rbind(noShow.data,sameDay) # No Shows + Same day canceled, bumped, rescheduled
arrivedNoShow.data <- rbind(arrived.data,noShow.data) ## Arrived + No Show data: Arrived and No Show


# Filter utilization data in last 60 days
max_date_util <- max(arrivedNoShow.data$Appt.DateYear) - 60
scheduled.data <- arrivedNoShow.data #%>% filter(Appt.DTTM >= max_date_util) ## All appts scheduled

# Function for formatting date and time by hour
system_date <- function(time){
  result <- as.POSIXct(paste0(as.character(Sys.Date())," ",time), format="%Y-%m-%d %H:%M:%S")
  return(result)
}

util.function <- function(time, df){
  result <- ifelse(system_date(time) %within% df$time.interval == TRUE,
                   ifelse(difftime(df$Appt.End.Time, system_date(time), units = "mins") >= 60, 60,
                          as.numeric(difftime(df$Appt.End.Time, system_date(time), units = "mins"))),
                   ifelse(floor_date(df$Appt.Start.Time, "hour") == system_date(time),
                          ifelse(floor_date(df$Appt.End.Time, "hour") == system_date(time),
                                 difftime(df$Appt.End.Time, df$Appt.Start.Time, units = "mins"),
                                 difftime(system_date(time) + 60*60, df$Appt.Start.Time, units = "mins")), 0))
  return(result)
}


# # Pre-process Utilization by Hour based on Scheduled Appointment Times --------------------------------------------------
data.hour.scheduled <- scheduled.data %>% filter(Appt.Status == "Arrived")
data.hour.scheduled$actual.visit.dur <- data.hour.scheduled$Appt.Dur * 1.2

data.hour.scheduled$Appt.Start <- as.POSIXct(data.hour.scheduled$Appt.DTTM, format = "%H:%M")
data.hour.scheduled$Appt.End <- as.POSIXct(data.hour.scheduled$Appt.Start + data.hour.scheduled$Appt.Dur*60, format = "%H:%M")

data.hour.scheduled$Appt.Start.Time <- as.POSIXct(paste0(Sys.Date()," ", format(data.hour.scheduled$Appt.Start, format="%H:%M:%S")))
data.hour.scheduled$Appt.End.Time <- as.POSIXct(paste0(Sys.Date()," ", format(data.hour.scheduled$Appt.End, format="%H:%M:%S")))


data.hour.scheduled$time.interval <- interval(data.hour.scheduled$Appt.Start.Time, data.hour.scheduled$Appt.End.Time)

# Excluding visits without Roomin or Visit End Tines
data.hour.scheduled$`00:00` <- util.function("00:00:00", data.hour.scheduled)
data.hour.scheduled$`01:00` <- util.function("01:00:00", data.hour.scheduled)
data.hour.scheduled$`02:00` <- util.function("02:00:00", data.hour.scheduled)
data.hour.scheduled$`03:00` <- util.function("03:00:00", data.hour.scheduled)
data.hour.scheduled$`04:00` <- util.function("04:00:00", data.hour.scheduled)
data.hour.scheduled$`05:00` <- util.function("05:00:00", data.hour.scheduled)
data.hour.scheduled$`06:00` <- util.function("06:00:00", data.hour.scheduled)
data.hour.scheduled$`07:00` <- util.function("07:00:00", data.hour.scheduled)
data.hour.scheduled$`08:00` <- util.function("08:00:00", data.hour.scheduled)
data.hour.scheduled$`09:00` <- util.function("09:00:00", data.hour.scheduled)
data.hour.scheduled$`10:00` <- util.function("10:00:00", data.hour.scheduled)
data.hour.scheduled$`11:00` <- util.function("11:00:00", data.hour.scheduled)
data.hour.scheduled$`12:00` <- util.function("12:00:00", data.hour.scheduled)
data.hour.scheduled$`13:00` <- util.function("13:00:00", data.hour.scheduled)
data.hour.scheduled$`14:00` <- util.function("14:00:00", data.hour.scheduled)
data.hour.scheduled$`15:00` <- util.function("15:00:00", data.hour.scheduled)
data.hour.scheduled$`16:00` <- util.function("16:00:00", data.hour.scheduled)
data.hour.scheduled$`17:00` <- util.function("17:00:00", data.hour.scheduled)
data.hour.scheduled$`18:00` <- util.function("18:00:00", data.hour.scheduled)
data.hour.scheduled$`19:00` <- util.function("19:00:00", data.hour.scheduled)
data.hour.scheduled$`20:00` <- util.function("20:00:00", data.hour.scheduled)
data.hour.scheduled$`21:00` <- util.function("21:00:00", data.hour.scheduled)
data.hour.scheduled$`22:00` <- util.function("22:00:00", data.hour.scheduled)
data.hour.scheduled$`23:00` <- util.function("23:00:00", data.hour.scheduled)

# Data Validation
# colnames(data.hour.scheduled[89])
data.hour.scheduled$sum <- rowSums(data.hour.scheduled[,which(colnames(data.hour.scheduled)=="00:00"):which(colnames(data.hour.scheduled)=="23:00")])
# data.hour.scheduled$sum <- rowSums(data.hour.scheduled [,66:89])
data.hour.scheduled$actual <- as.numeric(difftime(data.hour.scheduled$Appt.End.Time, data.hour.scheduled$Appt.Start.Time, units = "mins"))
data.hour.scheduled$comparison <- ifelse(data.hour.scheduled$sum ==data.hour.scheduled$actual, 0, 1)
data.hour.scheduled <- data.hour.scheduled %>% filter(comparison == 0)


# Pre-process Utilization by Hour based on Actual Room in to Visit End Times ---------------------------------------------------
data.hour.arrived.all <- scheduled.data %>% filter(Appt.Status == "Arrived")
data.hour.arrived.all$actual.visit.dur <- round(difftime(data.hour.arrived.all$Visitend.DTTM, data.hour.arrived.all$Roomin.DTTM, units = "mins"))

########### Analysis of % of visits with actual visit start and end times ############

data.hour.arrived.all$Appt.Start <- format(strptime(as.ITime(data.hour.arrived.all$Roomin.DTTM), "%H:%M:%S"),'%H:%M:%S')
data.hour.arrived.all$Appt.Start <- as.POSIXct(data.hour.arrived.all$Appt.Start, format = "%H:%M")
data.hour.arrived.all$Appt.End <- as.POSIXct(data.hour.arrived.all$Appt.Start + data.hour.arrived.all$actual.visit.dur, format = "%H:%M")

data.hour.arrived.all$Appt.Start.Time <- data.hour.arrived.all$Appt.Start
data.hour.arrived.all$Appt.End.Time <- data.hour.arrived.all$Appt.End

data.hour.arrived.all$time.interval <- interval(data.hour.arrived.all$Appt.Start.Time, data.hour.arrived.all$Appt.End.Time)

data.hour.arrived <- data.hour.arrived.all
# Excluding visits without Roomin or Visit End Tines
data.hour.arrived$`00:00` <- util.function("00:00:00", data.hour.arrived)
data.hour.arrived$`01:00` <- util.function("01:00:00", data.hour.arrived)
data.hour.arrived$`02:00` <- util.function("02:00:00", data.hour.arrived)
data.hour.arrived$`03:00` <- util.function("03:00:00", data.hour.arrived)
data.hour.arrived$`04:00` <- util.function("04:00:00", data.hour.arrived)
data.hour.arrived$`05:00` <- util.function("05:00:00", data.hour.arrived)
data.hour.arrived$`06:00` <- util.function("06:00:00", data.hour.arrived)
data.hour.arrived$`07:00` <- util.function("07:00:00", data.hour.arrived)
data.hour.arrived$`08:00` <- util.function("08:00:00", data.hour.arrived)
data.hour.arrived$`09:00` <- util.function("09:00:00", data.hour.arrived)
data.hour.arrived$`10:00` <- util.function("10:00:00", data.hour.arrived)
data.hour.arrived$`11:00` <- util.function("11:00:00", data.hour.arrived)
data.hour.arrived$`12:00` <- util.function("12:00:00", data.hour.arrived)
data.hour.arrived$`13:00` <- util.function("13:00:00", data.hour.arrived)
data.hour.arrived$`14:00` <- util.function("14:00:00", data.hour.arrived)
data.hour.arrived$`15:00` <- util.function("15:00:00", data.hour.arrived)
data.hour.arrived$`16:00` <- util.function("16:00:00", data.hour.arrived)
data.hour.arrived$`17:00` <- util.function("17:00:00", data.hour.arrived)
data.hour.arrived$`18:00` <- util.function("18:00:00", data.hour.arrived)
data.hour.arrived$`19:00` <- util.function("19:00:00", data.hour.arrived)
data.hour.arrived$`20:00` <- util.function("20:00:00", data.hour.arrived)
data.hour.arrived$`21:00` <- util.function("21:00:00", data.hour.arrived)
data.hour.arrived$`22:00` <- util.function("22:00:00", data.hour.arrived)
data.hour.arrived$`23:00` <- util.function("23:00:00", data.hour.arrived)

# Data Validation
# colnames(data.hour.arrived[89])
data.hour.arrived$sum <- rowSums(data.hour.arrived[,which(colnames(data.hour.arrived)=="00:00"):which(colnames(data.hour.arrived)=="23:00")])
data.hour.arrived$actual <- as.numeric(difftime(data.hour.arrived$Appt.End.Time, data.hour.arrived$Appt.Start.Time, units = "mins"))
data.hour.arrived$comparison <- ifelse(data.hour.arrived$sum == data.hour.arrived$actual, 0, 1)
data.hour.arrived$comparison[is.na(data.hour.arrived$comparison)] <- "No Data"
data.hour.arrived <- data.hour.arrived %>% filter(comparison != 1)

# Combine Utilization Data
data.hour.scheduled$util.type <- "scheduled"
data.hour.arrived$util.type <- "actual"

timeOptionsHr_filter <- c("07:00","08:00","09:00",
                          "10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00","18:00","19:00",
                          "20:00") ## Time Range by Hour Filter

utilization.data <- rbind(data.hour.scheduled, data.hour.arrived)
utilization.data <- utilization.data %>%
  select(Campus, Campus.Specialty, Department, Resource, Provider,
         Visit.Method, Appt.Type, Appt.Status,
         Appt.DateYear, Appt.MonthYear, Appt.Year, Appt.Week, Appt.Day, Appt.TM.Hr, holiday, util.type,
         timeOptionsHr_filter, sum, comparison, NPI)


### Population Data

zipcode_ref <- read_csv(here::here("Data/Oncology System Data - Zip Code Groupings 4.13.2021.csv"))
zipcode_ref <- zipcode_ref[1:(length(zipcode_ref)-7)]
zipcode_ref$`Zip Code Layer: A`[which(zipcode_ref$`Zip Code Layer: A` == "Long island")] <- "Long Island"

zipcode <- as.data.frame(read_feather("Data/zipcode.feather"))


population.data <- arrived.data
population.data$new_zip <- normalize_zip(population.data$Zip.Code)
population.data <- merge(population.data, zipcode_ref, by.x="new_zip", by.y="Zip Code", all.x = TRUE)

population.data <- merge(population.data, zipcode, by.x="new_zip", by.y="zip", all.x = TRUE)

population.data$`Zip Code Layer: A`[(is.na(population.data$`Zip Code Layer: A`) &
                                       (!is.na(population.data$state) | population.data$state != "NY"))] <- "Out of NYS"
population.data <- population.data %>%
  mutate(`Zip Code Layer: B` = ifelse(`Zip Code Layer: A` == "Out of NYS" & is.na(`Zip Code Layer: B`),
                                      ifelse(state == "NJ", "New Jersey",
                                             ifelse(state == "CT", "Connecticut",
                                                    ifelse(state == "FL", "Florida",
                                                           ifelse(state == "PA", "Pennsylvania", "Other")))), `Zip Code Layer: B`))


population.data_filtered <- population.data %>% filter(!is.na(`Zip Code Layer: A`))



