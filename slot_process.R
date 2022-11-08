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

process_data <- function(slot.data.raw){
  data.subset.new <- readRDS("/data/Ambulatory/Data/historical_data.rds")
  data.subset.new <- data.subset.new %>% select(Department,Campus.Specialty,Campus)
  holid <- read_feather("/data/Ambulatory/Data/holid.feather")
  ## Pre-processing for Slot Data -----------------------------
  # Replace NAs with 0 in minutes columns
  slot.data.raw[,4:19][is.na(slot.data.raw[,4:19])] <- 0
  
  
  # Clean up department names (X_..._DEACTIVATED)
  slot.data.raw <- slot.data.raw %>%
    mutate(DEPARTMENT_NAME = ifelse(str_detect(DEPARTMENT_NAME, "DEACTIVATED"),
                               gsub('^.{2}|.{12}$', '', DEPARTMENT_NAME), 
                               ifelse(startsWith(DEPARTMENT_NAME,"X_"),
                                      gsub('^.{2}', '', DEPARTMENT_NAME), DEPARTMENT_NAME)))
  
  # Dept Specialty Manual Mapping to slot data 
  dept_specialty <- unique(data.subset.new[,c("Department","Campus.Specialty")])
  slot.data.raw$DEPT_SPECIALTY_NAME <- dept_specialty$Campus.Specialty[match(slot.data.raw$DEPARTMENT_NAME,dept_specialty$Department)]
  
  
  # Crosswalk Campus to Site by Department Name
  # slot.data.raw$Campus_new <- site_ref$`Site`[match(slot.data.raw$DEPARTMENT_NAME,site_ref$`Department Name`)]
  dept_campus <- unique(data.subset.new[,c("Department","Campus")])
  slot.data.raw$Campus_new <- dept_campus$Campus[match(slot.data.raw$DEPARTMENT_NAME,dept_campus$Department)]
  slot.data.raw <- slot.data.raw %>% filter(!Campus_new == "NA") %>% filter(!Campus_new %in% c("Other","OTHER","EHS")) ## Exclude Mapped Sites: Other, OTHER, EHS
  
  # Data fields incldued for analysis
  original.cols.slots <- c("Campus_new",
                           "DEPT_SPECIALTY_NAME",
                           "DEPARTMENT_NAME","PROVIDER_NAME",
                           "SLOT_BEGIN_TIME","NUM_APTS_SCHEDULED","SLOT_LENGTH",
                           "AVAIL_MINUTES","BOOKED_MINUTES","ARRIVED_MINUTES","CANCELED_MINUTES","NOSHOW_MINUTES","LEFTWOBEINGSEEN_MINUTES",
                           "AVAIL_SLOTS","BOOKED_SLOTS","ARRIVED_SLOTS","CANCELED_SLOTS","NOSHOW_SLOTS","LEFTWOBEINGSEEN_SLOTS",
                           "ORG_REG_OPENINGS","ORG_OVBK_OPENINGS","PRIVATE_YN","DAY_UNAVAIL_YN","TIME_UNAVAIL_YN","DAY_HELD_YN","TIME_HELD_YN","OUTSIDE_TEMPLATE_YN","VISIT_PROV_STAFF_RESOURCE_C")
  
  # Subset raw slot usage data
  slot.data.subset <- slot.data.raw[original.cols.slots]
  
  # Rename data columns to match schduling data
  new.cols.slots <- c("Campus",
                      "Campus.Specialty",
                      "Department","Provider",
                      "SLOT_BEGIN_TIME","NUM_APTS_SCHEDULED","SLOT_LENGTH",
                      "AVAIL_MINUTES","BOOKED_MINUTES","ARRIVED_MINUTES","CANCELED_MINUTES","NOSHOW_MINUTES","LEFTWOBEINGSEEN_MINUTES",
                      "AVAIL_SLOTS","BOOKED_SLOTS","ARRIVED_SLOTS","CANCELED_SLOTS","NOSHOW_SLOTS","LEFTWOBEINGSEEN_SLOTS",
                      "ORG_REG_OPENINGS","ORG_OVBK_OPENINGS","PRIVATE_YN","DAY_UNAVAIL_YN","TIME_UNAVAIL_YN","DAY_HELD_YN","TIME_HELD_YN","OUTSIDE_TEMPLATE_YN","Resource")
  
  colnames(slot.data.subset) <- new.cols.slots
  
  # Create additional columns for Slot Data
  slot.data.subset$BOOKED_MINUTES <- slot.data.subset$BOOKED_MINUTES + slot.data.subset$CANCELED_MINUTES # Booked + Canceled Minutes 
  slot.data.subset$Appt.DTTM <- as.POSIXct(slot.data.subset$SLOT_BEGIN_TIME,format="%Y-%m-%d %H:%M:%S",tz=Sys.timezone(),origin = "1970-01-01")
  slot.data.subset$Appt.DateYear <- as.Date(slot.data.subset$SLOT_BEGIN_TIME, format="%Y-%m-%d") ## Create day of week colunm
  slot.data.subset$Appt.MonthYear <- format(as.Date(slot.data.subset$SLOT_BEGIN_TIME, format="%Y-%m-%d %H:%M:%S"), "%Y-%m") ## Create month - year column
  slot.data.subset$Appt.Year <- format(as.Date(slot.data.subset$Appt.DTTM, format="%Y-%m-%d %H:%M:%S"), "%Y") ## Create year column
  slot.data.subset$Appt.Month <- format(as.Date(slot.data.subset$Appt.DTTM, format="%Y-%m-%d %H:%M:%S"), "%b") ## Create month colunm
  slot.data.subset$Appt.Quarter <- quarters(as.Date(slot.data.subset$Appt.DTTM)) ## Create quarter column 
  slot.data.subset$Appt.Week <- floor_date(as.Date(slot.data.subset$Appt.DateYear, "%Y-%m-%d"), unit="week", week_start = 1) # Create week column
  slot.data.subset$Appt.Day <- format(as.Date(slot.data.subset$SLOT_BEGIN_TIME, format="%Y-%m-%d %H:%M:%S"), "%a") ## Create day of week colunm
  slot.data.subset$Time <- format(as.POSIXct(as.ITime(slot.data.subset$SLOT_BEGIN_TIME, format = "%H:%M"),origin = "1970-01-01"), "%H:%M") ## Create Slot Time column
  slot.data.subset$Appt.TM.Hr <- format(strptime(as.ITime(round_date(slot.data.subset$Appt.DTTM, "hour")), "%H:%M:%S"),'%H:%M') ## Appt time rounded by hour
  
  slot.data.subset$holiday <- holid$holiday[match(slot.data.subset$Appt.DateYear, holid$date)] ## Identify US Holidays in Data
  slot.data.subset$Visit.Method <- "IN PERSON"
  slot.data.subset$Resource <- ifelse(slot.data.subset$Resource == 1, "Provider", "Resource")
  
  slot.data.subset <- slot.data.subset %>%
    mutate(siteSpecialty = paste0(Campus, " - ", Campus.Specialty)) %>%
    group_by(Campus, siteSpecialty, Campus.Specialty, Department, Provider,
             Appt.DateYear, Appt.MonthYear, Appt.Year, Appt.Week, Appt.Day, Appt.TM.Hr,
             Resource, Visit.Method, holiday) %>%
    dplyr::summarise(`Available Hours` = sum(AVAIL_MINUTES)/60,
                     `Booked Hours` = sum(BOOKED_MINUTES)/60,
                     `Arrived Hours` = sum(ARRIVED_MINUTES)/60,
                     `Canceled Hours` = sum(CANCELED_MINUTES)/60,
                     `No Show Hours` = sum(NOSHOW_MINUTES , LEFTWOBEINGSEEN_MINUTES)/60)
}

con <- dbConnect(odbc(), Driver = "Oracle",
                 Host = "msx01-scan.mountsinai.org",
                 Port = 1521,
                 SVC = "PRD_MSX_TAF.msnyuhealth.org",
                 UID = "villea04",
                 PWD = "villea04123$"
)

slot_date_1 <- "2022-01-01"
slot_date_2 <- Sys.Date()-1

slot_sql <- paste0("SELECT DEPARTMENT_NAME,PROVIDER_NAME,
	       SLOT_BEGIN_TIME,NUM_APTS_SCHEDULED,SLOT_LENGTH,
	       AVAIL_MINUTES,BOOKED_MINUTES,ARRIVED_MINUTES,
	       CANCELED_MINUTES,NOSHOW_MINUTES,LEFTWOBEINGSEEN_MINUTES,
	       AVAIL_SLOTS,BOOKED_SLOTS,ARRIVED_SLOTS,CANCELED_SLOTS,
	       NOSHOW_SLOTS,LEFTWOBEINGSEEN_SLOTS,ORG_REG_OPENINGS,
	       ORG_OVBK_OPENINGS,PRIVATE_YN,DAY_UNAVAIL_YN,TIME_UNAVAIL_YN,
	       DAY_HELD_YN,TIME_HELD_YN,OUTSIDE_TEMPLATE_YN,VISIT_PROV_STAFF_RESOURCE_C
	FROM CRREPORT_REP.Y_DM_BOOKED_FILLED_RATE
	WHERE SLOT_BEGIN_TIME BETWEEN TO_DATE('", slot_date_1, "00:00:00', 'YYYY-MM-DD HH24:MI:SS')
					AND TO_DATE('", slot_date_2, "23:59:59', 'YYYY-MM-DD HH24:MI:SS')")


slot_raw <- dbGetQuery(con, slot_sql)

slot.data.subset <- process_data(slot_raw)

#Create Slot
max_date_slot <- max(slot.data.subset$Appt.DateYear) - 455
slot.data.subset <- slot.data.subset %>% filter(Appt.DateYear >= max_date_slot)


