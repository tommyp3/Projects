############# TITLE: Weekly Ad Hoc Scrub
#############
############# WRITTEN BY: THOMAS PETERSON
############# CREATED ON: 01/31/2020
############# NOTES: 
############# DOCUMENTATION: 



####LOADING BASE R PACKAGES
source("T:/InterTeam/FEP-Ops-Metrics/99 Documentation/R Templates/UTIL/global_library.R")
source(str_c(gsub("Documents","Desktop",getwd()),"/R/Credentials/FDS_PROD.R"))


####LOAD ALL CLAIM NUMBERS FOR WEEKLY AD HOC
all_claims <- read.xlsx("C:/Users/r635223/Desktop/ID 4624 Unique Claims 07-14-2019 to 01-20-2020.xlsx")

names(all_claims)[1] <- "CLM_NUM"

# CREATE LIST OF ALL PREVIOUS REPORTS FROM SHARED FOLDER
all_reports <- list.files("T:/InterTeam/FEP-Ops-Metrics/2 Ad-Hoc Requests/1 Weekly Adhocs/ID 4624", pattern='*.xlsx')

#CONVERT LIST TO DATA FRAME
all_reports <- as.data.frame(all_reports)
names(all_reports)[1] <- "REPORT_NAME"

#SELECT ONLY WEEKLY REPORTS 
weekly_reports <- sqldf('SELECT REPORT_NAME
                            FROM all_reports
                            WHERE REPORT_NAME LIKE "ID 4624%"
                            AND REPORT_NAME NOT LIKE "%01-20-2020%"')

stage_claim <- data.frame(CLM_NUM = character())

for (i in 1:nrow(weekly_reports)){                                           
  
    tmp<-read.xlsx(str_c("T:/InterTeam/FEP-Ops-Metrics/2 Ad-Hoc Requests/1 Weekly Adhocs/ID 4624/",weekly_reports[i,1]),
                   sheet = "Unique Claims")
    
    names(tmp) <- c("CLM_NUM","TMP")
    
    tmp$TMP <- NULL
    
    stage_claim <- merge(stage_claim, tmp, all=T)
}


scrubbed_claims <- sqldf('SELECT a.CLM_NUM, s.CLM_NUM
                          FROM all_claims a
                          LEFT JOIN stage_claim s
                          ON a.CLM_NUM = s.CLM_NUM
                          WHERE s.CLM_NUM IS NULL')

#### PRINTING LIST OF CLAIMS TO DESKTOP

filename <- paste("C:/Users/r635223/Documents/4624_scrubbed_claims", ".xlsx", sep="")
write.xlsx(scrubbed_claims, file = filename, row.names = FALSE, na="")



