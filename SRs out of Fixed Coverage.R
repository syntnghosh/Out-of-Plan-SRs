#load libraries
library(readxl)
library(openxlsx)

#read data
#setwd("C:/Users/saghosh/Desktop/R Test/MBO_1/")

SR_History <- read_excel("Fixed_Plan_Out_of_Coverage_SRs_dataset.xlsx",
                   sheet = "SR Details")

SWM <- read_excel("Fixed_Plan_Out_of_Coverage_SRs_dataset.xlsx",
                  sheet = "SWM")


#backup data
backupSRHistory<-SR_History
backupSWM<-SWM
#-----------------------------------------------------------------------------------------------------------------------------------------
#
#start - re map SWM and SR History
SR_History<-backupSRHistory
SWM<-backupSWM


#rename columns
colnames(SWM)<-c("SWMSiteNum","SWMPMGroup","SWMPMName","SWMSPNum","SWMSPName","SWMPlanStartDate","SWMPlanEndDate")

colnames(SR_History)<-c("SRNum","SRSiteNum","SRSiteState","SRWorkDate","SRPMGroup","SRShortDescription","SRInvoiceStatus","SRSubstatus",
                        "SRSPNum","SRSPName","SRAPAmount")

#remove SEIbilling from SWM-when no fixed PM in place
#SWM<-SWM[SWM$SWMSPName!="SEI_BILLING",]

#create key to join
SWM$key<-paste(SWM$SWMPMGroup,SWM$SWMSiteNum)
SR_History$key<-paste(SR_History$SRPMGroup,SR_History$SRSiteNum)

#join SWM with SR history
combineddataset<- merge(SR_History,SWM,by="key",all.x=TRUE)

#remove entries that do not fall within SWM date range~ removes wrong joins and SRs completed when there was no fixed PM
combineddataset<-combineddataset[which(combineddataset$SRWorkDate < combineddataset$SWMPlanEndDate),]
combineddataset<-combineddataset[which(combineddataset$SRWorkDate > combineddataset$SWMPlanStartDate),]



#create final data file
results<-combineddataset[combineddataset$SRSPNum!=combineddataset$SWMSPNum,]
results$key=NULL
results$SWMSiteID=NULL
results$SWMLOS=NULL


#write.xlsx(results, "C:/Users/saghosh/Desktop/R Test/MBO_1/outofplandata.xlsx") 
write.xlsx(results, "OutofPlanSrs.xlsx") 



#write.xlsx(combineddataset, "Combined.xlsx") 
