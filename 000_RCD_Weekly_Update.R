### LOAD LIBRABRIES
library (readxl)
library(dplyr)
library (lubridate)
library(openxlsx)
library (RDCOMClient)
library(DIZtools)


##### LOAD LINKS ######
documents_folder <- "C://Users//h409422//OneDrive - Santander Office 365//Documents//"
forecast_folder <- "C://Users//h409422//OneDrive - Santander Office 365//"
Sharepoint_folder <- "C://Users//h409422//Santander Office 365//Sales - Quell_Dateien/"
RCD_folder <-"C://Users/h409422/Santander Office 365/HCBE Dashboard Retail - Retail Contract Details via Databricks/"
setwd(documents_folder)


teams_list <- sort( list.files(RCD_folder) ,  decreasing=TRUE)
teams_list<- teams_list[  ! grepl (  "KMD",  teams_list) ]
teams_list<- teams_list[  ! grepl (  "HMD",  teams_list) ]
teams_list<- teams_list[  ! grepl (  "KMD",  teams_list) ]
teams_list<- teams_list[  ! grepl (  "desktop",  teams_list) ]
teams_list<- teams_list[   grepl (  ".xlsx",  teams_list) ]
teams_list<- teams_list[   grepl (  "Retail",  teams_list) ]
teams_list<- teams_list[   grepl ( 2025,  teams_list) ]
teams_list<- teams_list[   grepl ( paste0 ( "_0" , month(today()-3 )),  teams_list) ]
RCD_file <-  teams_list [1]

##### READ THE RCD FILE 


df <-    readxl::read_excel(paste0(RCD_folder ,  RCD_file) ,
                            guess_max = 50000  ,
                            sheet = 1)

df <- as.data.frame(df)

print ("RCD was extracted properly")

df <- df %>%
  rename (     "Storno J/N" = "StornoJ/N") %>%
  
  
  rename (     "Fahrgestellnr." = "Fahrgestellnr") %>%
  
  rename (   "Rabattierfähiger Betrag netto"= "Rabattierfähiger Betrag Netto" ) %>%
  rename (    "Fin. Betrag brutto"  = "Fin Betrag brutto" ) %>%
  rename (    "Fin. Betrag netto" = "Fin Betrag netto" ) %>%
  rename (    "1. Rate brutto" = "1 Rate brutto" ) %>%
  rename (    "1. Rate netto" = "1 Rate netto" ) %>%
  rename (      "WKZ Hdl. netto"= "WKZ Hdl netto" ) %>%
  rename (      "Händlernr."= "Händler Nr") %>%  rename (      "Händler Nr."= "Händlernr")


### correct format for the dates 
df$Anfragedatum <- as_date(df$Anfragedatum)
df$Aktivierung <- as_date(df$Aktivierung)
df$Erstzulassung <- as_date(df$Erstzulassung)

### only for safety in case we have cells with NA then they are 0. 
df$`RSV - Prämie` [  is.empty( df$`RSV - Prämie` ) ]   <- 0
df$`GAP - Prämie` [  is.empty( df$`GAP - Prämie` ) ]   <- 0
df$`MOBI PLUS - Prämie` [  is.empty( df$`MOBI PLUS - Prämie` ) ]   <- 0
df$`EW - Prämie` [  is.empty( df$`EW - Prämie` ) ]   <- 0


df$`Rabattierter Kaufpreis brutto` [is.empty(df$`Rabattierter Kaufpreis brutto` )] <-0
df$`Rabattierter Kaufpreis netto` [is.empty(df$`Rabattierter Kaufpreis netto` )] <-0
df$`Anzahlung in EUR brutto` [is.empty(df$`Anzahlung in EUR brutto` )] <-0
df$`Anzahlung in EUR netto` [is.empty(df$`Anzahlung in EUR netto` )] <-0



### Fin b1 and Fin b2
df$Finb1 <-
  if_else(
    df$Produktgattung == "Finanzierung",
    df$`Rabattierter Kaufpreis brutto` - df$`Anzahlung in EUR brutto`,
    df$`Rabattierter Kaufpreis netto` - df$`Anzahlung in EUR netto`
  )


df$Finb2 <-  if_else(
  df$Produktgattung == "Finanzierung",
  df$`Rabattierter Kaufpreis brutto` - df$`Anzahlung in EUR brutto` +     df$`RSV - Prämie` + df$`GAP - Prämie` + df$`MOBI PLUS - Prämie`  + df$`EW - Prämie`,
  df$`Rabattierter Kaufpreis netto` - df$`Anzahlung in EUR netto`  +    df$`RSV - Prämie` + df$`GAP - Prämie` + df$`MOBI PLUS - Prämie` +  df$`EW - Prämie`,
)



#### stop here
#####  in case of empty dealers 
Zuordnung <-   readxl::read_excel(    paste0(   Sharepoint_folder ,  "00_Zuordnung_Hdl_Nr_Vertragsnr.xlsx") ,        sheet= "Sheet1")
Zuordnung$Vertragsnummer <- Zuordnung$`Angebots-nummer`

A<-left_join(   df [  is.empty(df$Händlernr.)  , ]  , Zuordnung[ , c("Vertragsnummer", "HDL Nummer") ] )

empty_dealers<- which(is.empty(df$Händlernr.))

for (x in empty_dealers) {
  
  df$Händlernr. [x]  <-  if (  is.empty(df$Händlernr.[x] )   & length( A [ A$Vertragsnummer == df$Vertragsnummer[x] , "HDL Nummer"]) >0  )
  {
    A [ A$Vertragsnummer == df$Vertragsnummer[x] , "HDL Nummer"]
  }
  else {       df$Händlernr.[x]   }
}


df$Kraftstoffart[  df$Kraftstoffart == "NA"] <- ""

### in case of empty fuel description fields 
empty_krafstoff <- which ( is.empty(df$Kraftstoffart))

B<-left_join(   df [ empty_krafstoff  , ]  , df  [ , c("Kraftstoffart","Fahrzeugmodell") ] ,by=  "Fahrzeugmodell" , multiple = "first")

B<- B [ !duplicated(B$Fahrzeugmodell),]




for (z in empty_krafstoff) {
  
  df$Kraftstoffart [z]  <- {   B [ B$Fahrzeugmodell == df$Fahrzeugmodell[z] , "Kraftstoffart.y"]
  }
  
}

df$Kraftstoffart <- if_else(is.empty(df$Kraftstoffart) , "Benzin", df$Kraftstoffart)



### CREATE CHECK TABLE ###
RCD_Check <- data.frame("RCD Quality Check")
RCD_Check$Empty_vertragsnummer <- sum(is.empty(df$Vertragsnummer))
RCD_Check$Empty_Status <- sum(is.empty(df$Status))
RCD_Check$Empty_Handlernr <- sum(is.empty(df$Händlernr.))
RCD_Check$Empty_Kraftorfart <- sum(is.empty(df$Kraftstoffart))
RCD_Check$Empty_Dealer_Codes <- sum(is.empty(df$Händlernr.))
RCD_Check$Emtpy_Fahrzeugtyp <- sum(is.empty(df$Fahrzeugtyp))
RCD_Check$Wrong_Dealer_Code <-    sum(substr(df$Händlernr., 1, 3) != "C07"  , na.rm = TRUE)

print ( RCD_Check)


### SAVE  RCD EXCEL FILE IN SHAREPOINT 

wb <- openxlsx::createWorkbook()     
addWorksheet(wb, "Export" )
writeDataTable(wb,  "Export", x = df)
setColWidths(wb, 1, cols = c(1: ncol( df)), widths = "auto")
openxlsx:: saveWorkbook(wb, file = "HCBE_RCD_Mainsource.xlsx", overwrite = TRUE)

# library(openxlsx2)
# wb<- wb_load(  "HCBE_RCD_Mainsource.xlsx" )
# wb$remove_worksheet("Export")
# wb$add_worksheet("Export")
# wb$add_data("Export", df)
# wb_add_mips ( wb , xml = LABEL)
# wb_save(wb,  HCBE_RCD_Mainsource.xlsx )
#LABEL <-"<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_Enabled\"><vt:lpwstr>true</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"3\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_SetDate\"><vt:lpwstr>2024-05-27T10:38:16Z</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"4\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_Method\"><vt:lpwstr>Privileged</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"5\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_Name\"><vt:lpwstr>Public O365</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"6\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_SiteId\"><vt:lpwstr>35595a02-4d6d-44ac-99e1-f9ab4cd872db</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"7\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_ActionId\"><vt:lpwstr>a2770939-90dd-4f2f-a907-706f17ad8bcb</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"8\" name=\"MSIP_Label_41b88ec2-a72b-4523-9e84-0458a1764731_ContentBits\"><vt:lpwstr>0</vt:lpwstr></property"

fs::file_copy(
  path  = "HCBE_RCD_Mainsource.xlsx",
  new_path   =  paste0( Sharepoint_folder, "HCBE_RCD_Mainsource.xlsx"),
  overwrite = TRUE
)

## SAVE SUMMARY EXCEL FILE IN MY LOCAL PC 
wb <- openxlsx::createWorkbook()
addWorksheet(wb, "Check" )
writeDataTable(wb,  "Check", x = RCD_Check)
setColWidths(wb, 1, cols = c(1: ncol( RCD_Check)), widths = "auto")
openxlsx:: saveWorkbook(wb, file = paste0(documents_folder, "RCD_Check.xlsx"), overwrite = TRUE)

#### SEND EMAIL AUTOMATIC 
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["To"]] = c("juan.torrealba@de.hcs.com",  "anastasia.taranuha@de.hcs.com","philipp.hedwig@de.hcs.com" ) 
outMail[["subject"]] = paste ("RCD_Check_", today())
outMail[["body"]] = "This an automatic email: With the weekly check on RCD."
outMail[["attachments"]]$Add(paste0(documents_folder, "RCD_Check.xlsx"))
outMail$Send()






# 
# 
# BOOKINGS <- df %>% filter ( Brand == "Hyundai", Aktivierung >= "2024-01-01", Status=="Active", Fahrzeugtyp != "Gebrauchtwagen")
# 
# 
# BOOKINGS$ENGINE_TYPE<- if_else ( 
#   
#   grepl ( "Benzin" ,  BOOKINGS$Kraftstoffart)  & (     grepl ( "HEV" ,  BOOKINGS$Fahrzeugmodell) |   grepl ( "Hybrid" ,  BOOKINGS$Fahrzeugmodell)   )    ,
#                                  "PHEV/HEV", if_else( BOOKINGS$Kraftstoffart ==  "Elektrischer Strom",  "EV", "ICE")) 
# 
# 
#     




