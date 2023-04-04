#### Proposal Form Import Tool - Pilot
#### Wildlife Conservation Society - 2023
#### Script by Chloé Debyser

#### Workspace ####
# Packages
library(openxlsx)
library(tidyverse)
library(magrittr)
library(sf)
library(stringi)

# Input
proposalFolder <- "C:/Users/CDebyser/OneDrive - Wildlife Conservation Society/4. Analyses/Test workspace/KBACanadaProposalForms" # Change to path to folder on the server where proposal forms get stored
KBASite_input <- st_read("C:/Users/CDebyser/OneDrive - Wildlife Conservation Society/4. Analyses/Test workspace/TestGDB.gdb", "KBASite") # Pass this from Python instead

#### Processing ####
formPaths <- list.files(proposalFolder)
errors <- c()

for(formPath in formPaths){
  
  # ** LOAD KEY INFORMATION **
  # Load proposal form
  wb <- loadWorkbook(paste0(proposalFolder, "/", formPath))
  
  # Load proposal form sheets
        # HOME
  home <- read.xlsx(wb, sheet="HOME")
  
        # 1. PROPOSER
  proposer1 <- read.xlsx(wb, sheet="1. PROPOSER") %>%
    .[,2:3] %>%
    rename(Field = X2, Entry = X3)
  
        # 2. SITE
  site2 <- read.xlsx(wb, sheet="2. SITE") %>%
    .[,2:4] %>%
    rename(Field = X2)
  
  # Load key values
        # KBA Canada Form version
  formVersion <- home[1,1] %>%
    gsub("Version ", "", .) %>%
    as.numeric()
  
        # Site name
  siteName <- site2 %>% filter(Field == "National name") %>% pull(GENERAL) %>% trimws()
  
        # Site ID
  KBASiteID <- site2 %>% filter(Field == "KBA-EBAR Database ID") %>% pull(GENERAL) %>% as.integer()
  
        # Date assessed
  dateAssessed <- site2 %>%
    filter(Field == "Date (dd/mm/yyyy)") %>%
    pull(GENERAL) %>%
    as.character() %>%
    as.Date(., format="%d/%m/%Y") %>%
    as.POSIXlt(., format='%Y/%m/%d')
  
  # KBA Service data
  KBASite <- KBASite_input %>%
    filter(kbasiteid == KBASiteID)
  
  # ** CHECKS **
  # Check that the site name is identical in the proposal form and in the database
  if(!siteName == KBASite$nationalname){
    errors <- c(errors, paste("ERROR -", formPath, "not processed: site name is different in the proposal form and in the database."))
    next()
  }
  
  # ** POPULATE KBA SERVICE **
  # Site disclaimer
  KBASite %<>%
    mutate(disclaimer_en = paste0("<p>The information provided is current as of the date of KBA assessment (", format(dateAssessed, '%Y/%m/%d'), "). If you have more current information about the site, please <a href='/contact-us/'>contact us</a>.</p>"),
           disclaimer_fr = paste0("<p>Les informations présentées sont à jour en date du ", format(dateAssessed, '%Y/%m/%d'), " (date d'évaluation de la KBA). Si vous avez des informations plus récentes à propos du site, merci de <a href='/fr/contact-us/'>nous contacter</a>.</p>"))
  
  # ** COMPILE FINAL TABLE **
  if(!exists("KBASite_output")){
    KBASite_output <- KBASite
  }else{
    KBASite_output <- bind_rows(KBASite_output, KBASite)
  }
}
rm(KBASite_input, KBASite, formPath, formPaths, formVersion, KBASiteID, proposalFolder, dateAssessed, home, proposer1, site2, siteName, wb)

# NOTE:
# - KBASite_output: Use this simple feature collection to update KBASite in the database
# - errors: A vector of text entries describing errors encountered
