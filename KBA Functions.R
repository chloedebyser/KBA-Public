#### Identification of Canadian KBAs
#### Wildlife Conservation Society - 2021-2023
#### Script by Chloé Debyser and Zachary Moore

#### KEY FUNCTIONS

#### KBA Canada Proposal Form - Load and format ####
read_KBACanadaProposalForm <- function(formPath, final){
  
  # Load full workbook
  wb <- loadWorkbook(formPath)
  
  # Check that it is a KBA Canada Proposal Form
  if(sum(c("HOME", "1. PROPOSER", "2. SITE", "3. SPECIES","4. ECOSYSTEMS & C", "5. THREATS", "6. REVIEW", "7. CITATIONS", "8. CHECK") %in% getSheetNames(formPath)) != 9) {
    result <<- F
    message <<- "The file provided is not a KBA Canada Proposal Form."
    stop()
  }
  
  # Load individual sheets
        # Visible sheets
  home <- read.xlsx(wb, sheet = "HOME")
  proposer <- read.xlsx(wb, sheet = "1. PROPOSER")
  site <- read.xlsx(wb, sheet = "2. SITE")
  species <- read.xlsx(wb, sheet = "3. SPECIES")
  ecosystems <- read.xlsx(wb, sheet = "4. ECOSYSTEMS & C")
  threats <- read.xlsx(wb, sheet = "5. THREATS")
  review <- read.xlsx(wb, sheet = "6. REVIEW")
  citations <- read.xlsx(wb, sheet = "7. CITATIONS")
  check <- read.xlsx(wb, sheet = "8. CHECK")
  
        # Invisible sheets
  checkboxes <- read.xlsx(wb, sheet = "checkboxes")
  resultsSpecies <- read.xlsx(wb, sheet = "results_species")
  resultsEcosystems <- read.xlsx(wb, sheet = "results_ecosystems")
  
  # Site name
        # Get the site name
  nationalName <<- site[1,"GENERAL"]
  
        # Check that the name exists
  if(is.na(nationalName)){
    result <<- F
    message <<- "KBA site must have a name."
    stop()
  }
  
        # Check that the name is not too long
  if(nchar(nationalName) > 255){
    result <<- F
    message <<- "KBA name is too long."
    stop()
  }
  
        # Check that the name does not contain any paragraph symbols
  if(grepl("\n", nationalName, fixed=T)){
    result <<- F
    message <<- "KBA name should not include paragraph breaks."
    stop()
  }
  
  # Form version
        # Get the form version
  formVersion <<- home[1,1] %>%
    gsub("Version ", "", .) %>%
    as.numeric()
  
        # Check form version compatibility
  if(!formVersion %in% c(1, 1.1, 1.2)){
    result <<- F
    message <<- "KBA Canada Form version not supported."
    stop()
  }
  
  # KBA level
  KBALevel <- home %>%
    filter(X3 == "Criteria met") %>%
    pull(X4) %>%
    {ifelse(grepl("g", ., T), "Global", "National")}
  
  # KBA criteria
  if(formVersion < 1.2){
    KBACriteria <- str_split(home[13,4], "; ")[[1]]
    
  }else{
    KBACriteria <- str_split(home[16,4], "; ")[[1]]
  }
  
  # 1. PROPOSER
        # Overall formatting
  proposer %<>%
    .[, 2:3] %>%
    rename(Field = X2, Entry = X3) %>%
    filter(!is.na(Field))
  
        # Handle differing form versions
  if(formVersion < 1.2){
    
    proposer %<>%
      pivot_wider(., names_from = Field, values_from = Entry) %>%
      mutate(`Other affiliations with a KBA Partner` = ifelse(is.na(`Other affiliations with a KBA Partner`) & (!`KBA Partner represented` == "WCS"), `KBA Partner represented`, `Other affiliations with a KBA Partner`)) %>%
      select(-c(Address, `Country of residence`, `Membership in a KBA National Coordination Group`, `KBA Partner represented`, `Membership in an IUCN Specialist group`, `Name of the IUCN Specialist group`, `Main country of interest`, `Second country of interest`, `Prior experience proposing KBAs`, `Details of prior experience`, `Email (please re-enter)`, `Main taxon or group of interest`, `Second taxon or group of interest`)) %>%
      mutate(`Name of contact person` = "Ciara Raudsepp-Hearne",
             Address = "Suite 204, 344 Bloor Street West, Toronto, Ontario, M5S 3A7",
             `Country of residence` = "Canada",
             `Email of contact person` = "craudsepp@wcs.org",
             `Organization of contact person` = "Wildlife Conservation Society Canada",
             `Membership in a KBA National Coordination Group` = "Canada",
             `KBA Partner represented` = "WCS",
             .before = "Name") %>%
      relocate(`Other affiliations with a KBA Partner`, .before = "Name") %>%
      mutate(`Membership in an IUCN Specialist group` = NA,
             `Name of the IUCN Specialist group` = NA,
             `Main country of interest` = "Canada",
             `Second country of interest` = NA,
             `Prior experience proposing KBAs` = "Yes",
             `Details of prior experience` = "The KBA Canada Secretariat, hosted by Wildlife Conservation Society Canada, Birds Canada, and NatureServe Canada, is coordinating the work of KBA identification in Canada.",
             .before = "Name") %>%
      rename(`Name of proposal development lead` = "Name",
             `Email of proposal development lead` = "Email",
             `Organization of proposal development lead` = "Organization") %>%
      relocate(`Names and affiliations`, .after = "Organization of proposal development lead") %>%
      mutate(`Name(s) to be displayed publicly` = NA,
             .after ="Names and affiliations") %>%
      pivot_longer(everything(),
                   names_to = "Field",
                   values_to = "Entry") %>%
      mutate(Field = replace(Field, Field == "I agree to the data in this form being stored in the World Database of KBAs and used for the purposes of KBA identification and conservation.", "I agree to the data in this form being stored in the Canadian KBA Registry and in the World Database of KBAs, and used for the purposes of KBA identification and conservation."))
  }
  
        # Spell out "WCS"
  proposer %<>%
    mutate(Entry = case_when(Field %in% c("Organization of contact person", "Organization of proposal development lead", "Names and affiliations") ~ gsub("WCS", "Wildlife Conservation Society", Entry, fixed=T), .default=Entry))
  
  # 2. SITE
        # Overall formatting
  site %<>%
    .[, 2:4] %>%
    rename(Field = X2) %>%
    filter(!Field == "Ongoing                                                                                           Needed                                                  ")
  
        # Handle differing form versions
  if(formVersion < 1.2){
    
    site %<>%
      mutate(Field = replace(Field, Field == "Site management", "Conservation"),
             Field = replace(Field, Field == "Longitude (dd.dddd)", "Longitude (ddd.dddd)")) %>%
      add_row(Field = "Disclaimer", .after = site %>% with(which(Field == "Additional biodiversity"))) %>%
      add_row(Field = "Site history", .before = site %>% with(which(Field == "Forest"))) %>%
      add_row(Field = "Customary jurisdiction source", .after = site %>% with(which(Field == "Customary jurisdiction")))
    
  }else{
    
    site %<>%
      add_row(Field = "Names and emails", .before = site %>% with(which(Field == "Country")))
  }
  
        # Assessment date
  dateAssessed <- site %>%
    filter(Field == "Date (dd/mm/yyyy)") %>%
    pull(GENERAL)
  
  if(final | !is.na(dateAssessed)){
    
    if(grepl("/", dateAssessed)){
      
      if(substr(dateAssessed, start=3, stop=3) == "/"){
        
        dateAssessed %<>%
          as.Date(., format="%d/%m/%Y") %>%
          as.character()
        
      }else{
        stop("Date format not supported")
      }
      
    }else{
      
      dateAssessed %<>%
        as.integer() %>%
        as.Date(., origin = "1899-12-30") %>%
        as.character()
    }
    
    if(is.na(dateAssessed)){
      stop("Date assessed not correctly processed")
    }
    
    site[which(site$Field == "Date (dd/mm/yyyy)"), "GENERAL"] <- dateAssessed
  }
  
        # Site version
  if(formVersion < 1.2){
    
    siteVersion <- NA
    
  }else{
    
    siteVersion <- site %>%
      filter(Field == "Site version") %>%
      pull(GENERAL) %>%
      as.integer()
  }
  
        # Site area
              # Compute and round site area
  siteArea <- site %>%
    filter(Field == "Site area (km2)") %>%
    pull(GENERAL) %>%
    as.numeric() %>%
    round(., digits=2)
    
              # Replace value
  site %<>%
    mutate(GENERAL = replace(GENERAL, Field == "Site area (km2)", siteArea))
  
        # Latitude and longitude
              # Round the values
  lat <- site %>%
    filter(Field == "Latitude (dd.dddd)") %>%
    pull(GENERAL) %>%
    as.numeric() %>%
    round(., digits=4)
  
  lon <- site %>%
    filter(Field == "Longitude (ddd.dddd)") %>%
    pull(GENERAL) %>%
    as.numeric() %>%
    round(., digits=4)
  
              # Replace the values 
  site %<>%
    mutate(GENERAL = replace(GENERAL, Field == "Latitude (dd.dddd)", lat),
           GENERAL = replace(GENERAL, Field == "Longitude (ddd.dddd)", lon))
  
        # Conservation actions
  actionsCol <- which(colnames(checkboxes) == "2..Conservation.actions")
  actions <- checkboxes %>%
    .[, actionsCol:(actionsCol+2)]
  colnames(actions) <- actions[1,]
  actions %<>%
    .[2:nrow(.),]
  
  # 3. SPECIES
        # Overall formatting
  colnames(species) <- species[1,]
  species %<>%
    .[2:nrow(.),] %>%
    filter(!is.na(`Common name`)) %>%
    mutate_if(is.character, trimws) %>%
    mutate(`Derivation of best estimate` = ifelse(`Derivation of best estimate` == "Other (please add further details in column AA)", "Other", `Derivation of best estimate`),
           `Red List SIS number` = as.integer(`Red List SIS number`)) %>%
    arrange(`Scientific name`)
  
        # Extract species for delisting
  sppDelisting <- species %>%
    filter(substr(species$`Explanation of site estimates`, start=1, stop=10) == "DELISTING:") %>%
    select(`Common name`, `Scientific name`, `Explanation of site estimates`) %>%
    mutate(`Delisting reason` = trimws(sub("DELISTING:", "", sub("RATIONALE:.*", "", `Explanation of site estimates`))) %>% gsub(".", "", ., fixed=T),
           `Rationale text` = trimws(sub("RATIONALE:", "", sub(".*RATIONALE:", "", `Explanation of site estimates`)))) %>%
    select(-`Explanation of site estimates`)
  
  species %<>%
    filter(is.na(`Explanation of site estimates`) | !substr(species$`Explanation of site estimates`, start=1, stop=10) == "DELISTING:")
  
        # Handle scientific names of the type "sp. #"
  species$`Scientific name` <- sapply(species$`Scientific name`, function(x) trimws(ifelse(!grepl(" sp. ", x, fixed=T), gsub(" sp.", " sp. ", x, fixed=T), x))) %>% as.character()
  
        # Handle differing form versions
  if(formVersion < 1.2){
    
    species %<>%
      rename(`RU source` = `RU Source`,
             `Seasonal distribution` = `Seasonal Distribution`,
             `Evidence year` = `Evidence Year`) %>%
      mutate(display_taxonomicgroup = "Yes",
             display_taxonname = "Yes",
             display_alternativename = NA,
             display_assessmentinfo = "Yes",
             display_biodivelementdist = "Yes")
    
  }else{
    
    species %<>%
      rename(display_taxonomicgroup = `Display taxonomic group?`,
             display_taxonname = `Display taxon name?`,
             display_alternativename = `Alternative name to display`,
             display_assessmentinfo = `Display assessment information?`,
             display_biodivelementdist = `Display internal boundary?`)
  }
  
  # 4. ECOSYSTEMS & C
        # Ecological integrity
  ecologicalIntegrity <- ecosystems
  start <- as.integer(rownames(ecosystems[which(ecosystems$X1=="Name of ecoregion"),]))
  colnames(ecologicalIntegrity) <- ecologicalIntegrity[start,]
  ecologicalIntegrity %<>%
    .[(start+1):nrow(.),] %>%
    .[, colnames(.)[which(!is.na(colnames(.)))]] %>%
    drop_na(`Name of ecoregion`)
  
  if(nrow(ecologicalIntegrity) > 0){
    result <<- F
    message <<- "Criterion C KBAs not yet supported."
    stop()
  }
  
        # Ecosystems (A2 and B4)
  stop <- start-3
  start <- as.integer(rownames(ecosystems[which(ecosystems$X1=="Name of ecosystem type"),]))
  colnames(ecosystems) <- ecosystems[start,]
  
  if(start < stop){
    ecosystems %<>%
      .[(start+1):stop,]
  }else{
    ecosystems %<>% .[0,]
  }
  
  if(nrow(ecosystems) > 0){
    ecosystems %<>% 
      mutate(`Name of ecosystem type` = trimws(`Name of ecosystem type`)) %>%
      arrange(`Name of ecosystem type`)
  }
  
  # 5. THREATS
          # Verify whether "No Threats" checkbox is checked
  noThreatsCol <- which(colnames(checkboxes) == "5..Threats")
  noThreats <- checkboxes[2, (noThreatsCol+1)] %>% as.logical()
  
          # If there are threats, get that information
  if(!noThreats){
    colnames(threats) <- threats[3,]
    threats %<>% .[4:nrow(.),]
    colnames(threats)[ncol(threats)] <- "Notes"
    threats %<>% mutate(`Specific biodiversity element` = trimws(`Specific biodiversity element`))
    
    if(nrow(threats) == 0){
      result <<- F
      message <<- "The 'No threats' box isn't checked, and yet there aren't any threats entered in tab 5."
      stop()
    }
    
  }else{
    colnames(threats) <- threats[3,]
    colnames(threats)[ncol(threats)] <- "Notes"
    threats %<>% filter(!row_number() %in% rownames(threats))
  }
  
        # Add SpeciesIDs, where applicable
  threats <- species %>%
    select(`Common name`, `NatureServe Element Code`) %>%
    distinct() %>%
    left_join(threats, ., by=c("Specific biodiversity element"="Common name")) %>%
    left_join(., DB_BIOTICS_ELEMENT_NATIONAL[which(!is.na(DB_BIOTICS_ELEMENT_NATIONAL$element_code)), c("element_code", "speciesid")], by=c("NatureServe Element Code" = "element_code")) %>%
    relocate(., speciesid, .after = 'Specific biodiversity element') %>%
    select(-`NatureServe Element Code`)
  
        # Add EcosystemIDs, where applicable
  threats %<>%
    left_join(., DB_BIOTICS_ECOSYSTEM[which(!is.na(DB_BIOTICS_ECOSYSTEM$cnvc_english_name)), c("ecosystemid", "cnvc_english_name")], by=c("Specific biodiversity element" = "cnvc_english_name")) %>%
    relocate(., ecosystemid, .after = 'Specific biodiversity element')
  
  # 6. REVIEW
        # Overall formatting
  review %<>%
    drop_na(X2) %>%
    fill(`INSTRUCTIONS:`)
  
        # Technical review
  technicalReview <- review %>%
    filter(`INSTRUCTIONS:` == 1) %>%
    select(-`INSTRUCTIONS:`)
  
  colnames(technicalReview) <- technicalReview[2,]
  
  if(nrow(technicalReview) > 2){
    technicalReview %<>% .[3:nrow(.),]
  }else{
    technicalReview[3,] <- c("No reviewers listed", "", "", "")
    technicalReview %<>% .[3:nrow(.),]
  }
  
        # General review
  generalReview <- review %>%
    filter(`INSTRUCTIONS:` == 2) %>%
    select(-c(`INSTRUCTIONS:`))
  
  if(is.na(generalReview[2,4])){
    generalReview %<>% select(-X5)
  }
  
  colnames(generalReview) <- generalReview[2,]
  
  if(nrow(generalReview) > 2){
    generalReview %<>% .[3:nrow(.),]
  }else{
    generalReview[3,] <- c("No reviewers listed", rep("", (ncol(generalReview)-1)))
    generalReview %<>% .[3:nrow(.),]
  }
  
        # No review provided
  noFeedback <- review$X3[which(review$X2 == "Provide information about any organizations you contacted and that did not provide feedback.")]
  noFeedback <- ifelse(is.na(noFeedback), "None", noFeedback)
  
  # 7. CITATIONS
  colnames(citations) <- c("Short citation", "Long citation", "DOI", "URL")
  citations %<>%
    .[3:nrow(.), 1:4] %>%
    drop_na(`Short citation`) %>%
    mutate(`Short citation` = trimws(`Short citation`),
           Sensitive = ifelse(grepl('[SENSITIVE]', `Short citation`, fixed=T), 1, 0),
           `Short citation` = gsub("[SENSITIVE]", "", `Short citation`, fixed=T),
           `Short citation` = trimws(`Short citation`),
           `Long citation` = trimws(`Long citation`))
  
  # 8. CHECK
        # Column names
  colnames(check) <- c("Check", "Item")
  
        # Get checkbox results
  check_checkboxes <- checkboxes %>%
    .[2:nrow(.),] %>%
    select("8..Checks") %>%
    drop_na()
  
  if(formVersion %in% c(1, 1.1)){
    check_checkboxes %<>% .[c(1:5,7:nrow(.)),] # Cell N8 is obsolete in v1.1 of the Proposal Form (it doens't link to any actual checkbox)
  }else{
    check_checkboxes %<>% pull(`8..Checks`)
  }
  
        # Verify that there are as many checkbox results as there are checkboxes
  if(!(nrow(check) == length(check_checkboxes))){
    convert_res[step,"Result"] <- emo::ji("prohibited")
    convert_res[step,"Message"] <- "Inconsistencies between the 8. CHECKS tab and checkbox results. This error originates from the Excel formulas themselves. Please contact Chloé and provide her with this error message."
    KBAforms[step] <- NA
    next
  }
  
        # Add checkbox results to the 8. CHECK tab
  check %<>%
    select(-Check) %>%
    mutate(Check = check_checkboxes)
  
  # RESULTS - Ecosystems
  resultsCols_start <- which(colnames(resultsEcosystems) == "RESULT.COLUMNS")
  resultsCols_end <- which(colnames(resultsEcosystems) == "MANDATORY.FIELDS.COMPLETED")-1
  resultsCols <- paste0("RESULT_", resultsEcosystems[1, resultsCols_start:resultsCols_end])
  
  mandatoryCols <- paste0("MANDATORY_", resultsEcosystems[1, (resultsCols_end+1):ncol(resultsEcosystems)])
  
  starterCols <- resultsEcosystems[1, 1:(resultsCols_start-1)]
  
  colnames(resultsEcosystems) <- c(starterCols, resultsCols, mandatoryCols)
  
  resultsEcosystems %<>%
    .[3:nrow(.),]
  
  # RESULTS - Species
  resultsCols_start <- which(colnames(resultsSpecies) == "RESULT.COLUMNS")
  resultsCols_end <- which(colnames(resultsSpecies) == "MANDATORY.FIELDS.COMPLETED")-1
  resultsCols <- paste0("RESULT_", resultsSpecies[1, resultsCols_start:resultsCols_end])
  
  mandatoryCols <- paste0("MANDATORY_", resultsSpecies[1, (resultsCols_end+1):ncol(resultsSpecies)])
  
  starterCols <- resultsSpecies[1, 1:(resultsCols_start-1)]
  
  colnames(resultsSpecies) <- c(starterCols, resultsCols, mandatoryCols)
  
  resultsSpecies %<>%
    .[3:nrow(.),]
    
  # Return outputs
  PF_home <<- home
  PF_formVersion <<- formVersion
  PF_KBALevel <<- KBALevel
  PF_KBACriteria <<- KBACriteria
  PF_proposer <<- proposer
  PF_site <<- site
  PF_nationalName <<- nationalName
  PF_siteVersion <<- siteVersion
  PF_actions <<- actions
  PF_species <<- species
  PF_ecosystems <<- ecosystems
  PF_ecologicalIntegrity <<- ecologicalIntegrity
  PF_noThreats <<- noThreats
  PF_threats <<- threats
  PF_technicalReview <<- technicalReview
  PF_generalReview <<- generalReview
  PF_noReview <<- noFeedback
  PF_citations <<- citations
  PF_check <<- check
  PF_resultsEcosystems <<- resultsEcosystems
  PF_resultsSpecies <<- resultsSpecies
  PF_delistedSpecies <<- sppDelisting
}

#### KBA Canada Proposal Form - Convert to Global Multi-Site Form ####
convert_toGlobalMultiSiteForm <- function(templatePath, reviewedProposal, speciesOrder){
  
  # Check that there are global criteria met
  criteriaMet <- PF_home %>%
    filter(X3 == "Criteria met") %>%
    pull(X4)
  
  if(!grepl("g", criteriaMet, fixed=T)){
    stop("No Global Criteria met.")
  }
  
  # Load template for the Multi-Site Global Proposal Form
  multiSiteForm_wb <- loadWorkbook(templatePath) %>%
    suppressWarnings()
  
  # Only retain global triggers
        # Species
              # Classify species by level  
                    # Species assessed nationally
  speciesN <- PF_species %>%
    filter(`KBA level` == "National") %>%
    pull(`NatureServe Element Code`)
    
                    # Species assessed globally
  speciesG <- PF_species %>%
    filter(`KBA level` == "Global") %>%
    pull(`NatureServe Element Code`)
    
              # Species assessed nationally and not globally
                    # Format dataset
  speciesNnotG <- speciesN[which(!speciesN %in% speciesG)]
  speciesNnotG <- PF_species %>%
    filter(`NatureServe Element Code` %in% speciesNnotG) %>%
    mutate(Trigger = ifelse(is.na(`Criteria met`), "No", "Yes")) %>%
      select(`Common name`, `Scientific name`, Trigger) %>%
      distinct() %>%
      group_by(`Common name`, `Scientific name`) %>%
      arrange(desc(Trigger)) %>%
      filter(row_number()==1) %>%
      ungroup() %>%
      mutate(Text = paste0(`Common name`, " (", `Scientific name`, ")"))
  
                    # Create text
  speciesNnotG_triggers <- speciesNnotG %>%
    filter(Trigger == "Yes") %>%
    pull(Text) %>%
    paste(., collapse="; ")
  speciesNnotG_triggers <- ifelse(!speciesNnotG_triggers=="",
                                  paste0("Taxa that meet national KBA criteria in Canada: ", speciesNnotG_triggers, "."),
                                  "")
    
  speciesNnotG_notTriggers <- speciesNnotG %>%
    filter(Trigger == "No") %>%
    pull(Text) %>%
    paste(., collapse="; ")
  speciesNnotG_notTriggers <- ifelse(!speciesNnotG_notTriggers=="",
                                     paste("Taxa assessed against national KBA criteria in Canada, but that did not meet criteria:", speciesNnotG_notTriggers),
                                     "")
  
  speciesNnotG_text <- paste(speciesNnotG_triggers, speciesNnotG_notTriggers) %>% trimws()
    
                # Only retain global triggers for further processing
  PF_species %<>%
    filter(`KBA level` == "Global")
    
  if(nrow(PF_species) > 0){
    PF_species$SpeciesIndex <- sapply(1:nrow(PF_species), function(x) paste0("spp", x))
  }
    
  PF_species <<- PF_species
  
  PF_resultsSpecies %<>%
    filter(`KBA level` == "Global")
    
  if(nrow(PF_resultsSpecies) > 0){
    PF_resultsSpecies$SpeciesIndex <- sapply(1:nrow(PF_resultsSpecies), function(x) paste0("spp", x))
  }
    
  PF_resultsSpecies <<- PF_resultsSpecies
    
          # Ecosystems
                # Classify ecosystems by level  
                      # Ecsosystemss assessed nationally
  ecosystemsN <- PF_ecosystems %>%
    filter(`KBA level` == "National") %>%
    pull(`Name of ecosystem type`)
  
                      # Ecosystems assessed globally
  ecosystemsG <- PF_ecosystems %>%
    filter(`KBA level` == "Global") %>%
    pull(`Name of ecosystem type`)
    
                # Ecosystems assessed nationally and not globally
                      # Format dataset
  ecosystemsNnotG <- ecosystemsN[which(!ecosystemsN %in% ecosystemsG)]
  ecosystemsNnotG <- PF_ecosystems[which(PF_ecosystems$`Name of ecosystem type` == ecosystemsNnotG), ] %>%
    mutate(Trigger = ifelse(is.na(`Criteria met`), "No", "Yes")) %>%
    select(`Name of ecosystem type`, Trigger) %>%
    distinct() %>%
    group_by(`Name of ecosystem type`) %>%
    arrange(desc(Trigger)) %>%
    filter(row_number()==1) %>%
    ungroup()
    
                      # Create text
  ecosystemsNnotG_triggers <- ecosystemsNnotG %>%
    filter(Trigger == "Yes") %>%
    pull(`Name of ecosystem type`) %>%
    paste(., collapse="; ")
  ecosystemsNnotG_triggers <- ifelse(!ecosystemsNnotG_triggers=="",
                                     paste0("Ecosystem types that meet national KBA criteria in Canada: ", ecosystemsNnotG_triggers, "."),
                                     "")
    
  ecosystemsNnotG_notTriggers <- ecosystemsNnotG %>%
    filter(Trigger == "No") %>%
    pull(`Name of ecosystem type`) %>%
    paste(., collapse="; ")
  ecosystemsNnotG_notTriggers <- ifelse(!ecosystemsNnotG_notTriggers=="",
                                        paste("Ecosystem types assessed against national KBA criteria in Canada, but that did not meet criteria:", ecosystemsNnotG_notTriggers),
                                        "")
    
  ecosystemsNnotG_text <- paste(ecosystemsNnotG_triggers, ecosystemsNnotG_notTriggers) %>% trimws()
  
                # Only retain global triggers for further processing
  PF_ecosystems %<>%
    filter(`KBA level` == "Global")
  
  if(nrow(PF_ecosystems) > 0){
    PF_ecosystems$EcosystemIndex <- sapply(1:nrow(PF_ecosystems), function(x) paste0("ecosys", x))
  }
    
  PF_ecosystems <<- PF_ecosystems
  
  PF_resultsEcosystems %<>%
    filter(`KBA level` == "Global")
    
  if(nrow(PF_resultsEcosystems) > 0){
    PF_resultsEcosystems$EcosystemIndex <- sapply(1:nrow(PF_resultsEcosystems), function(x) paste0("ecosys", x))
  }
    
  PF_resultsEcosystems <<- PF_resultsEcosystems
    
          # Ecological integrity (criterion C)
  if(sum(!PF_ecologicalIntegrity$`Assessment level` == "Global") > 0){
    stop("Use case not implemented: national criterion C sites")
  }
    
  if(nrow(PF_ecologicalIntegrity) > 0){
    PF_ecologicalIntegrity$EcoregionIndex <- sapply(1:nrow(PF_ecologicalIntegrity), function(x) paste0("ecoreg", x))
  }
    
    # Only retain threats associated with global triggers
  PF_threats %<>%
    filter((Category == "Entire site") | (`Specific biodiversity element` %in% PF_species$`Common name`) | (`Specific biodiversity element` %in% PF_ecosystems$`Name of ecosystem type`))
    
  PF_threats <<- PF_threats
    
    # Prepare review texts
          # Technical review
  technicalReview_text <- ""
    
  if(nrow(PF_technicalReview) > 0){
    
    for(i in 1:nrow(PF_technicalReview)){
      inParentheses <- c(PF_technicalReview$Affiliation[i], PF_technicalReview$Email[i]) %>% .[which(!is.na(.))]
      technicalReview_text <- paste0(technicalReview_text, paste0(PF_technicalReview$Name[i], ifelse(length(inParentheses)>0, paste0(" (", paste0(inParentheses, collapse=", "), ")"), ""), ifelse(!is.na(PF_technicalReview$`Description of role`[i]), paste0(": ", PF_technicalReview$`Description of role`[i]), "")), sep="\n")
    }
  }
          # General review
  generalReview_text <- ""
    
  if(nrow(PF_generalReview) > 0){
      
    for(i in 1:nrow(PF_generalReview)){
      inParentheses <- c(PF_generalReview$Affiliation[i], PF_generalReview$Email[i]) %>% .[which(!is.na(.))]
      generalReview_text <- paste0(ifelse(!is.na(PF_generalReview$Name[i]), PF_generalReview$Name[i],""), ifelse(length(inParentheses)>0, paste0(" (", paste0(inParentheses, collapse=", "), ")"), "")) %>%
        trimws() %>%
        ifelse(substr(., start=1, stop=1) == "(", substr(., start=2, stop=nchar(.)-1), .) %>%
        paste0(generalReview_text, ., sep="\n")
    }
  }
    
          # Reviewers who did not participate
  if(!is.na(nchar(PF_noReview))){
    generalReview_text %<>% paste0(., "\n\n", PF_noReview)
  }
    
    # Remove HTML tags from citations
  PF_citations %<>%
    mutate(`Long citation` = gsub("<i>", "", `Long citation`)) %>%
    mutate(`Long citation` = gsub("</i>", "", `Long citation`))
    
    # Get site IDs
          # Site record number in WDKBA (sitrecid)
  WDKBAnumber <- PF_site %>%
    filter(Field == "WDKBA number") %>%
    pull(GENERAL)
    
          # Site code
  CASiteCode <- PF_site %>%
    filter(Field == "Canadian Site Code") %>%
    pull(GENERAL)
    
  if(is.na(CASiteCode)){
    stop("No site code.")
  }
    
    # Populate the global multi-site form
          # 1. Proposer
                # Name
  contactPerson <- PF_proposer %>% filter(Field == "Name of contact person") %>% pull(Entry)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=contactPerson, xy=c(2,3))
    
                # Address
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Address") %>% pull(Entry), xy=c(2,4))
    
                # Country of residence
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Country of residence") %>% pull(Entry), xy=c(2,5))
    
                # Email
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Email of contact person") %>% pull(Entry), xy=c(2,6))
    
                # Email (please re-enter)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Email of contact person") %>% pull(Entry), xy=c(2,7))
    
                # Organisation
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Organization of contact person") %>% pull(Entry), xy=c(2,8))
    
                # Names and affiliations of any co-proposers for this KBA nomination
                      # Get proposal development lead(s)
                            # Names
  proposalLeads <- PF_proposer %>%
    filter(Field == "Name of proposal development lead") %>%
    pull(Entry) %>%
    gsub("; ", ";", ., fixed=T) %>%
    strsplit(., split=";") %>%
    .[[1]]
    
                            # Organizations
  proposalLeadOrganizations <- PF_proposer %>%
    filter(Field == "Organization of proposal development lead") %>%
    pull(Entry) %>%
    gsub("; ", ";", ., fixed=T) %>%
    strsplit(., split=";") %>%
    .[[1]]
  
  if((length(proposalLeadOrganizations) == 1) & (length(proposalLeads) > 1)){
    proposalLeadOrganizations <- rep(proposalLeadOrganizations, length(proposalLeads))
      
  }else if(!length(proposalLeadOrganizations) == length(proposalLeads)){
    stop("Not the same number of proposal development lead names and organizations")
  }
          
                            # Full text
  proposalLeads_full <- paste0(paste(proposalLeads, proposalLeadOrganizations, sep=" ("), ")")
    
                      # Get co-proposers
  coproposers_full <- PF_proposer %>%
    filter(Field == "Names and affiliations") %>%
    pull(Entry) %>%
    gsub("; ", ";", ., fixed=T) %>%
    strsplit(., split=";") %>%
    .[[1]]
    
                      # All proposers
  proposers <- c(proposalLeads_full, coproposers_full) %>%
    unique() %>%
    {.[which(!is.na(.))]} %>%
    {.[which(!grepl(contactPerson, .))]} %>%
    paste(., collapse="; ")
    
                      # Save
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=proposers, xy=c(2,9))
    
                # Do you represent a KBA Partner organisation? Leave blank if not
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "KBA Partner represented") %>% pull(Entry), xy=c(2,12))
    
                # Do you have other affiliations with a KBA Partner organisation - if so which (leave blank otherwise)?
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Other affiliations with a KBA Partner") %>% pull(Entry), xy=c(2,13))
    
                # What is your main country of interest?
  mainCountry <- PF_proposer %>% filter(Field == "Main country of interest") %>% pull(Entry)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=mainCountry, xy=c(2,14))
    
                # Are you a member of this country's KBA National Coordination Group?
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=ifelse(mainCountry == PF_proposer %>% filter(Field == "Membership in a KBA National Coordination Group") %>% pull(Entry), "Yes", "No"), xy=c(2,15))
    
                # Do you have a second country of interest? (optional)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Second country of interest") %>% pull(Entry), xy=c(2,16))
    
                # What is your main taxon or group of interest?

                # Do you have a second taxon of interest? (optional)
    
                # Are you, or one of the co-proposers, a member of an IUCN Specialist Group? (optional)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Membership in an IUCN Specialist group") %>% pull(Entry), xy=c(2,19))
    
                # Name of Group
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Name of the IUCN Specialist group") %>% pull(Entry), xy=c(4,19))
    
                # Have you or one of your co-proposers, identified or proposed one or more KBAs before? (optional)
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Prior experience proposing KBAs") %>% pull(Entry), xy=c(2,20))
    
                # Please give further details
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "Details of prior experience") %>% pull(Entry), xy=c(5,20))
    
                # Date of proposal or nomination (dd/mm/yyyy):
  date <- PF_site %>%
    filter(Field == "Date (dd/mm/yyyy)") %>%
    pull(GENERAL) %>%
    as.Date()
    
  assessmentYear <<- date %>%
    format(., format="%Y") %>%
    as.integer()
    
  date %<>% format(., "%d/%m/%Y")
    
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=date, xy=c(2,24))
    
                # Is this a proposal or a nomination?:
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=ifelse(nVersion==0, "Proposal (before review)", "Nomination (following review)"), xy=c(2,25))
    
                # I agree to the data in this form being stored in the WDKBA and used for the purposes of KBA identification and conservation
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "I agree to the data in this form being stored in the Canadian KBA Registry and in the World Database of KBAs, and used for the purposes of KBA identification and conservation.") %>% pull(Entry), xy=c(2,30))
    
                # If no, please give further details
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=PF_proposer %>% filter(Field == "If no, please provide further details:") %>% pull(Entry), xy=c(5,30))
    
                # Reviewers
  suggestedReviewers <- PF_site %>%
    filter(Field == "Names and emails") %>%
    pull(GENERAL) %>%
    {ifelse(is.na(.), "A thorough review of this site has been conducted within Canada, as documented in the 'Consultation' section. If additional review is required, please get touch with us.", .)}
  writeData(multiSiteForm_wb, sheet = "1. Proposer", x=suggestedReviewers, xy=c(1,38))
    
          # 2. Site data
                # Site record number in WDKBA (sitrecid)
  if(!is.na(WDKBAnumber)){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x=WDKBAnumber, xy=c(1,5))
  }
    
                # Site code if new site
  if(is.na(WDKBAnumber)){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x=CASiteCode, xy=c(2,5))
  }
    
                # Site name (english)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "International name") %>% pull(GENERAL), xy=c(3,5))
    
                # Site name (national)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "National name") %>% pull(GENERAL), xy=c(4,5))
    
                # Country
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Country") %>% pull(GENERAL), xy=c(5,5))
    
                # Purpose of proposal for the site
  if(!reviewedProposal){
    
    if(globalPurpose[it] == "New site"){
      writeData(multiSiteForm_wb, sheet = "2. Site data", x="Propose a new KBA that does not intersect any existing KBAs", xy=c(6,5))
        
    }else if(globalPurpose[it] == "New biodiversity"){
      writeData(multiSiteForm_wb, sheet = "2. Site data", x="Add new qualifying biodiversity element to an existing KBA", xy=c(6,5))
    
    }else if(globalPurpose[it] == "Reassessed biodiversity"){
      writeData(multiSiteForm_wb, sheet = "2. Site data", x="Re-assess existing qualifying biodiversity elements", xy=c(6,5))
      
    }else if(globalPurpose[it] == "Amalgamated KBA"){  
      writeData(multiSiteForm_wb, sheet = "2. Site data", x="Propose a new KBA that amalgamates one or more existing KBAs into a new KBA", xy=c(6,5))
        
    }else{
      stop("Purpose of global submission not recognized")
    }
  }
  
                # A1
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gA1", criteriaMet, fixed=T), "Yes", "No"), xy=c(7,5))
    
                # A2
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gA2", criteriaMet, fixed=T), "Yes", "No"), xy=c(8,5))
    
                # B1
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gB1", criteriaMet, fixed=T), "Yes", "No"), xy=c(9,5))
    
                # B2
  writeData(multiSiteForm_wb, sheet = "2. Site data", x="No", xy=c(10,5))
    
                # B3
  writeData(multiSiteForm_wb, sheet = "2. Site data", x="No", xy=c(11,5))
    
                # B4
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gB4", criteriaMet, fixed=T), "Yes", "No"), xy=c(12,5))
    
                # C
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gC", criteriaMet, fixed=T), "Yes", "No"), xy=c(13,5))
    
                # D1
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gD1", criteriaMet, fixed=T), "Yes", "No"), xy=c(14,5))
    
                # D2
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gD2", criteriaMet, fixed=T), "Yes", "No"), xy=c(15,5))
    
                # D3
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=ifelse(grepl("gD3", criteriaMet, fixed=T), "Yes", "No"), xy=c(16,5))
    
                # Names of existing KBAs intersected that will be replaced by this site.
  if(!reviewedProposal){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x=globalKBAReplaced[it], xy=c(17,5))
  }
    
                # Why existing qualifying element needs to be changed
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Reassessment rationale") %>% pull(GENERAL), xy=c(18,5))
    
                # State/province
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Province or Territory") %>% pull(GENERAL), xy=c(19,5))
    
                # Site area (sq km)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$areakm2, xy=c(20,5))
    
                # Latitude of mid point (dd.dddd)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$lat_wgs84, xy=c(21,5))
    
                # Longitude of mid point (dd.dddd)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$long_wgs84, xy=c(22,5))
    
                # Lowest altitude (m asl) or deepest Bathymetry level (m)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$altitudemin, xy=c(23,5))
    
                # Highest altitude (m asl) or shallowest Bathymetry level (m)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$altitudemax, xy=c(24,5))
    
                # Altitude or Bathymetry (m)
  
                # System (largest)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Largest") %>% pull(GENERAL), xy=c(26,5))
    
                # System (2nd largest)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "2nd largest") %>% pull(GENERAL), xy=c(27,5))
    
                # System (3rd largest)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "3rd largest") %>% pull(GENERAL), xy=c(28,5))
    
                # System (smallest)
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Smallest") %>% pull(GENERAL), xy=c(29,5))
    
                # Rationale for site nomination -text that will be on fact sheets
  nominationRationale <- PF_site %>%
    filter(Field == "Rationale for nomination") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    nominationRationale %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=nominationRationale, xy=c(30,5))
    
                # Site description - text will be on fact sheets
  siteDescription <- PF_site %>%
    filter(Field == "Site description") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    siteDescription %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=siteDescription, xy=c(31,5))
    
                # How is the site managed or potentially manageable?
  siteManagement <- PF_site %>%
    filter(Field == "Conservation") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .) %>%
    {ifelse(. == "None", "There are no known conservation actions at this site.", .)}
    
  for(citationIndex in 1:nrow(PF_citations)){
    siteManagement %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=siteManagement, xy=c(32,5))
  
                # Delineation rationale - text will be on fact sheet
  delineationRationale <- PF_site %>%
    filter(Field == "Delineation rationale") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    delineationRationale %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=delineationRationale, xy=c(33,5))
    
                # Are you attaching a boundary file with this application?
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=PF_site %>% filter(Field == "Boundary file provided?") %>% pull(GENERAL), xy=c(34,5))
    
                # How much of the site is covered by protected areas?
  if(DBS_KBASite$percentprotected < 0.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="0% - completely unprotected", xy=c(35,5))
    
  }else if(DBS_KBASite$percentprotected < 10.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="1-10%", xy=c(35,5))
    
  }else if(DBS_KBASite$percentprotected < 20.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="11-20%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 30.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="21-30%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 40.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="31-40%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 50.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="41-50%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 60.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="51-60%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 70.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="61-70%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 80.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="71-80%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 90.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="81-90%", xy=c(35,5))
      
  }else if(DBS_KBASite$percentprotected < 99.5){
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="91-99%", xy=c(35,5))
      
  }else{
    writeData(multiSiteForm_wb, sheet = "2. Site data", x="100% - completely protected", xy=c(35,5))
  }
    
                # If covered by protected area which is it?
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$protectedareas_en, xy=c(36,5))
    
                # Relationship with protected area boundaries
  if(!reviewedProposal){
    PARelationship <- c("")
      
    if(DBS_KBASite$percentprotected == 0){
      PARelationship <- "KBA boundary does not intersect any PA boundaries"
    
    }else{
        
      PA_intersection <- st_intersection(DBS_KBASite, cpcad) %>%
        st_make_valid() %>%
        select(NAME_E) %>%
        mutate(Areakm2 = as.numeric(st_area(.))/1000000,
               Percent = round(100*Areakm2/DBS_KBASite$areakm2, 1)) %>%
        filter(Percent >= 0.1)
          
      if(nrow(PA_intersection) > 0){
          
        for(PA in 1:nrow(PA_intersection)){
            
          PA_name <- PA_intersection$NAME_E[PA]
            
          PANotSite <- cpcad %>%
            filter(NAME_E == PA_name) %>%
            st_difference(., DBS_KBASite) %>%
            select(NAME_E) %>%
            mutate(Areakm2 = as.numeric(st_area(.))/1000000,
                   Percent = round(100*Areakm2/DBS_KBASite$areakm2, 1)) %>%
            filter(Percent >= 0.1)
          
          SiteNotPA <- cpcad %>%
            filter(NAME_E == PA_name) %>%
            st_difference(DBS_KBASite, .) %>%
            select(NAME_E) %>%
            mutate(Areakm2 = as.numeric(st_area(.))/1000000,
                   Percent = round(100*Areakm2/DBS_KBASite$areakm2, 1)) %>%
            filter(Percent >= 0.1)
          
          if(nrow(PANotSite) + nrow(SiteNotPA) == 0){
            PARelationship <- c(PARelationship, "Identical")
            
          }else if(nrow(PANotSite) == 0){
            PARelationship <- c(PARelationship, "Contains")
            
          }else{
            PARelationship <- c(PARelationship, "Overlaps")
          }
        }
          
        if("Identical" %in% PARelationship){
          PARelationship <- "KBA boundary exactly follows a PA boundary"
          
        }else if("Overlaps" %in% PARelationship){
          PARelationship <- "KBA overlaps one or more PA boundaries"
          
        }else{
          PARelationship <- "KBA contains one or more PA boundaries"
        }
        
      }else{
        stop("Percent protected > 0 and no CPCAD intersection")
      }
    }
      
    writeData(multiSiteForm_wb, sheet = "2. Site data", x=PARelationship, xy=c(37,5))
  }
  
                # Additional biodiversity values at site
  additionalBiodiversity <- PF_site %>%
    filter(Field == "Additional biodiversity") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    additionalBiodiversity %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  additionalBiodiversity %<>% paste(., speciesNnotG_text, ecosystemsNnotG_text) %>% trimws()
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=additionalBiodiversity, xy=c(38,5))
    
                # Does the site include land/water/resources belonging to Indigenous people or subject to customary use rights. If yes name the indigenous group.
  customaryJurisdiction <- PF_site %>%
    filter(Field == "Customary jurisdiction") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    customaryJurisdiction %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=customaryJurisdiction, xy=c(39,5))
    
                # Land-use regimes at site
  landUseRegimes <- PF_site %>%
    filter(Field == "Land-use regimes") %>%
    pull(GENERAL) %>%
    gsub("<i>", "", .) %>%
    gsub("</i>", "", .)
    
  for(citationIndex in 1:nrow(PF_citations)){
    landUseRegimes %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
  }
    
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=landUseRegimes, xy=c(40,5))
    
                # Does the site overlap an OECM. If yes put the name of the OECM here otherwise leave blank.
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=DBS_KBASite$oecmsatsite_en, xy=c(41,5))
    
                # Notes
    
                # Academic, scientific and expert consultation
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=technicalReview_text, xy=c(43,5))
    
                # Consultation with other stakeholders (government, NGOs, local people etc):
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=generalReview_text, xy=c(44,5))
    
                # Potential Reviewers for the site
  writeData(multiSiteForm_wb, sheet = "2. Site data", x=suggestedReviewers, xy=c(45,5))
    
          # 3. Biodiversity elements data
                # If reviewed proposal, determine order in which to enter species data
  if(reviewedProposal){
    
    speciesOrder %<>%
      mutate(order = paste(speciesid, assessmentParameter, sep="_"))
    
    PF_species %<>%
      left_join(., DB_BIOTICS_ELEMENT_NATIONAL[,c("speciesid", "element_code")], by=c("NatureServe Element Code" = "element_code")) %>%
      mutate(order = paste(speciesid, `Assessment parameter`, sep="_")) 
    
    if(sum(!speciesOrder$order %in% PF_species$order) > 0){
      
      commentsInRemovedRows <- speciesOrder %>%
        filter(!order %in% PF_species$order) %>%
        .[,5:22]
      
      if(sum(is.na(commentsInRemovedRows)) == nrow(commentsInRemovedRows)*18){
        
        speciesOrder %<>%
          filter(order %in% PF_species$order)
        
      }else{
        stop("Some species entries with comments not recognized")
      }
      
    }else if(sum(!PF_species$order %in% speciesOrder$order) > 0){
      stop("Some species entries not recognized")
    }
    
    PF_species %<>%
      arrange(match(order, speciesOrder$order)) %>%
      select(-c(speciesid, order))
  }
  
  elementIndices <- c(PF_species$SpeciesIndex, PF_ecosystems$EcosystemIndex, PF_ecologicalIntegrity$EcoregionIndex)
    
  if(length(elementIndices) > 0){
    line <- 5
      
    for(index in elementIndices){
        
                # Get element information
      elementType <- ifelse(grepl("spp", index),
                            "species",
                            ifelse(grepl("ecosys", index),
                                   "ecosystem",
                                   "ecoregion"))
        
      elementNumber <- ifelse(elementType == "species",
                              as.integer(substr(index, start=4, stop=nchar(index))),
                              as.integer(substr(index, start=7, stop=nchar(index))))
        
                # Site record number in WDKBA (sitrecid)
      if(!is.na(WDKBAnumber)){
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=WDKBAnumber, xy=c(1,line))
      }
        
                # Site code if new site
      if(is.na(WDKBAnumber)){
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=CASiteCode, xy=c(2,line))
      }
        
      if(elementType == "species"){
          
                # INTERMEDIATE: Get species identifiers
        SISnumber <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Red List SIS number`)
          
        WDKBAsppnumber <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`WDKBA number`)
        
        elementCode <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`NatureServe Element Code`)
          
        if(is.na(elementCode)){
          stop("Missing Element Code")
        }
          
        sciName <- DB_BIOTICS_ELEMENT_NATIONAL %>%
          filter(element_code == elementCode) %>%
          left_join(., DB_Species[,c("speciesid", "iucn_name")], by="speciesid") %>%
          pull(iucn_name) %>%
          {ifelse(is.na(.), PF_species %>% filter(SpeciesIndex == index) %>% pull(`Scientific name`), .)}
          
                # Species ID Number in SIS (sis_id)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=SISnumber, xy=c(4,line))
          
                # Species ID Number in WDKBA (spcrecid)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=WDKBAsppnumber, xy=c(5,line))
          
                # Your species ID code
        if(is.na(SISnumber) & is.na(WDKBAsppnumber)){
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=elementCode, xy=c(6,line))
        }
          
                # Species (common name)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Common name`), xy=c(7,line))
          
                # Scientific binomial name
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=sciName, xy=c(8,line))
          
                # Taxonomic group
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Taxonomic group`), xy=c(9,line))
          
                # Red List category
        status <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(Status)
        
        statusAssessmentAgency <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Status assessment agency`)
          
        if(!is.na(status)){
            
          if(!statusAssessmentAgency == "IUCN"){
              
            if(PF_species %>% filter(SpeciesIndex == index) %>% pull(`KBA criterion`) == "A1 or B1"){
              
              if(statusAssessmentAgency == "NatureServe"){
                  
                status <- ifelse(status %in% c("GH", "TH", "NH"),
                                 "Critically Endangered (Possibly Extinct)",
                                 ifelse(status %in% c("G1", "T1", "N1"),
                                        "Endangered (EN)",
                                        ifelse(status %in% c("G2", "T2", "N2"),
                                               "Vulnerable (VU)",
                                               "")))
                  
              }else if(statusAssessmentAgency == "COSEWIC"){
                  
                status <- ifelse(status == "E",
                                 "Endangered (EN)",
                                 ifelse(status == "T",
                                        "Vulnerable (VU)",
                                        ""))
              }
            }else{
              status <- NA
            }
          }
          
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=status, xy=c(10,line))
        }
          
                # IUCN Red list criteria
        if(!is.na(status)){
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Assessment criteria`), xy=c(11,line))
        }
          
                # Equivalent system if not IUCN Red List.
        if(!is.na(status)){
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=ifelse((statusAssessmentAgency == "IUCN") | is.na(status), "", statusAssessmentAgency), xy=c(12,line))
        }
          
                # Assess against A1c/A1d?
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_resultsSpecies %>% filter(SpeciesIndex == index) %>% pull(`Eligible for A1c/d`), xy=c(13,line))
          
                # Range-restricted?
          
                # Range restriction determined by
          
                # Eco/bioregion-restricted? B3a, B3b, No
          
                # Eco/bioregion map used
          
                # Name of eco/bioregion
          
                # Is species on KBA list of eco/bioregion restricted species?
          
                # Number of species required to meet B3a/B3b thresholds to trigger KBA status.
          
                # Assessment parameter applied
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Assessment parameter`), xy=c(21,line))
          
                # Min (global)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Min reference estimate`) %>% as.numeric(), xy=c(22,line))
          
                # Best (global)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Best reference estimate`) %>% as.numeric(), xy=c(23,line))
          
                # Max (global)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Max reference estimate`) %>% as.numeric(), xy=c(24,line))
          
                # Source global data
        sourceGlobalEstimate <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Sources of reference estimates`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          sourceGlobalEstimate %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
        
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=sourceGlobalEstimate, xy=c(25,line))
          
                # Notes on global data
        explanationGlobalEstimate <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Explanation of reference estimates`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          explanationGlobalEstimate %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=explanationGlobalEstimate, xy=c(26,line))
          
                # Evidence the species is present at the site
        descriptionOfEvidence <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Description of evidence`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          descriptionOfEvidence %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=descriptionOfEvidence, xy=c(27,line))
          
                # Most recent year for which there is evidence species is present
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Evidence year`), xy=c(28,line))
          
                # Min. number of reproductive units (RU) at site
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Reproductive Units (RU)`), xy=c(29,line))
          
                # What are 10 RUs comprised of
        composition10RUs <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Composition of 10 RUs`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          composition10RUs %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=composition10RUs, xy=c(30,line))
          
                # Source of reproductive unit data (including year of data)
        RUSource <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`RU source`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          RUSource %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=RUSource, xy=c(31,line))
          
                # Min (site)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Min site estimate`) %>% as.numeric(), xy=c(32,line))
          
                # Best (site)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Best site estimate`) %>% as.numeric(), xy=c(33,line))
          
                # Max (site)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Max site estimate`) %>% as.numeric(), xy=c(34,line))
          
                # Derivation of estimate
        derivationOfEstimate <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Derivation of best estimate`) %>%
          ifelse(. == "Other (please add further details in column AA)", "Other (please add further details in Notes (column AU)", .)
        
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=derivationOfEstimate, xy=c(35,line))
          
                # Source of data
        sourceSiteEstimate <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Sources of site estimates`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          sourceSiteEstimate %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=sourceSiteEstimate, xy=c(36,line))
          
                # Year of site population estimates
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Year of site estimate`), xy=c(37,line))
          
                # Are you applying A1,B1,B2,B3 or D1a or D2 or D3
        KBAcriterion <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`KBA criterion`)
          
        value <- ifelse(KBAcriterion == "A1 or B1",
                        "Regularly held  in one or more life cycle stages (A1, B1, B2 or B3)",
                        ifelse(KBAcriterion == "D1",
                               "Aggregation predictably held  in one or more life cycle stages (D1a)",
                               ifelse(KBAcriterion == "D2",
                                      "Supported by site as a refugium (D2)",
                                      "Produced by site as a recruitment source (D3)")))
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=value, xy=c(38,line))
          
                # Seasonal distribution applied to
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_species %>% filter(SpeciesIndex == index) %>% pull(`Seasonal distribution`), xy=c(39,line))
          
                # Reason for selection in column AL if not left as "Regularly held ..."
        oneOf10LargestAgg <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`One of 10 largest aggregations?`)
          
        oneOf10LargestAgg <- ifelse(oneOf10LargestAgg == "Globally", "Yes", "No")
          
        criterionDRationale <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Criterion D rationale`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          criterionDRationale %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        if(oneOf10LargestAgg == "No"){
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=criterionDRationale, xy=c(40,line))
        }
          
                # Globally most important 5% of occupied habitat? (B3c)
          
                # Assessment parameter for globally most important 5% of occupied habitat
          
                # Source for being globally most impt. 5% of occupied habitat
          
                # One of 10 largest aggregations?
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=oneOf10LargestAgg, xy=c(44,line))
          
                # Justification that the species is aggregated
        if(oneOf10LargestAgg == "Yes"){
          writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=criterionDRationale, xy=c(45,line))
        }
          
                # Source for being one of 10 largest aggregations
        source10LargestAggregations <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Source for being one of 10 largest aggregations`)
        
        for(citationIndex in 1:nrow(PF_citations)){
          source10LargestAggregations %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=source10LargestAggregations, xy=c(46,line))
          
                # Notes on site data for species
        explanationSiteEstimate <- PF_species %>%
          filter(SpeciesIndex == index) %>%
          pull(`Explanation of site estimates`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          explanationSiteEstimate %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=explanationSiteEstimate, xy=c(47,line))
      }
        
      if(elementType == "ecosystem"){
          
                # Ecosystem code number
        ecosystemCode <- PF_ecosystems %>%
          filter(EcosystemIndex == index) %>%
          pull(`WDKBA number`)
          
        ecosystemCode <- ifelse(is.na(ecosystemCode), index, ecosystemCode)
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=ecosystemCode, xy=c(48,line))
          
                # Name of ecosystem type
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Name of ecosystem type`), xy=c(49,line))
          
                # Red List of Ecosystems category
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Status in the IUCN Red List of Ecosystems`), xy=c(50,line))
          
                # Global extent
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Reference extent (km2)`) %>% as.numeric(), xy=c(51,line))
          
                # Extent at site Min)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Min site extent (km2)`) %>% as.numeric(), xy=c(52,line))
          
                # Extent at site (best estimate)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Best site extent (km2)`) %>% as.numeric(), xy=c(53,line))
          
                # Extent at site (Max)
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecosystems %>% filter(EcosystemIndex == index) %>% pull(`Max site extent (km2)`) %>% as.numeric(), xy=c(54,line))
          
                # Date of assessment
          
                # Source of ecosystem data
        dataSource <- PF_ecosystems %>%
          filter(EcosystemIndex == index) %>%
          pull(`Data source`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          dataSource %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=dataSource, xy=c(56,line))
      }
        
      if(elementType == "ecoregion"){
          
                # Ecological integrity Ecoregion
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecologicalIntegrity %>% filter(EcoregionIndex == index) %>% pull(`Name of ecoregion`), xy=c(57,line))
          
                # Number of criterion C sites in ecoregion
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=PF_ecologicalIntegrity %>% filter(EcoregionIndex == index) %>% pull(`Number of existing criterion C sites`), xy=c(58,line))
          
                # Ecological Integrity - Evidence for low human impact
        evidenceLowHumanImpact <- PF_ecologicalIntegrity %>%
          filter(EcoregionIndex == index) %>%
          pull(`Evidence of low human impact`)
        
        for(citationIndex in 1:nrow(PF_citations)){
          evidenceLowHumanImpact %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=evidenceLowHumanImpact, xy=c(59,line))
          
                # Ecological integrity - evidence for intact ecological communities
        evidenceIntactEcologicalCommunities <- PF_ecologicalIntegrity %>%
          filter(EcoregionIndex == index) %>%
          pull(`Evidence of intact ecological communities`)
          
        for(citationIndex in 1:nrow(PF_citations)){
          evidenceIntactEcologicalCommunities %<>% gsub(PF_citations$`Short citation`[citationIndex], PF_citations$`Long citation`[citationIndex], .)
        }
          
        writeData(multiSiteForm_wb, sheet = "3. Biodiversity elements data", x=evidenceIntactEcologicalCommunities, xy=c(60,line))
          
                # Date of assessment
      }
      
        # Go to next line
      line <- line + 1
    }
  }else{
    stop(paste0(nationalName, " should be a global KBA, but no trigger elements were found that meet global criteria."))
  }
    
          # 4. Conservation Actions
  actionsOngoing <- PF_actions %>%
    filter(Ongoing == "TRUE") %>%
    pull(Action)
    
  actionsNeeded <- PF_actions %>%
    filter(Needed == "TRUE") %>%
    pull(Action)

  nOngoing <- length(actionsOngoing)
    
  for(actionType in c("Ongoing", "Needed")){
      
    actions <- get(paste0("actions", actionType))
    line <- 3
      
    if(length(actions) > 0){
        
      for(action in actions){
          
                # Site record number in WDKBA (sitrecid)
        if(!is.na(WDKBAnumber)){
          writeData(multiSiteForm_wb, sheet = "4. Conservation Actions", x=WDKBAnumber, xy=c(1,line))
        }
          
                # Site code if new site
        if(is.na(WDKBAnumber)){
          writeData(multiSiteForm_wb, sheet = "4. Conservation Actions", x=CASiteCode, xy=c(2,line))
        }
          
                #	Ongoing conservation Actions at site
        if(actionType == "Ongoing"){
          writeData(multiSiteForm_wb, sheet = "4. Conservation Actions", x=action, xy=c(4,line))
        }
          
                # Conservation actions needed at site
        if(actionType == "Needed"){
          writeData(multiSiteForm_wb, sheet = "4. Conservation Actions", x=action, xy=c(5,line))
        }
          
          # Go to next line
        line <- line + 1
      }
    }
  }
    
          # 5. Habitat types
  if(PF_formVersion <= 1.1){
      
    startRow <- which(PF_site$Field == "Forest")
    habitats <- PF_site[42:nrow(PF_site),] %>%
      drop_na(GENERAL) %>%
      select(-FRENCH)
      
    line <- 3
      
    if(nrow(habitats) > 0){
        
      for(habitat in habitats$Field){
          
                # Site record number in WDKBA (sitrecid)
        if(!is.na(WDKBAnumber)){
          writeData(multiSiteForm_wb, sheet = "5. Habitat types", x=WDKBAnumber, xy=c(1,line))
        }
          
                # Site code if new site
        if(is.na(WDKBAnumber)){
          writeData(multiSiteForm_wb, sheet = "5. Habitat types", x=CASiteCode, xy=c(2,line))
        }
          
                #	Major habitat types at site
        writeData(multiSiteForm_wb, sheet = "5. Habitat types", x=habitat, xy=c(4,line))
          
                # Percentage cover of habitat at site
        writeData(multiSiteForm_wb, sheet = "5. Habitat types", x=habitats %>% filter(Field == habitat) %>% pull(GENERAL), xy=c(5,line))
          
        line <- line + 1
      }
    }
  }
    
          # 6. Threats data
  line <- 4
    
  if(nrow(PF_threats) > 0){
      
    for(index in 1:nrow(PF_threats)){
        
                # Site record number in WDKBA (sitrecid)
      if(!is.na(WDKBAnumber)){
        writeData(multiSiteForm_wb, sheet = "6. Threats data", x=WDKBAnumber, xy=c(1,line))
      }
        
                # Site code if new site
      if(is.na(WDKBAnumber)){
        writeData(multiSiteForm_wb, sheet = "6. Threats data", x=CASiteCode, xy=c(2,line))
      }
        
      if(PF_threats$Category[index] == "Species"){
          
                #	Species ID Number in SIS (sis_id) - indicate a species here if the threat applies to a species otherwise leave blank for threats that apply to the site as a whole
        SISnumber <- PF_threats[index,] %>%
          left_join(., PF_species[,c("Common name", "Red List SIS number")], by=c("Specific biodiversity element" = "Common name")) %>%
          pull(`Red List SIS number`) %>%
          unique()
        
        writeData(multiSiteForm_wb, sheet = "6. Threats data", x=SISnumber, xy=c(4,line))
          
                # Species ID Number in WDKBA (spcrecid)
        WDKBAsppnumber <- PF_threats[index,] %>%
          left_join(., PF_species[,c("Common name", "WDKBA number")], by=c("Specific biodiversity element" = "Common name")) %>%
          pull(`WDKBA number`) %>%
          unique()
          
        writeData(multiSiteForm_wb, sheet = "6. Threats data", x=WDKBAsppnumber, xy=c(5,line))
          
                # Your species ID code
        elementCode <- PF_threats[index,] %>%
          left_join(., PF_species[,c("Common name", "NatureServe Element Code")], by=c("Specific biodiversity element" = "Common name")) %>%
          pull(`NatureServe Element Code`) %>%
          unique()
          
        if(is.na(SISnumber) & is.na(WDKBAsppnumber)){
          writeData(multiSiteForm_wb, sheet = "6. Threats data", x=elementCode, xy=c(6,line))
        }
      }
        
      if(PF_threats$Category[index] == "Ecosystem"){
          
                # Ecosystem code
        ecosystemCode <- PF_threats[index,] %>%
          left_join(., PF_ecosystems[,c("Name of ecosystem type", "WDKBA number", "EcosystemIndex")], by=c("Specific biodiversity element" = "Name of ecosystem type")) %>%
          select(`WDKBA number`, EcosystemIndex) %>%
          distinct()
        
        if(nrow(ecosystemCode) > 1){
          stop("Use case not supported")
        }
          
        ecosystemCode <- ifelse(is.na(ecosystemCode$`WDKBA number`), ecosystemCode$EcosystemIndex, ecosystemCode$`WDKBA number`)          
          
        writeData(multiSiteForm_wb, sheet = "6. Threats data", x=ecosystemCode, xy=c(8,line))
      }
        
                # Level 1 threat
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$`Level 1`[index], xy=c(10,line))
        
                # Level 2 threat
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$`Level 2`[index], xy=c(11,line))
        
                # Level 3 threat
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$`Level 3`[index], xy=c(12,line))
        
                # Timing
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$Timing[index], xy=c(13,line))
        
                # Scope
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$Scope[index], xy=c(14,line))
      
                # Severity
      writeData(multiSiteForm_wb, sheet = "6. Threats data", x=PF_threats$Severity[index], xy=c(15,line))
        
                # Go to next line
      line <- line + 1
    }
  }
   
  # Hide helper sheets
  sheetVisibility(multiSiteForm_wb)[9:10] <- "hidden"
    
  # Return the Global Mutli-Site Form
  return(multiSiteForm_wb)
}

#### KBA-EBAR Database - Load data ####
read_KBAEBARDatabase <- function(datasetNames, type, environmentPath, account, epsg, rangeMapID, ebarCategories, speciesids){

  # Load password and CRS
  if(!missing(environmentPath)){
    load(environmentPath)
  }
  
  # Change CRS if specified in parameters
  if(!missing(epsg)){
    crs <- epsg
  }
  
  # Database parameters
  DBusername <- account
  DBpswd <- get(paste0(account, "_pswd"))
  DBdatasets <- list(c("KBASite", "KBA_View/FeatureServer/0", T),
                     c("SpeciesAtSite", "KBA_View/FeatureServer/7", F),
                     c("Species", "KBA_View/FeatureServer/8", F),
                     c("BIOTICS_ELEMENT_NATIONAL", "Restricted/FeatureServer/4", F),
                     c("SpeciesAssessment", "KBA_View/FeatureServer/14", F),
                     c("PopSizeCitation", "KBA_View/FeatureServer/13", F),
                     c("EcosystemAtSite", "KBA_View/FeatureServer/15", F),
                     c("Ecosystem", "KBA_View/FeatureServer/16", F),
                     c("BIOTICS_ECOSYSTEM", "Restricted/FeatureServer/36", F),
                     c("EcosystemAssessment", "KBA_View/FeatureServer/17", F),
                     c("ExtentCitation", "KBA_View/FeatureServer/18", F),
                     c("KBACitation", "KBA_View/FeatureServer/10", F),
                     c("KBAThreat", "KBA_View/FeatureServer/12", F),
                     c("KBAAction", "KBA_View/FeatureServer/9", F),
                     c("KBALandCover", "KBA_View/FeatureServer/11", F),
                     c("KBAProtectedArea", "KBA_View/FeatureServer/19", F),
                     c("OriginalDelineation", "KBA_View/FeatureServer/3", T),
                     c("BiodivElementDistribution", "KBA_View/FeatureServer/4", T),
                     c("KBACustomPolygon", "KBA_View/FeatureServer/1", T),
                     c("KBAInputPolygon", "KBA_View/FeatureServer/6", F),
                     c("KBAAcceptedSite", "KBA_Accepted_Sites/FeatureServer/0", T),
                     c("KBAInProgressSite", "KBA_InProgress_Sites/FeatureServer/0", T),
                     c("DatasetSource", "Restricted/FeatureServer/5", F),
                     c("InputDataset", "Restricted/FeatureServer/7", F),
                     c("ECCCRangeMap", "Restricted/FeatureServer/2", T),
                     c("COSEWICRangeMap", "Restricted/FeatureServer/2", T),
                     c("RangeMap", "Restricted/FeatureServer/10", F),
                     c("EmptyRangeMap", "Summary/FeatureServer/8", F),
                     c("EBARMap", "EcoshapeRangeMap/FeatureServer/0", T),
                     c("InputPolygonRelToKBAs", "Restricted/FeatureServer/2", T),
                     c("RangeMapEcoshape", "Restricted/FeatureServer/12", F),
                     c("InputPolygon", "Restricted/FeatureServer/2", T),
                     c("InputLine", "Restricted/FeatureServer/1", T),
                     c("InputPoint", "Restricted/FeatureServer/0", T),
                     c("EOPolygon", "EO_Polygons/FeatureServer/0", T),
                     c("ACCDCPoint", "Restricted/FeatureServer/0", T),
                     c("ACCDCPolygon", "Restricted/FeatureServer/2", T),
                     c("FP_KBASite", "FP_KBASite/FeatureServer/0", T),
                     c("FP_SpeciesAtSite", "FP_KBASite/FeatureServer/2", F),
                     c("FP_EcosystemAtSite", "FP_KBASite/FeatureServer/1", F))
  
  # Only retain datasets that are desired
  if(!missing(datasetNames)){
    
    if(type == "include"){
      DBdatasets <- sapply(DBdatasets, function(x) x[1] %in% datasetNames) %>%
        {DBdatasets[which(.)]}
    }
    
    if(type == "exclude"){
      DBdatasets <- sapply(DBdatasets, function(x) !x[1] %in% datasetNames) %>%
        {DBdatasets[which(.)]}
    }
  }

  # Get KBA-EBAR token
  response <- httr::POST("https://gis.natureserve.ca/portal/sharing/rest/generateToken",
                         body = list(username=DBusername, password=DBpswd, referer=":6443/arcgis/admin", f="json"),
                         encode = "form")
  token <- content(response)$token
  
  # Load KBA-EBAR data
  for(i in 1:length(DBdatasets)){
    
    # Parameters
    name <- DBdatasets[[i]][1]
    address <- DBdatasets[[i]][2]
    spatial <- DBdatasets[[i]][3] %>% as.logical()
    
    # Query
    if(name == "ECCCRangeMap"){
      
      query <- "inputdatasetid IN (18612)"
      
    }else if(name == "COSEWICRangeMap"){
      
      query <- DB_InputDataset[which(DB_InputDataset$datasetsourceid == 1120), "inputdatasetid"] %>%
        {paste0("inputdatasetid IN (", paste(., collapse=","), ")")}
      
    }else if(name == "EBARMap"){
      
      query <- paste0("(rangemapid = ", rangeMapID, ") AND (presence in (", paste(paste0("'", ebarCategories, "'"), collapse=","), "))")
      
    }else if(name == "InputPolygonRelToKBAs"){
      
      query <- DB_KBAInputPolygon %>%
        filter(!is.na(inputpolygonid)) %>%
        pull(inputpolygonid) %>%
        unique() %>%
        {paste0("inputpolygonid IN (", paste(., collapse=","), ")")}
      
    }else if(name == "EmptyRangeMap"){
      
      query <- "rangemapid >= 0"
    
    }else if(name == "RangeMapEcoshape"){
      
      query <- paste0("(rangemapid = ", rangeMapID, ") AND (presence IN (", paste(paste0("'", ebarCategories, "'"), collapse=", "), "))")
      
    }else if(name %in% c("InputPolygon", "InputLine", "InputPoint", "EOPolygon")){
      
      query <- paste0("speciesid IN (", paste(speciesids, collapse=","), ")")
      
    }else if(name == "ACCDCPolygon"){
      
      query <- DB_InputDataset[which(DB_InputDataset$datasetsourceid == 828), "inputdatasetid"] %>%
        {paste0("(inputdatasetid IN (", paste(., collapse=","), ")) AND (speciesid IN (", paste(speciesids, collapse=","), "))")}
    
    }else if(name == "ACCDCPoint"){
      
      query <- DB_InputDataset[which(DB_InputDataset$datasetsourceid == 825), "inputdatasetid"] %>%
        {paste0("(inputdatasetid IN (", paste(., collapse=","), ")) AND (speciesid IN (", paste(speciesids, collapse=","), "))")}
      
    }else{
      
      query <- "OBJECTID >= 0"
    }
    
    # Get GeoJSON
    url <- parse_url("https://gis.natureserve.ca/arcgis/rest/services")
    url$path <- paste(url$path, paste0(ifelse(name=="EOPolygon", "", "EBAR-KBA/"), address, "/query"), sep = "/")
    url$query <- list(where = query,
                      outFields = "*",
                      returnGeometry = ifelse(spatial, "true", "false"),
                      f = "geojson")
    request <- build_url(url)
    response <- VERB(verb = "GET",
                     url = request,
                     add_headers(`Authorization` = paste("Bearer ", token)))
    data <- content(response, as="text")
    
    # Check that query limit wasn't exceeded
    if(grepl("exceededTransferLimit", data, fixed=T)){
      stop(paste("Query limit exceeded for", name))
    }
    
    # Convert to sf
    data %<>% geojson_sf()
    
    # If non-spatial, drop geometry
    if(!spatial){
      data %<>% st_drop_geometry()
    
    # If spatial and not the default GeoJSON CRS, transform to the specified CRS
    }else if(!crs == 4326){
      data %<>% st_transform(., crs)
    }
    
    # Format IDs
          # Get columns that hold IDs
    idColumns <- colnames(data) %>%
      .[which(endsWith(., "id"))] %>%
      .[which(!. == "globalid")]
    
          # Format ID columns
    data %<>% mutate_at(.vars = idColumns, .funs = as.integer)
    
    # Format dates
    if("created_date" %in% colnames(data)){
      data %<>% mutate(created_date = as.POSIXct(as.numeric(created_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("last_edited_date" %in% colnames(data)){
      data %<>% mutate(last_edited_date = as.POSIXct(as.numeric(last_edited_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("statuschangedate" %in% colnames(data)){
      data %<>% mutate(statuschangedate = as.POSIXct(as.numeric(statuschangedate)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("confirmdate" %in% colnames(data)){
      data %<>% mutate(confirmdate = as.POSIXct(as.numeric(confirmdate)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("publishdate" %in% colnames(data)){
      data %<>% mutate(publishdate = as.POSIXct(as.numeric(publishdate)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("dateassessed" %in% colnames(data)){
      data %<>% mutate(dateassessed = as.POSIXct(as.numeric(dateassessed)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("iucn_assessmentdate" %in% colnames(data)){
      data %<>% mutate(iucn_assessmentdate = as.POSIXct(as.numeric(iucn_assessmentdate)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("cosewic_date" %in% colnames(data)){
      data %<>% mutate(cosewic_date = as.POSIXct(as.numeric(cosewic_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("g_rank_review_date" %in% colnames(data)){
      data %<>% mutate(g_rank_review_date = as.POSIXct(as.numeric(g_rank_review_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("g_rank_change_date" %in% colnames(data)){
      data %<>% mutate(g_rank_change_date = as.POSIXct(as.numeric(g_rank_change_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("n_rank_review_date" %in% colnames(data)){
      data %<>% mutate(n_rank_review_date = as.POSIXct(as.numeric(n_rank_review_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("n_rank_change_date" %in% colnames(data)){
      data %<>% mutate(n_rank_change_date = as.POSIXct(as.numeric(n_rank_change_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("sara_status_date" %in% colnames(data)){
      data %<>% mutate(sara_status_date = as.POSIXct(as.numeric(sara_status_date)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    if("maxdate" %in% colnames(data)){
      data %<>% mutate(maxdate = as.POSIXct(as.numeric(maxdate)/1000, origin = "1970-01-01", tz = "GMT"))
    }
    
    # Format integers
    if("sitestatus" %in% colnames(data)){
      data %<>% mutate(sitestatus = as.integer(sitestatus))
    }
    
    if("ru_min" %in% colnames(data)){
      data %<>% mutate(ru_min = as.integer(ru_min))
    }
    
    # Format doubles
    if("percentatsite" %in% colnames(data)){
      data %<>% mutate(percentatsite = round(as.numeric(percentatsite), 2))
    }
    
    if("siteestimate_min" %in% colnames(data)){
      data %<>% mutate(siteestimate_min = round(as.numeric(siteestimate_min), 2))
    }
    
    if("siteestimate_best" %in% colnames(data)){
      data %<>% mutate(siteestimate_best = round(as.numeric(siteestimate_best), 2))
    }
    
    if("siteestimate_max" %in% colnames(data)){
      data %<>% mutate(siteestimate_max = round(as.numeric(siteestimate_max), 2))
    }
    
    if("referenceestimate_min" %in% colnames(data)){
      data %<>% mutate(referenceestimate_min = round(as.numeric(referenceestimate_min), 2))
    }
    
    if("referenceestimate_best" %in% colnames(data)){
      data %<>% mutate(referenceestimate_best = round(as.numeric(referenceestimate_best), 2))
    }
    
    if("referenceestimate_max" %in% colnames(data)){
      data %<>% mutate(referenceestimate_max = round(as.numeric(referenceestimate_max), 2))
    }
    
    # Format misc.
    if("ivc_formatted_scientific_name_f" %in% colnames(data)){
      data %<>% rename(ivc_formatted_scientific_name_fr = ivc_formatted_scientific_name_f)
    }
    
    # Rename dataset
    assign(paste0("DB_", name), data, envir = .GlobalEnv)
  }
}

#### KBA-EBAR Database - Filter data for one or several sites ####
filter_KBAEBARDatabase <- function(KBASiteIDs, RMUnfilteredDatasets, datasetNames, inputPrefix, outputPrefix){
  
  # Database parameters
  DBdatasets <- list(c("KBASite", "KBA_View/FeatureServer/0", T),
                     c("SpeciesAtSite", "KBA_View/FeatureServer/7", F),
                     c("Species", "KBA_View/FeatureServer/8", F),
                     c("BIOTICS_ELEMENT_NATIONAL", "Restricted/FeatureServer/4", F),
                     c("SpeciesAssessment", "KBA_View/FeatureServer/14", F),
                     c("PopSizeCitation", "KBA_View/FeatureServer/13", F),
                     c("EcosystemAtSite", "KBA_View/FeatureServer/15", F),
                     c("Ecosystem", "KBA_View/FeatureServer/16", F),
                     c("BIOTICS_ECOSYSTEM", "Restricted/FeatureServer/36", F),
                     c("EcosystemAssessment", "KBA_View/FeatureServer/17", F),
                     c("ExtentCitation", "KBA_View/FeatureServer/18", F),
                     c("KBACitation", "KBA_View/FeatureServer/10", F),
                     c("KBAThreat", "KBA_View/FeatureServer/12", F),
                     c("KBAAction", "KBA_View/FeatureServer/9", F),
                     c("KBALandCover", "KBA_View/FeatureServer/11", F),
                     c("KBAProtectedArea", "KBA_View/FeatureServer/19", F),
                     c("OriginalDelineation", "KBA_View/FeatureServer/3", T),
                     c("BiodivElementDistribution", "KBA_View/FeatureServer/4", T),
                     c("KBACustomPolygon", "KBA_View/FeatureServer/1", T),
                     c("KBAInputPolygon", "KBA_View/FeatureServer/6", T),
                     c("KBAAcceptedSite", "KBA_Accepted_Sites/FeatureServer/0", T),
                     c("KBAInProgressSite", "KBA_InProgress_Sites/FeatureServer/0", T),
                     c("InputPolygonRelToKBAs", "Restricted/FeatureServer/2", T))
  
  # Only retain datasets that are desired
  if(!missing(datasetNames)){
    DBdatasets <- sapply(DBdatasets, function(x) x[1] %in% datasetNames) %>%
      {DBdatasets[which(.)]}
    
  }else{
    DBdatasets <- sapply(DBdatasets, function(x) exists(paste0("DB_", x[1]))) %>%
      {DBdatasets[which(.)]}
  }
  
  # Set prefixes, if not specified manually
  if(missing(inputPrefix)){
    inputPrefix <- "DB"
  }
  
  if(missing(outputPrefix)){
    outputPrefix <- "DBS"
  }
  
  # Filter KBA-EBAR data
  for(i in 1:length(DBdatasets)){
    
    # Parameters
    name <- DBdatasets[[i]][1]
    data <- get(paste0(inputPrefix, "_", name))
    
    # Filtering
    if(nrow(data) > 0){
      
      if(name %in% c("KBASite", "SpeciesAtSite", "EcosystemAtSite", "KBACitation", "KBAThreat", "KBAAction", "KBALandCover", "KBAProtectedArea", "OriginalDelineation", "KBAAcceptedSite", "KBAInProgressSite")){
        data %<>% filter(kbasiteid %in% KBASiteIDs)
      }
      
      if(name %in% c("Species", "BIOTICS_ELEMENT_NATIONAL")){
        data %<>% filter(speciesid %in% SpeciesAtSite$speciesid)
      }
      
      if(name == "SpeciesAssessment"){
        data %<>% filter(speciesatsiteid %in% SpeciesAtSite$speciesatsiteid)
      }
      
      if(name == "PopSizeCitation"){
        data %<>% filter(speciesassessmentid %in% SpeciesAssessment$speciesassessmentid)
      }
      
      if(name %in% c("Ecosystem", "BIOTICS_ECOSYSTEM")){
        data %<>% filter(ecosystemid %in% EcosystemAtSite$ecosystemid)
      }
      
      if(name == "EcosystemAssessment"){
        data %<>% filter(ecosystematsiteid %in% EcosystemAtSite$ecosystematsiteid)
      }
      
      if(name == "ExtentCitation"){
        data %<>% filter(ecosystemassessmentid %in% EcosystemAssessment$ecosystemassessmentid)
      }
      
      if(name == "BiodivElementDistribution"){
        biodivelementdistributionids <- c(SpeciesAtSite$biodivelementdistributionid, EcosystemAtSite$biodivelementdistributionid) %>% .[which(!is.na(.))]
        data %<>% filter(biodivelementdistributionid %in% biodivelementdistributionids)
      }
      
      if(name %in% c("KBAInputPolygon", "KBACustomPolygon")){
        data %<>% filter((speciesatsiteid %in% SpeciesAtSite$speciesatsiteid) | (ecosystematsiteid %in% EcosystemAtSite$ecosystematsiteid))
      }
      
      if(name == "InputPolygonRelToKBAs"){
        data %<>% filter(inputpolygonid %in% KBAInputPolygon$inputpolygonid)
      }
    }
    
    # Rename dataset
    assign(name, data)
    assign(paste0(outputPrefix, "_", name), data, envir = .GlobalEnv)
    
    # Remove unfiltered dataset, if applicable
    if(RMUnfilteredDatasets){
      rm(list=paste0(inputPrefix, "_", name), envir = .GlobalEnv)
    }
  }
}

#### KBA-EBAR Database - Trim dataset ####
trim_KBAEBARDataset <- function(dataset, id){
  
  # Get key information
        # Name of the dataset in the KBA-EBAR database
  datasetName <- gsub("DBS_", "", dataset)
  
        # Data
  data <- get(dataset)
  
        # Columns to retain
  cols <- crosswalk %>%
    filter(Layer_WCSC == datasetName) %>%
    pull(Name_WCSC) %>%
    c(id, .) %>%
    tolower()
  
  # Format data
        # If spatial data, drop geometry
  if("geometry" %in% colnames(data)){
    data %<>% st_drop_geometry()
  }
  
        # Only keep the columns to retain
  data %<>% select(all_of(cols))
  
  # Return data
  return(data)
}

#### KBA-EBAR Database - Update data and identify deletions ####
update_KBAEBARDataset <- function(dataset, id){
 
  # Get datasets
  oldDataset <- get(dataset)
  newDataset <- get(paste0(dataset, "_new"))
  DBDataset <- get(paste0("DB_", dataset))
  
  # Remove records that are duplicated in the old dataset but not in the new dataset
  if(nrow(oldDataset) > 0){
    
        # Get record frequency
    oldFreq <- oldDataset %>%
      group_by(pick(-all_of(id))) %>%
      summarise(freq = n(), .groups="drop")
    
    newFreq <- newDataset %>%
      group_by(pick(-all_of(id))) %>%
      summarise(freq = n(), .groups="drop")
  
        # Find duplicates in the old dataset that are not retained in the new dataset
    extraRecords <- left_join(oldFreq, newFreq, by=cols[which(!cols == id)]) %>%
      mutate(extra = freq.x - freq.y) %>%
      select(-freq.x, -freq.y) %>%
      left_join(oldDataset, ., by=cols[which(!cols == id)]) %>%
      filter(extra > 0) %>%
      mutate(extra = as.integer(extra))
    
        # Delete those duplicates
    if(nrow(extraRecords) > 0){
        
      deletions <- extraRecords %>%
        group_by(pick(-all_of(id))) %>%
        summarise(delete = list(pick(all_of(id)) %>% slice_max(get(id), n = first(extra))), .groups="drop") %>%
        unnest(delete) %>%
        pull(all_of(id)) %>%
        add_row(deletions, Dataset = dataset, ID = .)
        
      oldDataset %<>%
        filter(!get(id) %in% deletions$ID)
    }
  }
  
  # Get primary key
  if(nrow(newDataset) > 0){
    
    # Convert numeric fields to character, in order to ensure accurate join
          # Old dataset
    oldDataset_noNum <- oldDataset
    
    for(col in cols[which(!cols == id)]){
      
      if(oldDataset %>% pull(all_of(col)) %>% is.numeric(.)){
        
        if(nrow(oldDataset) > 0){
          oldDataset_noNum[,col] <- sapply(oldDataset_noNum[,col], function(x) as.character(x))
          
        }else{
          oldDataset_noNum[,col] <- as.character()
        }
      }
    }
    
          # New dataset
    newDataset_noNum <- newDataset
    
    for(col in cols[which(!cols == id)]){
      
      if(newDataset %>% pull(all_of(col)) %>% is.numeric(.)){
        
        newDataset_noNum[,col] <- sapply(newDataset_noNum[,col], function(x) as.character(x))
      }
    }
    
    # Get primary key values
    newDataset_noNum %<>%
      left_join(., oldDataset_noNum, by=cols[which(!cols == id)]) %>%
      mutate({{id}} := get(paste0(id, ".y"))) %>%
      select(all_of(cols))
    
    # Add primary key values to the new dataset
    newDataset[,id] <- newDataset_noNum[,id]
  }
  
  # Detect deletions
  if(sum(!oldDataset[[id]] %in% newDataset[[id]]) > 0){
    deletions %<>% add_row(Dataset = dataset,
                           ID = oldDataset[which(!oldDataset[[id]] %in% newDataset[[id]]), id])
  }
  
  # Add primary key, where missing
        # Get maximum ID already assigned
  maxID_final <- ifelse(exists(paste0(dataset, "_final")),
                        max(get(paste0(dataset, "_final"))[[id]]),
                        0)
  
  maxID_DB <- ifelse(nrow(DBDataset) > 0,
                     max(DBDataset[[id]]),
                     0)
  
  maxID <- max(maxID_final, maxID_DB)
  
        # Assign new IDs
  finalDataset <- newDataset %>%
    mutate({{id}} := replace(get(id), is.na(get(id)), (maxID+1):(maxID+nrow(.[which(is.na(.[[id]])),]))))
  
  # Check final columns
  if((!length(cols) == length(colnames(finalDataset))) || (!sum(sort(cols) == sort(colnames(finalDataset))) == length(cols))){
    stop(paste("Some", dataset, "columns were not correctly processed."))
  }
  
  # Save to the parent environment
  assign(dataset, finalDataset, envir = parent.frame())
  assign("deletions", deletions, envir = parent.frame())
  
  # Add to previous sites
  if(nrow(finalDataset) > 0){
    
    if(exists(paste0(dataset, "_final"))){
      assign(paste0(dataset, "_final"), bind_rows(get(paste0(dataset, "_final")), finalDataset), envir = parent.frame())
      
    }else{
      assign(paste0(dataset, "_final"), finalDataset, envir = parent.frame())
    }
  }
}

#### KBA-EBAR Database - Check primary key uniqueness ####
primaryKey_KBAEBARDataset <- function(dataset, id){
  
  # Get dataset
  data <- get(dataset)
  
  # Get maximum frequency of primary key
  maxFreq <- data[, id] %>%
    table() %>%
    as.data.frame() %>%
    .[, 2] %>%
    max()
  
  # If maximum > 1, return error message
  if(maxFreq > 1){
    error <- T
    message <- paste("The primary key for", dataset, "is not unique.")
    
  }else{
    error <- F
    message <- NA
  }
  
  # Return end parameters
  return(list(error, message))
}

#### Full Site Proposal - Check data validity ####
check_KBADataValidity <- function(final, postTranslation, priorAcceptance, checkURLs){
  
  # Starting parameters
  error <- F
  warning <- NULL
  message <- c()
  
  # KBA CANADA PROPOSAL FORM
        # Check that a national name is provided
  if(length(PF_nationalName) == 0 | nchar(PF_nationalName) == 0 | is.na(PF_nationalName)){
    error <- T
    message <- c(message, "There is no national name entered in the proposal form.")
  }
  
        # Check that a global name is provided, if site is global
  if(!PF_KBALevel == "National"){
    
    globalName <- PF_site %>%
      filter(Field == "International name") %>%
      pull(GENERAL)
    
    if(is.na(globalName)){
      error <- T
      message <- c(message, "This is a global KBA, yet there is no international name entered in the proposal form.")
    }
  }
  
        # Check that a French name is provided, if and only if the site is in NB
  Jurisdiction_PF <- PF_site %>%
    filter(Field == "Province or Territory") %>%
    pull(GENERAL) %>%
    trimws()
  
  if(postTranslation){
    frenchName <- PF_site %>%
      filter(Field == "National name") %>%
      pull(FRENCH)
    
    if(Jurisdiction_PF == "New Brunswick"){
      if(is.na(frenchName)){
        error <- T
        message <- c(message, "The site is in New Brunswick and is missing a Francophone national name.")
      }
    }else{
      if(!is.na(frenchName)){
        error <- T
        message <- c(message, "A Francophone national name is provided even though the site is not in New Brunswick.")
      }
    }
  }
  
        # Check that a site code is provided
  if(final){
    SiteCode_PF <- PF_site %>%
      filter(Field == "Canadian Site Code") %>%
      pull(GENERAL)
    
    if(length(SiteCode_PF) == 0 | nchar(SiteCode_PF) == 0 | is.na(SiteCode_PF)){
      error <- T
      message <- c(message, "There is no site code entered in the proposal form.")
    }
  }
  
        # Check that all triggers are valid taxonomic concepts
  if(!priorAcceptance){
    SpeciesValidity <- DB_BIOTICS_ELEMENT_NATIONAL %>%
      filter(national_scientific_name %in% PF_species$`Scientific name`) %>%
      unique() %>%
      mutate(IsValid = case_when(inactive_ind == "Y" ~ F, is.na(bcd_style_n_rank) ~ T, bcd_style_n_rank == "NSYN" ~ F, .default = T)) %>%
      select(national_scientific_name, IsValid) %>%
      filter(!IsValid)
    
    if(nrow(SpeciesValidity) > 0){
      error <- T
      message <- c(message, paste0("The following triggers are not valid taxonomic concepts: ", paste(SpeciesValidity$national_scientific_name, collapse=", ")))
    }
  }
  
        # Check that SIS numbers are correct
  incorrectSISNumbers <- PF_species %>%
    left_join(., DB_BIOTICS_ELEMENT_NATIONAL[,c("element_code", "speciesid")], by=c("NatureServe Element Code"="element_code")) %>%
    left_join(., DB_Species[,c("speciesid", "iucn_internaltaxonid")], by="speciesid") %>%
    mutate(iucn_internaltaxonid = ifelse(is.na(iucn_internaltaxonid), 0, iucn_internaltaxonid),
           `Red List SIS number` = ifelse(is.na(`Red List SIS number`), 0, `Red List SIS number`)) %>%
    mutate(CorrectSISID = (iucn_internaltaxonid == `Red List SIS number`)) %>%
    filter(!CorrectSISID & (`KBA level` == "Global")) %>%
    mutate(iucn_internaltaxonid = ifelse(iucn_internaltaxonid == 0, NA, iucn_internaltaxonid))
  
  if(final & (nrow(incorrectSISNumbers) > 0)){
    error <- T
    message <- c(message, paste0("Please enter the correct Red List SIS numbers (SPECIES tab) for the following global triggers (SIS numbers in parentheses): ", paste(paste0(paste(incorrectSISNumbers$`Common name`, incorrectSISNumbers$iucn_internaltaxonid, sep=" ("), ")"), collapse="; ")))
  }
  
        # Check that the correct conservation statuses are entered
  if(!priorAcceptance){
    
              # Species
    if(nrow(PF_species) > 0){
    
                    # If the species has a conservation status of level 1 or 2 (whether in the master species list or in the proposal form)
      conservationStatuses <- PF_species %>%
        left_join(., DB_BIOTICS_ELEMENT_NATIONAL[,c("speciesid", "element_code")], by=c("NatureServe Element Code" = "element_code")) %>%
        left_join(., DB_Species[,c("speciesid", "kbatrigger_g_a1_status", "kbatrigger_n_a1_status")], by="speciesid") %>%
        mutate(Status = case_when(Status == "Vulnerable (VU)" ~ "VU",
                                  Status == "Endangered (EN)" ~ "EN",
                                  Status == "Critically Endangered (CR)" ~ "CR",
                                  Status == "Critically Endangered (Possibly Extinct)" ~ "CR",
                                  Status %in% c("E", "T", "GH", "G1", "G2", "TH", "T1", "T2", "NH", "N1", "N2") ~ Status,
                                  .default=NA)) %>%
        rowwise() %>%
        mutate(CorrectGStatus = ifelse(is.na(Status), ifelse(is.na(kbatrigger_g_a1_status), T, F), grepl(paste0(Status, " (", `Status assessment agency`, ")"), kbatrigger_g_a1_status, fixed=T)),
               CorrectNStatus = ifelse(is.na(Status), ifelse(is.na(kbatrigger_n_a1_status), T, F), grepl(paste0(Status, " (", `Status assessment agency`, ")"), kbatrigger_n_a1_status, fixed=T))) %>%
        mutate(CorrectGStatus = ifelse(is.na(CorrectGStatus),
                                       ifelse(is.na(kbatrigger_g_a1_status), T, F),
                                       CorrectGStatus),
               CorrectNStatus = ifelse(is.na(CorrectNStatus),
                                       ifelse(is.na(kbatrigger_n_a1_status), T, F),
                                       CorrectNStatus)) %>%
        mutate(CorrectGStatus = case_when((!`KBA level` == "Global") | (!`KBA criterion` == "A1 or B1") ~ T, .default = CorrectGStatus),
               CorrectNStatus = case_when(!`KBA level` == "National" | (!`KBA criterion` == "A1 or B1") ~ T, .default = CorrectNStatus))
      
      if((sum(!conservationStatuses$CorrectGStatus) + sum(!conservationStatuses$CorrectNStatus)) > 0){
        error <- T
        wrongConservationStatus <- conservationStatuses %>%
          filter(!CorrectGStatus | !CorrectNStatus) %>%
          pull(`Common name`) %>%
          paste(., collapse="; ")
        message <- c(message, paste("There are incorrect conservation statuses provided for:", wrongConservationStatus))
      }
    
                    # Otherwise
      conservationStatuses <- PF_species %>%
        left_join(., DB_BIOTICS_ELEMENT_NATIONAL[,c("speciesid", "element_code", "cosewic_status")], by=c("NatureServe Element Code" = "element_code")) %>%
        left_join(., DB_Species[,c("speciesid", "iucn_cd", "precautionary_g_rank", "precautionary_n_rank")], by="speciesid") %>%
        mutate(Status = case_when(Status %in% c("Vulnerable (VU)", "Endangered (EN)", "Critically Endangered (CR)", "Critically Endangered (Possibly Extinct)", "E", "T", "GH", "G1", "G2", "TH", "T1", "T2", "NH", "N1", "N2") ~ NA,
                                  Status == "Near Threatened (NT)" ~ "NT",
                                  Status == "Least Concern (LC)" ~ "LC",
                                  Status == "Data Deficient (DD)" ~ "DD",
                                  .default=Status),
               CorrectStatus = ifelse(!is.na(Status),
                                       ifelse(`Status assessment agency` == "IUCN",
                                              ifelse(Status == iucn_cd,
                                                     T,
                                                     F),
                                              ifelse(`Status assessment agency` == "COSEWIC",
                                                     ifelse(Status == cosewic_status,
                                                            T,
                                                            F),
                                                     ifelse(Status %in% c(precautionary_g_rank, precautionary_n_rank),
                                                            T,
                                                            F))),
                                       T))
      
      if(sum(!conservationStatuses$CorrectStatus) > 0){
        error <- T
        wrongConservationStatus <- conservationStatuses %>%
          filter(!CorrectStatus) %>%
          pull(`Common name`) %>%
          paste(., collapse="; ")
        message <- c(message, paste("There are incorrect conservation statuses provided for:", wrongConservationStatus))
      }
    }
  
              # Ecosystems - TO ADD
  }
  
        # Check that conservation action section is filled out
  if(!priorAcceptance){
    
    if(sum(PF_actions$Ongoing == "TRUE") == 0){
      error <- T
      message <- c(message, "No `Ongoing` Conservation Actions entered.")
    }
    
    if(sum(PF_actions$Needed == "TRUE") == 0){
      error <- T
      message <- c(message, "No `Needed` Conservation Actions entered.")
    }
    
    actionsNone <- PF_actions %>% filter(Action == "None")
    actionsNotNone <- PF_actions %>% filter(!Action == "None")
    
    if((actionsNone$Ongoing == "TRUE") & (sum(actionsNotNone$Ongoing == "TRUE") > 0)){
      error <- T
      message <- c(message, "`Ongoing` Conservation Actions include 'None' and other actions.")
    }
    
    if((actionsNone$Needed == "TRUE") & (sum(actionsNotNone$Needed == "TRUE") > 0)){
      error <- T
      message <- c(message, "`Needed` Conservation Actions include 'None' and other actions.")
    }
    rm(actionsNone, actionsNotNone)
  }
  
        # Check that level 2 threats are provided
  if(sum(is.na(PF_threats$`Level 2`)) > 0){
    error <- T
    message <- c(message, "In the THREATS tab, some level 2 information is missing.")
  }
  
        # Check that threat levels are congruent with one another
  if(!PF_noThreats){
    
    for(threatRow in 1:nrow(PF_threats)){
      
      # Get levels
      l1 <- PF_threats[threatRow, "Level 1"] %>% sub(" .*", "", .)
      l2 <- PF_threats[threatRow, "Level 2"] %>% sub(" .*", "", .)
      l3 <- PF_threats[threatRow, "Level 3"] %>% sub(" .*", "", .)
      
      # Check congruence
      if(!is.na(l2) && !grepl(l1, l2, fixed=T)){
        error <- T
        message <- c(message, "In the THREATS tab, some level 2 information does not match level 1 information.")
      }
      
      if(!is.na(l3) && !grepl(l2, l3, fixed=T)){
        error <- T
        message <- c(message, "In the THREATS tab, some level 3 information does not match level 2 information.")
      }
    }
  }
  
        # Check that threats are correctly linked to triggers, where applicable
              # Check that 'Specific biodiversity elements' field is correctly populated
                    # If threat applies to the Entire site
  entireSite <- PF_threats %>%
    filter(Category == "Entire site") %>%
    pull(`Specific biodiversity element`) %>%
    unique()
  
  if((length(entireSite) > 0) & (!sum(is.na(entireSite) == length(entireSite)))){
    error <- T
    message <- c(message, "In the THREATS tab, some biodiversity elements are provided for threats that apply to the entire site.")
  }
  
                    # If threat applies to a Species
                          # Check that species name is provided
  speciesOnly <- PF_threats %>%
    filter(Category == "Species") %>%
    pull(`Specific biodiversity element`) %>%
    unique()
  
  if((length(speciesOnly) > 0) & (sum(is.na(speciesOnly)) > 0)){
    error <- T
    message <- c(message, "In the THREATS tab, species names are missing for some threats that apply to a specific Species.")
  }
  
                          # Check that speciesid was found
  speciesOnly <- PF_threats %>%
    filter(Category == "Species") %>%
    pull(speciesid) %>%
    unique()
  
  if((length(speciesOnly) > 0) & (sum(is.na(speciesOnly)) > 0)){
    error <- T
    message <- c(message, "In the THREATS tab, some species could not be matched to a SpeciesID.")
  }
  
                          # Check that the species match a species in the SPECIES tab
  speciesOnly <- PF_threats %>%
    filter(Category == "Species") %>%
    pull(`Specific biodiversity element`) %>%
    unique()
  
  if((length(speciesOnly) > 0) & !(sum(speciesOnly %in% PF_species$`Common name`) == length(speciesOnly))){
    error <- T
    message <- c(message, "In the THREATS tab, some species could not be matched to a species in the SPECIES tab.")
  }
  
                    # If threat applies to an Ecosystem
                          # Check that ecosystem name is provided
  ecosystemOnly <- PF_threats %>%
    filter(Category == "Ecosystem") %>%
    pull(`Specific biodiversity element`) %>%
    unique()
  
  if((length(ecosystemOnly) > 0) & (sum(is.na(ecosystemOnly)) > 0)){
    error <- T
    message <- c(message, "In the THREATS tab, ecosystem names are missing for some threats that apply to a specific Ecosystem.")
  }
  
                          # Check that ecosystemid was found
  ecosystemOnly <- PF_threats %>%
    filter(Category == "Ecosystem") %>%
    pull(ecosystemid) %>%
    unique()
  
  if((length(ecosystemOnly) > 0) & (sum(is.na(ecosystemOnly)) > 0)){
    error <- T
    message <- c(message, "In the THREATS tab, some ecosystems could not be matched to an EcosystemID.")
  }
  
                          # Check that the ecosystems match an ecosystem in the ECOSYSTEM tab
  ecosystemOnly <- PF_threats %>%
    filter(Category == "Ecosystem") %>%
    pull(`Specific biodiversity element`) %>%
    unique()
  
  if((length(ecosystemOnly) > 0) & !(sum(ecosystemOnly %in% PF_ecosystems$`Name of ecosystem type`) == length(ecosystemOnly))){
    error <- T
    message <- c(message, "In the THREATS tab, some ecosystems could not be matched to an ecosystem in the ECOSYSTEMS tab.")
  }
  
        # Check that citation URLs are valid
  if(checkURLs){
    
    # Check that all entries in the URL column are valid URLs 
    urls <- PF_citations %>%
      pull(URL) %>%
      na.omit() %>%
      as.vector()
    
    urlsNotValid <- sapply(urls, function(x) !is_valid_url(x)) %>%
      {urls[which(.)]}
    
    urlsValid <- urls[which(!urls %in% urlsNotValid)]
    
    urlsLive <- sapply(urlsValid, function(x) {
      try1 <- tryCatch({!http_error(x)}, error = function(e) {"error"})
      try2 <- tryCatch({!http_error(x, user_agent("httr"))}, error = function(e) {"error"})
      try3 <- tryCatch({status_code(GET(x)) == 200}, error = function(e) {"error"})
      ifelse(try1 == T | try2 == T | try3 == T, T, F)
      }) %>%
      {urlsValid[which(.)]}
    
    urlsNotLive <- urlsValid[which(!urlsValid %in% urlsLive)]
    
    urlsWithPb <- c(urlsNotValid, urlsNotLive)
    
    if(length(urlsWithPb) > 0){
      warning <- paste("In the CITATIONS tab, the following URLs are not found:", paste(urlsWithPb, collapse = ", "))
    }
    
    # Check that no URLs are entered in the Long Citation column
    longCitation <- PF_citations %>%
      pull(`Long citation`) %>%
      na.omit() %>%
      as.vector()
    
    if(sum(grepl("www.|http", longCitation)) > 0){
      error <- T
      message <- c(message, "In the CITATIONS tab, there are URLs entered in the `Long citation` column")
    }
  }
  
        # Check that short citations match information in the CITATIONS tab
              # Format CITATIONS tab
  KBACitation <- PF_citations %>%
    rename_with(.fn = function(x) gsub(" ", "", tolower(x))) %>%
    mutate(kbacitationid = NA,
           kbasiteid = KBASiteID,
           kbacitationid = 1:nrow(.))
  
              # Species
                    # Site population size
  PopSizeCitation_site <- PF_species %>%
    mutate(popsizecitationid = NA,
           siteestimate_sources = strsplit(`Sources of site estimates`, "; ")) %>%
    filter(!is.na(siteestimate_sources)) %>%
    unnest(siteestimate_sources) %>%
    mutate(siteestimate_sources = trimws(siteestimate_sources)) %>%
    filter(!grepl("personal communication", siteestimate_sources)) %>%
    filter(!grepl("pers. comm.", siteestimate_sources)) %>%
    filter(!grepl("communication personnelle", siteestimate_sources)) %>%
    filter(!grepl("comm. pers.", siteestimate_sources)) %>%
    select(popsizecitationid, siteestimate_sources) %>%
    distinct() %>%
    left_join(., KBACitation[,c("kbacitationid", "shortcitation")], by=c("siteestimate_sources" = "shortcitation")) %>%
    filter(!(is.na(kbacitationid) & grepl("unpublished data", .$siteestimate_sources, fixed=T)))
  
  if(sum(is.na(PopSizeCitation_site$kbacitationid)) > 0){
    error <- T
    message <- c(message, "Some short citations entered in field SiteEstimate_Sources for species do not match any entries in the CITATIONS tab.")
  }
  
                    # Reference population size
  PopSizeCitation_ref <- PF_species %>%
    mutate(popsizecitationid = NA,
           referenceestimate_sources = strsplit(`Sources of reference estimates`, "; ")) %>%
    filter(!is.na(referenceestimate_sources)) %>%
    unnest(referenceestimate_sources) %>%
    mutate(referenceestimate_sources = trimws(referenceestimate_sources)) %>%
    filter(!grepl("personal communication", referenceestimate_sources)) %>%
    filter(!grepl("pers. comm.", referenceestimate_sources)) %>%
    filter(!grepl("communication personnelle", referenceestimate_sources)) %>%
    filter(!grepl("comm. pers.", referenceestimate_sources)) %>%
    select(popsizecitationid, referenceestimate_sources) %>%
    distinct() %>%
    left_join(., KBACitation[,c("kbacitationid", "shortcitation")], by=c("referenceestimate_sources" = "shortcitation")) %>%
    filter(!(is.na(kbacitationid) & grepl("unpublished data", .$referenceestimate_sources, fixed=T)))
  
  if(sum(is.na(PopSizeCitation_ref$kbacitationid)) > 0){
    error <- T
    message <- c(message, "Some short citations entered in field ReferenceEstimate_Sources for species do not match any entries in the CITATIONS tab.")
  }
  
              # Ecosystems - TO DO: Update once site and reference extent sources are in separate fields
  ExtentCitation_site <- PF_ecosystems %>%
    mutate(extentcitationid = NA,
           siteestimate_sources = strsplit(`Data source`, "; ")) %>%
    filter(!is.na(siteestimate_sources)) %>%
    unnest(siteestimate_sources) %>%
    mutate(siteestimate_sources = trimws(siteestimate_sources)) %>%
    filter(!grepl("personal communication", siteestimate_sources)) %>%
    filter(!grepl("pers. comm.", siteestimate_sources)) %>%
    filter(!grepl("communication personnelle", siteestimate_sources)) %>%
    filter(!grepl("comm. pers.", siteestimate_sources)) %>%
    select(extentcitationid, siteestimate_sources) %>%
    distinct() %>%
    left_join(., KBACitation[,c("kbacitationid", "shortcitation")], by=c("siteestimate_sources" = "shortcitation")) %>%
    filter(!(is.na(kbacitationid) & grepl("unpublished data", .$siteestimate_sources, fixed=T)))
  
  if(sum(is.na(ExtentCitation_site$kbacitationid)) > 0){
    error <- T
    message <- c(message, "Some short citations entered in field 'Data source' for ecosystems do not match any entries in the CITATIONS tab.")
  }
  
        # Check that percentage at site is between 0 and 100
  PF_species %<>% mutate(percentatsite = 100*as.double(`Best site estimate`)/as.double(`Best reference estimate`))
  
  if(sum(PF_species$percentatsite < 0, na.rm=T) > 0){
    error <- T
    message <- c(message, "The percentage of the population at the site is less than 0% for some species")
  }   
  
  if(sum(PF_species %>% filter(!is.na(`Criteria met`)) %>% pull(percentatsite) == 0, na.rm=T) > 0){
    error <- T
    message <- c(message, "The percentage of the population at the site is 0% for some species meeting criteria")
  }   
  
  if(sum(PF_species$percentatsite > 100, na.rm=T) > 0){
    error <- T
    message <- c(message, "The percentage of the population at the site is more than 100% for some species")
  }
  
  # KBA-EBAR DATABASE
        # Site record was found in the database
  if(nrow(DBS_KBASite) == 0){
    stop("The site was not found in the database")
  }
  
        # There aren't multiple records with that national name and version in the database
  duplicateSites <- DB_KBASite %>%
    filter((nationalname == PF_nationalName) & (version == PF_siteVersion))
    
  if(nrow(duplicateSites)>1){
    error <- T
    message <- c(message, "There are several records with that National Name and Version in the database")
  }
  
        # There aren't multiple records with that site code and version in the database
  if(final){
    SiteCode_DB <- DBS_KBASite %>%
      pull(sitecode)
    
    duplicateSites <- DB_KBASite %>%
      filter((sitecode == SiteCode_DB) & (version == PF_siteVersion))
    
    if(nrow(duplicateSites)>1){
      error <- T
      message <- c(message, "There are several records with that Site Code and Version in the database")
    }
  }
  
        # Warn if there are any overlaps with accepted sites
              # Binary vector of overlaps
  overlaps <- DB_KBASite %>%
    filter(sitestatus %in% c(6,7,8)) %>% # Only keep accepted sites
    filter(!sitecode == DBS_KBASite$sitecode) %>% # Exclude sites with the same site code
    st_intersection(., DBS_KBASite) %>%
    mutate(overlaparea = st_area(.)) %>%
    st_drop_geometry() %>%
    filter(overlaparea > units::set_units(0.1, m^2)) %>% # Exclude overlaps less than 0.1 m2
    select(sitecode, nationalname) %>%
    mutate(label = paste0(nationalname, " (", sitecode, ")")) %>%
    pull(label) %>%
    suppressWarnings()
  
              # Produce warning
  if(length(overlaps) > 0){
    warning <- c(warning, paste0("WARNING - Site overlaps accepted KBA(s): ", paste0(overlaps, collapse="; ")))
  }
  
        # Boundary generalization
  if(nrow(DBS_OriginalDelineation) > 1){
    error <- T
    message <- c(message, "There are multiple OriginalDelineation records in the database")
  }
  
  if(DBS_KBASite$boundarygeneralization == "2"){
    error <- T
    message <- c(message, "The site boundary needs to be generalized!")
    
  }else if(DBS_KBASite$boundarygeneralization == "3"){
    
    if(nrow(DBS_OriginalDelineation) == 0){
      error <- T
      message <- c(message, "The site boundary is generalized, but there are no OriginalDelineation records in the database")
    }
    
  }else if(nrow(DBS_OriginalDelineation) > 0){
      error <- T
      message <- c(message, "There is an OriginalDelineation record even though the site boundary doesn't require generalization")
  }
  
        # SpeciesAtSite
  if(length(unique(DBS_SpeciesAtSite$speciesid)) < nrow(DBS_SpeciesAtSite)){
    error <- T
    message <- c(message, "There are duplicate SpeciesAtSite records in the database")
  }
  
        # EcosystemAtSite
  if(length(unique(DBS_EcosystemAtSite$ecosystemid)) < nrow(DBS_EcosystemAtSite)){
    error <- T
    message <- c(message, "There are duplicate EcosystemAtSite records in the database")
  }
  
        # MeetsCriteria - Species
  meetsCriteriaSpp <- DBS_SpeciesAtSite %>%
    pull(meetscriteria)
  
  if(sum(is.na(meetsCriteriaSpp)) > 0){
    error <- T
    message <- c(message, "There are some SpeciesAtSite records in the database with MeetsCriteria = NA")
  }
  
        # MeetsCriteria - Ecosystems
  meetsCriteriaEco <- DBS_EcosystemAtSite %>%
    pull(meetscriteria)
  
  if(sum(is.na(meetsCriteriaEco)) > 0){
    error <- T
    message <- c(message, "There are some EcosystemAtSite records in the database with MeetsCriteria = NA")
  }
  
        # MeetsCriteria - All
  meetsCriteria <- c(meetsCriteriaSpp, meetsCriteriaEco)
  
  if(!sum(meetsCriteria == "Y") >= 1){
    error <- T
    message <- c(message, "There are no biodiversity elements that meet criteria")
  }
  
        # PresentAtSite - Species
  presentAtSite <- DBS_SpeciesAtSite %>%
    filter(meetscriteria == "Y") %>%
    pull(presentatsite) %>%
    unique()
  
  if(sum(is.na(presentAtSite)) > 0 || sum(!presentAtSite == "Y") > 0){
    error <- T
    message <- c(message, "Some SpeciesAtSite records that meet criteria are not flagged as PresentAtSite = Yes")
  }
  
        # KBAInputPolygon
              # SpatialScope
  noSpatialScope <- DBS_KBAInputPolygon %>%
    filter(is.na(spatialscope))
  
  if(nrow(noSpatialScope) > 0){
    error <- T
    message <- c(message, "There are KBAInputPolygons with scope = NA")
  }
  
              # Neither InputPolygonID nor RangeMapID
  noRelatedID <- DBS_KBAInputPolygon %>%
    filter(is.na(inputpolygonid) & is.na(rangemapid))
  
  if(nrow(noRelatedID) > 0){
    error <- T
    message <- c(message, "There are KBAInputPolygon records without an InputPolygonID nor a RangeMapID")
  }
  
              # InputPolygonID and RangeMapID
  twoRelatedID <- DBS_KBAInputPolygon %>%
    filter(!is.na(inputpolygonid) & !is.na(rangemapid))
  
  if(nrow(twoRelatedID) > 0){
    error <- T
    message <- c(message, "There are KBAInputPolygon records with both an InputPolygonID and a RangeMapID")
  }
  
  # CONGRUENCE BETWEEN KBA CANADA PROPOSAL FORM AND KBA-EBAR DATABASE
        # Site name
  SiteName_DB <- DBS_KBASite %>%
    pull(nationalname)
  
  SiteName_PF <- PF_site %>%
    filter(Field == "National name") %>%
    pull(GENERAL)
  
  if(!SiteName_DB == SiteName_PF){
    error <- T
    message <- c(message, "There is a mismatch between the site name in the proposal form and in the database.")
  }
  
        # Site code
  if(final){
    if(!SiteCode_DB == SiteCode_PF){
      error <- T
      message <- c(message, "There is a mismatch between the site code in the proposal form and in the database.")
    }
  }
  
        # Site version
  SiteVersion_DB <- DBS_KBASite %>%
    pull(version) %>%
    as.integer() %>%
    ifelse(is.na(.), PF_siteVersion, .)
  
  SiteVersion_PF <- PF_siteVersion %>%
    ifelse(is.na(.), SiteVersion_DB, .)
  
  if(is.na(SiteVersion_DB) & is.na(SiteVersion_PF)){
    error <- T
    message <- c(message, "Site version information must be entered in the database.")
    
  }else if(!SiteVersion_PF == SiteVersion_DB){
    error <- T
    message <- c(message, "There is a mismatch between the site version in the database and in the proposal form.")
  }
  
        # Jurisdiction
  Jurisdiction_DB <- DBS_KBASite %>%
    pull(jurisdiction_en)
  
  if(!Jurisdiction_DB == Jurisdiction_PF){
    error <- T
    message <- c(message, "There is a mismatch between the jurisdiction in the proposal form and in the database.")
  }
  
        # Site area
  SiteArea_PF <- PF_site %>%
    filter(Field == "Site area (km2)") %>%
    pull(GENERAL)
  
  if(DBS_KBASite$boundarygeneralization == "1"){
    SiteArea_DB <- DBS_KBASite %>%
      st_area() %>%
      as.numeric()/1000000
    
  }else{
    SiteArea_DB <- DBS_OriginalDelineation %>%
      st_area() %>%
      as.numeric()/1000000
  }
  
  if(!is.na(SiteArea_PF) && ((!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=2)) & ((str_sub(formatC(round(SiteArea_DB, 2), format='f', digits=2), start=-1) == "0") & (!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=1))) & ((str_sub(formatC(round(SiteArea_DB, 2), format='f', digits=2), start=-2) == "00") & (!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=0))))){
    error <- T
    message <- c(message, "There is a mismatch between the site area provided in the proposal form and that computed from the shape in the database.")
  }
  
        # Latitude and longitude
  LatLon_DB <- DBS_KBASite %>%
    st_drop_geometry() %>%
    select(lat_wgs84, long_wgs84) %>%
    unlist() %>%
    as.vector() %>%
    as.character()
  
  LatLon_PF <- PF_site %>%
    filter(Field %in% c("Latitude (dd.dddd)", "Longitude (ddd.dddd)")) %>%
    pull(GENERAL)
  
  if(!sum(is.na(LatLon_PF)) == 2){
    if(!sum(LatLon_DB == LatLon_PF) == 2){
      error <- T
      message <- c(message, "There is a mismatch between the latitude and longitude provided in the proposal form and that provided in the database.")
    }
  }
  
        # Species
              # Based on NATIONAL_SCIENTIFIC_NAME
                    # All species
  if(!priorAcceptance){
    
    SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
      filter(national_scientific_name %in% PF_species$`Scientific name`) %>%
      pull(speciesid) %>%
      unique()
    
    SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
      pull(speciesid) %>%
      unique()
    
    if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
      error <- T
      message <- c(message, "There is a mismatch between the species in SpeciesAtSite and the species in the proposal form (based on NATIONAL_SCIENTIFIC_NAME).")
    }
  }
  
                    # Species meeting KBA criteria
  if(!priorAcceptance){
    
    SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
      filter(national_scientific_name %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "Scientific name"]) %>%
      pull(speciesid) %>%
      unique()
    
    SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
      filter(meetscriteria == "Y") %>%
      pull(speciesid) %>%
      unique()
    
    if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
      error <- T
      message <- c(message, "There is a mismatch between the species that meet criteria in SpeciesAtSite and those that meet criteria in the proposal form (based on NATIONAL_SCIENTIFIC_NAME).")
    }
  }
  
              # Based on ELEMENT_CODE
                    # All species
  SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
    filter(element_code %in% PF_species$`NatureServe Element Code`) %>%
    pull(speciesid) %>%
    unique()
  
  SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
    pull(speciesid) %>%
    unique()
  
  if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
    error <- T
    message <- c(message, "There is a mismatch between the species in SpeciesAtSite and the species in the proposal form (based on ELEMENT_CODE).")
  }
  
                    # Species meeting KBA criteria
  SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
    filter(element_code %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "NatureServe Element Code"]) %>%
    pull(speciesid) %>%
    unique()
  
  SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
    filter(meetscriteria == "Y") %>%
    pull(speciesid) %>%
    unique()
  
  if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
    error <- T
    message <- c(message, "There is a mismatch between the species that meet criteria in SpeciesAtSite and those that meet criteria in the proposal form (based on ELEMENT_CODE).")
  }
  
              # Based on NATIONAL_ENGL_NAME
                    # All species
  if(!priorAcceptance){
    
    SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
      filter(national_engl_name %in% PF_species$`Common name`) %>%
      pull(speciesid) %>%
      unique()
    
    if(length(SpeciesIDs_PF) > length(unique(PF_species$`Common name`))){
      SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
        filter((national_engl_name %in% PF_species$`Common name`) & (national_scientific_name %in% PF_species$`Scientific name`)) %>%
        pull(speciesid) %>%
        unique()
    }
    
    SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
      pull(speciesid) %>%
      unique()
    
    if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
      error <- T
      message <- c(message, "There is a mismatch between the species in SpeciesAtSite and the species in the proposal form (based on NATIONAL_ENGL_NAME).")
    }
  }
  
                    # Species meeting KBA criteria
  if(!priorAcceptance){
    
    SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
      filter(national_engl_name %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "Common name"]) %>%
      pull(speciesid) %>%
      unique()
    
    if(length(SpeciesIDs_PF) > length(unique(PF_species[which(!is.na(PF_species$`Criteria met`)),"Common name"]))){
      SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
        filter((national_engl_name %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "Common name"]) & (national_scientific_name %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "Scientific name"])) %>%
        pull(speciesid) %>%
        unique()
    }
    
    SpeciesIDs_DB <- DBS_SpeciesAtSite %>%
      filter(meetscriteria == "Y") %>%
      pull(speciesid) %>%
      unique()
    
    if((!sum(SpeciesIDs_PF %in% SpeciesIDs_DB)==length(SpeciesIDs_PF)) | (!sum(SpeciesIDs_DB %in% SpeciesIDs_PF)==length(SpeciesIDs_DB))){
      error <- T
      message <- c(message, "There is a mismatch between the species that meet criteria in SpeciesAtSite and those that meet criteria in the proposal form (based on NATIONAL_ENGL_NAME).")
    }
  }
  
        # Ecosystems
              # Based on CNVC_ENGLISH_NAME
                    # All ecosystems
  EcosystemIDs_PF <- DB_BIOTICS_ECOSYSTEM %>%
    filter(cnvc_english_name %in% PF_ecosystems$`Name of ecosystem type`) %>%
    pull(ecosystemid) %>%
    unique()
  
  EcosystemIDs_DB <- DBS_EcosystemAtSite %>%
    pull(ecosystemid) %>%
    unique()
  
  if((!sum(EcosystemIDs_PF %in% EcosystemIDs_DB)==length(EcosystemIDs_PF)) | (!sum(EcosystemIDs_DB %in% EcosystemIDs_PF)==length(EcosystemIDs_DB))){
    error <- T
    message <- c(message, "There is a mismatch between the ecostystems in EcosystemAtSite and the ecosystems in the proposal form (based on CNVC_ENGLISH_NAME).")
  }
  
                    # Ecosystems meeting KBA criteria
  EcosystemIDs_PF <- DB_BIOTICS_ECOSYSTEM %>%
    filter(cnvc_english_name %in% PF_ecosystems[which(!is.na(PF_ecosystems$`Criteria met`)), "Name of ecosystem type"]) %>%
    pull(ecosystemid) %>%
    unique()
  
  EcosystemIDs_DB <- DBS_EcosystemAtSite %>%
    filter(meetscriteria == "Y") %>%
    pull(ecosystemid) %>%
    unique()
  
  if((!sum(EcosystemIDs_PF %in% EcosystemIDs_DB)==length(EcosystemIDs_PF)) | (!sum(EcosystemIDs_DB %in% EcosystemIDs_PF)==length(EcosystemIDs_DB))){
    error <- T
    message <- c(message, "There is a mismatch between the species that meet criteria in EcosystemAtSite and those that meet criteria in the proposal form (based on CNVC_ENGLISH_NAME).")
  }
  
        # Biodiversity element distributions
  if(PF_formVersion < 1.2){
    
    biodivElementDist_form <- PF_site %>%
      filter(Field == "Species boundary provided?") %>%
      pull(GENERAL) %>%
      {ifelse(is.na(.), "No", .)}
    
    if((!biodivElementDist_form == "No") & (nrow(DBS_BiodivElementDistribution) == 0)){
      error <- T
      message <- c(message, "There is no internal boundary in the database, yet the KBA Canada Proposal Form suggests that there should be.")
    }
    
    if((biodivElementDist_form == "No") & (nrow(DBS_BiodivElementDistribution) > 0)){
      error <- T
      message <- c(message, "There is at least one internal boundary in the database, yet the KBA Canada Proposal Form suggests that there shouldn't be.")
    }
    
  }else if(PF_formVersion == 1.2){
    
    if(nrow(PF_species) > 0){
      
      biodivElementDist_form <- PF_species %>%
        filter(!display_biodivelementdist == "No internal boundary provided") %>%
        pull(`Common name`)
      
      biodivElementDist_DB <- DBS_BiodivElementDistribution %>%
        left_join(., DBS_SpeciesAtSite[,c("biodivelementdistributionid", "speciesid")], by="biodivelementdistributionid") %>%
        left_join(., DBS_EcosystemAtSite[,c("biodivelementdistributionid", "ecosystemid")], by="biodivelementdistributionid") %>%
        filter(!is.na(speciesid) | is.na(ecosystemid)) %>%
        left_join(., DB_BIOTICS_ELEMENT_NATIONAL[,c("speciesid", "national_engl_name")], by="speciesid") %>%
        pull(national_engl_name)
      
      if((length(biodivElementDist_form) > 0) & (sum(!biodivElementDist_form %in% biodivElementDist_DB) > 0)){
        error <- T
        message <- c(message, "The KBA Canada Proposal Form suggests that some taxa should have internal boundaries, yet those boundaries aren't found in the database.")
      }
      
      if((length(biodivElementDist_DB) > 0) & (sum(!biodivElementDist_DB %in% biodivElementDist_form) > 0)){
        error <- T
        message <- c(message, "There are internal boundaries in the database for taxa flagged as 'No internal boundary' in the proposal form.")
      }
    }
    
  }else{
    error <- T
    message <- c(message, "Data validity check for KBA Canada Form version 1.3 not yet implemented. Please contact Chloé.")
  }
  
  # Return end parameters
  return(list(error, message, warning))
}

#### Full Site Proposal - Summarize KBA criteria met ####
summary_KBAcriteria <-  function(prefix, language, referencePath){
  
  # Get crosswalks
        # KBA_Group
  KBAgroups <- read.xlsx(paste0(referencePath, "KBA_Group.xlsx"))
  
        # criterionHeader
  criterionHeader <- read.xlsx(paste0(referencePath, "CriterionHeader.xlsx"))  
  
  # Species
        # Get data
  speciesatsite <- get(paste0(prefix, "_", "SpeciesAtSite"))
  speciesBiotics <- get(paste0(prefix, "_", "BIOTICS_ELEMENT_NATIONAL"))
  
        # Add taxonomic information
  speciesatsite %<>%
    left_join(., speciesBiotics[, c("speciesid", "kba_group", "national_scientific_name")], by="speciesid") %>%
    select(speciesatsiteid, national_scientific_name, kba_group)
  
        # Get criteria met
  speciesatsite %<>%
    left_join(., PF_species[, c("Criteria met", "Scientific name", "KBA level")], by=c("national_scientific_name"="Scientific name")) %>%
    mutate(globalcriteria = as.character(ifelse(`KBA level` == "Global", gsub("g", "", `Criteria met`), NA)),
           nationalcriteria = as.character(ifelse(`KBA level` == "National", gsub("n", "", `Criteria met`), NA))) %>%
    select(speciesatsiteid, kba_group, globalcriteria, nationalcriteria)
  
        # Remove taxonomic group information if sensitive
  if(sensitiveSpp){
    speciesatsite %<>%
      left_join(., sensitiveSpeciesAtSite, by="speciesatsiteid") %>%
      mutate(display_taxonomicgroup = ifelse(is.na(display_taxonomicgroup), "Yes", display_taxonomicgroup)) %>%
      mutate(kba_group = ifelse(display_taxonomicgroup == "No", "Sensitive Species", kba_group)) %>%
      select(-display_taxonomicgroup)
  }
  
        # Remove speciesatsiteid
  speciesatsite %<>% select(-speciesatsiteid)
  
  # Ecosystems
        # Get data
  ecosystematsite <- get(paste0(prefix, "_", "EcosystemAtSite"))
  ecosystemBiotics <- get(paste0(prefix, "_", "BIOTICS_ECOSYSTEM"))
  ecosystems <- get(paste0(prefix, "_", "Ecosystem"))
  
        # Add classification information
  ecosystematsite %<>%
    left_join(., ecosystemBiotics[, c("ecosystemid", "cnvc_english_name")], by="ecosystemid") %>%
    left_join(., ecosystems[, c("ecosystemid", "kba_group")], by="ecosystemid") %>%
    select(cnvc_english_name, kba_group)
  
        # Get criteria met
  ecosystematsite %<>%
    left_join(., PF_ecosystems[, c("Criteria met", "Name of ecosystem type", "KBA level")], by=c("cnvc_english_name"="Name of ecosystem type")) %>%
    mutate(globalcriteria = as.character(ifelse(`KBA level` == "Global", gsub("g", "", `Criteria met`), NA)),
           nationalcriteria = as.character(ifelse(`KBA level` == "National", gsub("n", "", `Criteria met`), NA))) %>%
    select(kba_group, globalcriteria, nationalcriteria)
  
  # All biodiversity elements
  biodivelements <- bind_rows(speciesatsite, ecosystematsite) %>%
    pivot_longer(cols=c("globalcriteria", "nationalcriteria"), names_to="level", values_to="criteriamet") %>%
    mutate(level = gsub("criteria", "", level)) %>%
    separate_rows(criteriamet, sep="; ") %>%
    mutate(criteriamet = substr(criteriamet, start=1, stop=2)) %>%
    distinct() %>%
    drop_na(criteriamet)
  
  # Check that all biodiversity elements have an assigned kba_group
  if(sum(is.na(biodivelements$kba_group)) > 0){
    stop("Some biodiversity elements do not have an assigned kba_group.")
  }
  
  # Translate the kba_group
  if(language == "EN"){
    biodivelements %<>% mutate(kba_group_translated = kba_group)
  }
  
  if(language == "FR"){
    biodivelements %<>%
      left_join(., KBAgroups[, c("KBA_Group_EN", "KBA_Group_FR")], by=c("kba_group" = "KBA_Group_EN")) %>%
      rename(kba_group_translated = "KBA_Group_FR")
  }
  
  if(language == "ES"){
    biodivelements %<>%
      left_join(., KBAgroups[, c("KBA_Group_EN", "KBA_Group_ES")], by=c("kba_group" = "KBA_Group_EN")) %>%
      rename(kba_group_translated = "KBA_Group_ES")
  }
  
  # Sort groups by alphabetical order, for consistency
        # Move sensitive species to the bottom of the group list
  kba_groups_translated <- biodivelements %>%
    mutate(kba_group_translated = tolower(kba_group_translated)) %>%
    pull(kba_group_translated) %>%
    unique() %>%
    sort()
  
  if(sum(grepl("sensi", kba_groups_translated)) > 0){
    kba_groups_translated <- c(kba_groups_translated[which(!grepl("sensi", kba_groups_translated))], kba_groups_translated[which(grepl("sensi", kba_groups_translated))])
  }
  
        # Sort by level, criteria met, and kba_group
  biodivelements %<>%
    mutate(kba_group_translated = tolower(kba_group_translated)) %>%
    arrange(level, criteriamet, factor(kba_group_translated, levels=kba_groups_translated))
  
  # Add criterion header
  biodivelements %<>%
    left_join(., criterionHeader[,c("Criterion", paste0("Criterion_", language))], by=c("criteriamet" = "Criterion")) %>%
    rename(CriterionHeader = paste0("Criterion_", language))
  
  # Initialize final text
  finalText <- c()
  finalText_level <- c()
  
  # Create text
  for(global_national in unique(biodivelements$level)){
    
    for(criterion in biodivelements %>% filter(level == global_national) %>% pull(criteriamet) %>% unique()){
      
      # Get the biodiversity elements that meet the criterion at that level
      biodivelements_criterion <- biodivelements %>%
        filter((level == global_national) & (criteriamet == criterion))
      
      # Get criterion header
      criterionHeader <- biodivelements_criterion %>%
        pull(CriterionHeader) %>%
        unique()
      
      # Final text
      finalText <- c(finalText, paste0(criterion, " - ", criterionHeader, " (", paste(biodivelements_criterion$kba_group_translated, collapse=", "), ")"))
    }
  
    # Concatenate text for that level
    if(language == "EN"){
      text <- paste0(ifelse(global_national == "global", "GLOBAL: ", "NATIONAL: "), paste(finalText, collapse="; "))
    }
    
    if(language == "FR"){
      text <- paste0(ifelse(global_national == "global", "MONDIAL : ", "NATIONAL : "), paste(finalText, collapse="; "))
    }
    
    if(language == "ES"){
      text <- paste0(ifelse(global_national == "global", "GLOBAL: ", "NACIONAL: "), paste(finalText, collapse="; "))
    }
    
    # Final Text
    finalText_level <- c(finalText_level, text)
    
    # Reset finalText
    finalText <- c()
  }
  
  # Concatenate text across all levels
  finalText <- paste0(paste(finalText_level, collapse=". "), ".")

  # Return final text
  return(finalText)
}

#### Miscellaneous - Concatenate an enumeration, in EN, FR, or ES ####
pasteEnumeration <- function(string){
  
  # Get string length
  stringLength <- length(string)
  
  # Process
  if(stringLength == 1){
    finalText <- string
    
  }else if(stringLength == 2){
    
    finalText <- paste(string, collapse=" & ")
    
  }else{
    
    finalText <- paste(paste(string[1:(length(string)-1)], collapse=", "), string[length(string)], sep=" & ")
  }
  
  # Return final text
  return(finalText)
}

#### Miscellaneous - Convert m2 to km2 ####
m2tokm2 <- function(x){
  y <- x/1000000
  return(y)
}

#### Miscellaneous - Convert character codes to special characters ####
specialCharacters <- function(x){
  
  # é
  x <- gsub("&eacute;", "é", x)
  
  # É
  x <- gsub("&Eacute;", "É", x)
  
  # è
  x <- gsub("&egrave;", "è", x)
  
  # È
  x <- gsub("&Egrave;", "È", x)
  
  # ê
  x <- gsub("&ecirc;", "ê", x)
  
  # Ê
  x <- gsub("&Ecirc;", "Ê", x)
  
  # à
  x <- gsub("&agrave;", "à", x)
  
  # À
  x <- gsub("&Agrave;", "À", x)
  
  # â
  x <- gsub("&acirc;", "â", x)
  
  # Â
  x <- gsub("&Acirc;", "Â", x)
  
  # î
  x <- gsub("&icirc;", "î", x)
  
  # Î
  x <- gsub("&Icirc;", "Î", x)
  
  # ï
  x <- gsub("&iuml;", "ï", x)
  
  # Ï
  x <- gsub("&Iuml;", "Ï", x)
  
  # ù
  x <- gsub("&ugrave;", "ù", x)
  
  # Ù
  x <- gsub("&Ugrave;", "Ù", x)
  
  # û
  x <- gsub("&ucirc;", "û", x)
  
  # Û
  x <- gsub("&Ucirc;", "Û", x)
  
  # ô
  x <- gsub("&ocirc;", "ô", x)
  
  # Ô
  x <- gsub("&Ocirc;", "Ô", x)
  
  # œ
  x <- gsub("&oelig;", "œ", x)
  
  # Œ
  x <- gsub("&OElig;", "Œ", x)
  
  # ç
  x <- gsub("&ccedil;", "ç", x)
  
  # Ç
  x <- gsub("&Ccedil;", "Ç", x)
  
  # '
  x <- gsub("&rsquo;", "'", x)
  x <- gsub("&apos;", "'", x)
  
  # -
  x <- gsub("&ndash;", "-", x)
  
  # «
  x <- gsub("&laquo;", "«", x)
  
  # »
  x <- gsub("&raquo;", "»", x)
  
  # 
  x <- gsub("&nbsp;", " ", x)
  
  # Return result
  return(x)
}


#### map_KBASite function for 1. Pre Steering Committee Code ####
map_KBASite <- function(DBS_KBASite,pathway,map_service,map_type,cities) {
  
  # DBS_KBASite = KBA Shapefile (mutated so that name is NNAME and is 'valid')
  # pathway = filename (within the desired folder, doesn't need .pdf)
  # map_service = map_service from basemaps::get_maptypes()
  # map_type = map_type from basemaps::get_maptypes()
  
  #Extract bounding box for the KBA
  bb <- st_bbox(st_transform(DBS_KBASite,crs = 3857)) 
  
  #Extract width, length, and area for determining orientation and expansion
  length <- bb$ymax - bb$ymin
  width <- bb$xmax - bb$xmin
  area <- as.numeric(st_area(DBS_KBASite))
  
  #Determining buffer based on area (arbitrary numbers that looked good with example sites)
  buffer <- ifelse (area > 10^6, 10000, 2000)
  
  #Expanding bounding box by arbitrary buffer, otherwise just maps right to edge
  bb[1:2]<-bb[1:2] - buffer
  bb[3:4]<-bb[3:4] + buffer
  
  #Get cities located within the bounding box
  if(cities){
    
    # Convert bounding box to lat/lon
    bb_WGS <- st_as_sfc(bb) %>%
      st_transform(crs = 4326) %>%
      st_bbox()
    
    # Get largest cities within the bounding box
    citiesData <- GET(paste0("https://geogratis.gc.ca/services/geoname/en/geonames.geojson?category=O&bbox=", paste(bb_WGS, collapse=","))) %>%
      content(., "text") %>%
      geojson_sf() %>%
      st_centroid() %>%
      st_crop(bb_WGS) %>%
      filter(concise %in% c("CITY", "TOWN", "VILG", "HAM", "UNP")) %>%
      mutate(concise = factor(concise, levels = c("CITY", "TOWN", "VILG", "HAM", "UNP"))) %>%
      arrange(concise, desc(relevance)) %>%
      .[1:3,] %>%
      drop_na(name)
  }
  
  #Start writing PDF file
  pdf(paste0(pathway,".pdf"),width = ifelse(length>width,8.5,11), height = ifelse(length>width,11,8.5))
  
  #Make Map for Site
  KBASiteMap <- DBS_KBASite %>%
    st_transform(crs = 3857) %>% # coordinate system needed for base map function
    
    #start plot
    ggplot()+
    
    #basemap (from basemaps::)
    basemap_gglayer(ext = bb,map_service = map_service, map_type = map_type)+
    scale_fill_identity() +
    
    #annotations (scale and north arrow) (from ggspatial::)
    annotation_scale(aes(location = "br"), pad_x = unit(10,"cm"), pad_y = unit(0.1,"cm")) + 
    annotation_north_arrow(aes(location = "br"), pad_x = unit(7.5,"cm"), pad_y = unit(0.1,"cm")) +
    
    #polygon
    geom_sf(linewidth=1, show.legend = FALSE, aes(col="pink",fill=NA)) + #can show legend instead of ggtitle if desired
    
    #title 
    ggtitle(paste(DBS_KBASite$NName, "KBA")) +
    
    #limits, labels, and theme
    ylim(c(bb$ymin,bb$ymax)) + 
    xlim(c(bb$xmin,bb$xmax)) + 
    xlab("") + ylab("") +
    theme_classic() +
    theme(
      plot.title = element_text(size=18,hjust=0.5),
      axis.line = element_blank()
    )
  
  #Add cities
  if(cities){
    KBASiteMap <- KBASiteMap +
      
      geom_sf(data=citiesData, col="lightgrey")+
      geom_text_repel(data = bind_rows(citiesData, DBS_KBASite), aes(label = name, geometry = geometry, fontface = "bold"), stat = "sf_coordinates", min.segment.length = 10, colour = "lightgrey", box.padding = 0.5, label.padding = 1)
  }
  
  #Create inset map
  KBAInset <- 
    ggplot(DBS_KBASite) + 
    geom_prov(colour="black") +
    scale_fill_manual(guide = "none", values=rep("white",13))+
    geom_sf(linewidth=4, show.legend = FALSE, aes(col="pink",fill=NA))+
    theme(
      panel.border = element_rect(colour = "black",fill=NA),
      panel.background = element_blank(),
      panel.grid = element_blank(),
      axis.ticks = element_blank(),
      axis.text = element_blank()
    )
  
  print(
    ggdraw(KBASiteMap) +
      draw_plot(KBAInset,
                height=0.3,
                width=0.3,
                x=0.7,y=0.01)
  )
  
  #End PDF file
  dev.off()
}

#### map_KBASite_new function for website. Will eventually replace old mapping function. ####
map_KBASite_new <- function(
    KBASiteID, # which KBASiteID to map. Required. 
    pathway, # where you want to output the map files. Must end in /. Required. 
    osm_pbf, # an optional parameter to specify a local loading location for Open Street Map data. Saves processing time if on C: instead of G: 
    BM, # whether or not you want to output the simple basemap into the final PDF. Defaults to F.
    IM, # whether or not you want to output the aerial imagery map into the final PDF. Defaults to F. 
    LC, # whether or not you want to output the land cover map into the final PDF. Defaults to T. 
    BD, # whether or not you want to output the biodiversity element distribution map into the final PDF. Only populates if elements are identified. Defaults to T
    CPCAD, # whether or not to include CPCAD in the outputs
    OSM_Water, # whether or not to add OSM lines (water) to LC map
    OSM_Roads, # Whether or not to add OSM lines (roads and highways to LC map
    limit_filesize, # whether or not you want to limit the filesize of outputs through successive reductions in DPI, default is yes
    max_filesize, # the maximum file size in mb
    dpi, # the starting dpi, default is 300
    dpi_reduction_factor, # how many pixels to reduce with each successive time processing the map
    dpi_min, # the minimum acceptable dpi for the final product. 100 is too low. 
    open_pdf # whether or not you want to open the PDF map after it's completed
) {
  # Function Parameter Descriptions and Testing ####
  
  #Extract Start time
  start <- Sys.time()
  
  # Requires that Userparameters is loaded for Google Drive
  if (!exists("googleDrive")) {
    source("User Parameters.r") # detects user and creates userName, sensitiveDir, and googleDrive character strings.
    if(user == "Chlo\u00e9"){userName <- "cdebyser"}
  }
  
  # Check UserParameters successfully loaded
  if (!exists("googleDrive")) {
    stop("KBA Canada Team Google Drive location not detected. Update User Parameters.R.")
  }
  
  # requires site entry
  if(missing(KBASiteID)) (stop("Enter a value for KBASiteID") )
  
  # pathway = filename (within the desired folder, doesn't need .pdf)
  if(missing(pathway)) (stop("Enter a value for pathway ending in / to output PDFs. ") )
  
  # Do you want a simple base map included in the output?
  if(missing(BM)) (BM <- FALSE)
  
  # Do you want an imagery map included in the output?
  if(missing(IM)) (IM <- FALSE)
  
  # Do you want a land cover map included in the output?
  if(missing(LC)) (LC <- TRUE)
  
  # Do you want an biodiversity element distribution map included in the output?
  if(missing(BD)) (BD <- TRUE)
  
  # include cpcad in the outputs?
  if(missing(CPCAD)) (CPCAD <- FALSE)
  
  # Add OSM lines for water courses and roads?? Default to NO - saves space, some poor watercourse data in places
  if(missing(OSM_Water)) (OSM_Water <- FALSE)
  
  # Add OSM lines for roads?? Default is true, fine to include. 
  if(missing(OSM_Roads)) (OSM_Roads <- TRUE)
  
  # To save a vast amount of processing time, stored the .osm.pbf file for North America in this folder. 
  # To update, download from "https://download.geofabrik.de/north-america/canada-latest.osm.pbf"
  #osm_pbf <- "C:/GIS/OSM/canada-260308.osm.pbf" # local for testing, faster processing off Google Drive
  if(missing(osm_pbf)) (
    osm_pbf <- paste0(googleDrive, "/1. Source Datasets/OSM/canada-260308.osm.pbf")
  ) # slow processing on Google Drive, but accessible. 
  
  # Do you want to limit the file size to some limit? Default is YES
  if(missing(limit_filesize)) (limit_filesize <- TRUE)
  
  # What is the maximum filesize acceptable? Default is 2 MB
  if(missing(max_filesize)) (max_filesize <- 2)
  
  # What is the starting DPI you want? Default is 300. 
  if(missing(dpi)) (dpi <- 300)
  
  # How many pixels do you want to reduce the map by each time you process it?
  if(missing(dpi_reduction_factor)) (dpi_reduction_factor <- 50)
  
  # Minimum acceptable dpi for final image
  if(missing(dpi_min)) (dpi_min <- 150)
  
  # Open File logic
  if(missing(open_pdf)) (open_pdf <- FALSE)
  
  # Load Functions (if not already loaded)
  required_functions <- c("read_KBAEBARDatabase", "filter_KBAEBARDatabase")
  for (f in required_functions){
    if(!exists(f)){
      stop(paste("Function", f, "is not loaded."))
      #source_url("https://github.com/chloedebyser/KBA-Public/blob/main/KBA%20Functions.R?raw=TRUE")
    }
  }
  
  # Set Maptiler APIs (provides higher resolution base maps with better info. Can make a free account and use 100,000 pulls per month)
  maptiler_api <- "oetpeah7T8Oicb4nu9ab" #zmoore@wcs.org
  #maptiler_api <- "Axx7yCqgEupUoWmglqPO"
  
  # Admin #### 
  cat("\n\nStarting to create map for KBASiteID ", KBASiteID)
  
  # For loading files with non-English characters (e.g., some site names) and to ensure legend items are in English
  Sys.setlocale("LC_ALL", "English_United States.utf8") 
  
  ## Packages ####
  required_pkgs <- 
    c("tidyverse",
      "sf",
      "magrittr",
      "openxlsx",
      "elevatr",
      "terra",
      "officer",
      "ddpcr",
      "googledrive",
      "httr",
      "geojsonsf",
      "readxl",
      "stringr",
      "devtools",
      "basemaps",
      "ggspatial",
      "canadianmaps",
      "cowplot",
      "grid",
      "gtable",
      "gridExtra",
      "terra",
      "showtext", # for getting KBA comms fonts
      #"ggfx", # for halos on labels (no longer used",
      "png", # for handling pngs
      #"tidyterra", # Has to be run with at least 4.5.2 to use tidyterra
      #"osmdata", # alternative loading format
      "osmextract", # for extracting Open Street Map data
      "geomtextpath", # for adding labels on linear features (no longer used because of messiness",
      "gridtext", # for adding grobs in maps
      "gridSVG", # for adding grobs in maps
      "ggnewscale", # for adding multiple scale_fills onto ggplot
      "pdftools", # for converting the pdf into png and then reprinting to reduce file size.  
      "ggrepel", # for offsetting labels
      "viridisLite", # for colourblind friendly colour palette
      #remotes::install_github("ITSleeds/OSMTools"",
      "OSMtools" # required to save massive processing time in using OSM data. Contains OSMConvert tool in source files. 
    )
  
  # Check each package is loaded
  missing_pkgs <- required_pkgs[
    !vapply(required_pkgs, requireNamespace,
            FUN.VALUE = logical(1),
            quietly = TRUE)
  ]
  
  
  if (length(missing_pkgs) > 0) {
    stop(
      "Missing required packages: ",
      paste(missing_pkgs, collapse = ", ")
    )
  }
  ## Fonts ####
  font_add_google("Open Sans", "opensans")   # loads all weights
  font_add_google("Montserrat", "montserrat")
  #font_add_google("Noto Sans", "noto")
  showtext_auto()
  #
  ## KBA-EBAR database ####
  # DO NOT HAVE TO READ IN FUNCTION IF ALREADY LOADED IN ENV
  required_tables <- c("BIOTICS_ELEMENT_NATIONAL", "Species", "KBASite", "SpeciesAtSite", "EcosystemAtSite", "BiodivElementDistribution", "KBALandCover")
  
  #Load KBA EBAR Database for testing, comment out for function
  #read_KBAEBARDatabase(datasetNames = required_tables,type="include",environmentPath = sensitiveDir,account=userName) %>% suppressWarnings()
  
  # Check each table is loaded
  for (t in required_tables) {
    if(!exists(paste0("DB_", t))) {
      stop(paste0("DB_", t, " was not detected. Run Read_KBABARDatabase to load required tables: ", required_tables))
    }
  }
  
  # Filter KBAEBARDatabase
  filter_KBAEBARDatabase(KBASiteIDs = KBASiteID, RMUnfilteredDatasets = F)
  
  # Is published logic
  published_site <- DBS_KBASite$sitestatus %in% c(6,7,8,9,10)
  
  # Remove sensitive species at site 
  # just in case they're linked to internal polygons that are NOT supposed to be viewed, but are ALSO for a non-Sensitive species
  if(published_site){
    
    # Remove sensitive species
    DBS_SpeciesAtSite %<>%
      filter(
        display_taxonname == "Yes",
        display_assessmentinfo == "Yes"
      )
    
    # Determine which SpeciesAtSite entries need to omit internal polygons
    biodivToKeep <- DBS_SpeciesAtSite %>%
      filter(
        display_biodivelementdist == "Yes" #might want to include other sensitivity filters...
      ) %>%
      pull(biodivelementdistributionid) %>%
      unique(.)
    
  } else { #Unpublished sites don't have these fields filled out
    DBS_SpeciesAtSite %<>%
      filter(
        showonpublicdashboard == "Yes",
      )
    
    # Determine which SpeciesAtSite entries need to omit internal polygons
    biodivToKeep <- DBS_SpeciesAtSite %>%
      pull(biodivelementdistributionid) %>%
      unique(.)
  }
  
  # Remove internal polygons with only sensitive species
  DBS_BiodivElementDistribution %<>%
    filter(biodivelementdistributionid %in% biodivToKeep)
  
  #
  ## Legend Items ####
  # Define legend rows using colours and descriptions from NALCMS 2020 (Custom CSV) 
  legend_items <- 
    read.csv(
      paste0(googleDrive, "/2. Formatted Datasets - Tabular/Map_Legend.csv"),
      encoding = "UTF-8"
    )
  
  names(legend_items)[1] <- "type" # not sure why it's misreading this column name
  
  # Logic to exclude OSM_Water
  if(!OSM_Water){
    legend_items %<>%
      filter(!lcat %in% c("Watercourse"))
  }
  
  #Logic to exclude OSM_Roads
  if(!OSM_Roads){
    legend_items %<>%
      filter(!lcat %in% c("Major Road", "Minor Road", "Railway"))
  }
  
  # Only include CPCAD if desired
  if(!CPCAD){
    legend_items %<>% 
      filter(eng != "Canadian Protected and Conserved Areas Database")
  }
  
  # Add columns from KBA_LandCover (before wrapping*)
  legend_items$perc <- DBS_KBALandCover$percentcover[match(legend_items$eng, DBS_KBALandCover$landcover_en)]
  legend_items$perc <- ifelse(legend_items$perc < 1, "<1", round(legend_items$perc, 0))
  
  # Change type to boxnum so grob in legend will population with number inside
  if(exists("DBS_KBALandCover")){# this line allows both published and unpublished sites to be mapped. 
    legend_items$type <- ifelse (is.na(legend_items$perc), legend_items$type, "boxnum") 
  }
  
  # Wrap text in french and english columns
  wrap_length <- 50 # num char chosen based on Trial and Error, dependent on width of legend columns in Part 3
  wrap_text <- function(x) {x |> strwrap(width = wrap_length) |> paste(collapse = "\n")} 
  legend_items$fr <- sapply(legend_items$fr, wrap_text)
  legend_items$eng <- sapply(legend_items$eng, wrap_text)
  #                        
  # Extract bounding box for the KBA ####
  DBS_KBASite %<>% st_make_valid() %>% # make valid for bounding box function, otherwise can be weird with some sites
    st_transform(crs = 3857) # coordinate system needed for base map function
  
  # Create a fixed-aspect bbox polygon around any sf object
  fixed_aspect_bbox <- function(x, ratio = 5/4.5, pad = 0.05) {
    
    # x = sf object
    # ratio = width / height (e.g., 5/4), manually compared with existing maps
    # pad = proportion of final width/height to pad on each side (0.05 = 5%)
    
    # Stop if not a sf 
    stopifnot(inherits(x, "sf") || inherits(x, "sfc"))
    
    # Take bounding box from sf 'x' object
    bb <- st_bbox(x)
    
    # Current metrics
    w <- as.numeric(bb["xmax"] - bb["xmin"])
    h <- as.numeric(bb["ymax"] - bb["ymin"])
    if (w == 0 && h == 0) stop("Degenerate bbox: zero width and height (are these points with identical coords?)")
    if (w == 0) w <- 1e-9
    if (h == 0) h <- 1e-9
    
    # Current Aspect is width over height
    current_aspect <- w / h
    
    # Determine the center
    cx <- as.numeric(bb["xmin"] + w / 2)
    cy <- as.numeric(bb["ymin"] + h / 2)
    
    # Target width/height to match the requested aspect ratio
    if (current_aspect < ratio) {
      # too tall -> widen
      target_w <- ratio * h
      target_h <- h
    } else {
      # too wide -> grow height
      target_w <- w
      target_h <- w / ratio
    }
    
    # Apply padding as a fraction of the final target size
    target_w <- target_w * (1 + 2 * pad)
    target_h <- target_h * (1 + 2 * pad)
    
    # Create Final Bounding Box Corners at target ratio
    xmin <- cx - target_w / 2
    xmax <- cx + target_w / 2
    ymin <- cy - target_h / 2
    ymax <- cy + target_h / 2
    
    rect <- st_as_sfc(st_bbox(c(xmin = xmin, ymin = ymin, xmax = xmax, ymax = ymax), crs = st_crs(x)))
    st_sf(geometry = rect)
  }
  
  # Running function to create fixed aspect bbox
  bb_poly <- fixed_aspect_bbox(DBS_KBASite)
  bb <- st_bbox(bb_poly)
  
  # Extract width, length, and area for determining orientation and expansion
  length <- bb$ymax - bb$ymin
  width <- bb$xmax - bb$xmin
  area <- as.numeric(st_area(DBS_KBASite))
  
  # Bigsite logic
  big_site <- ifelse(area > 10^10, TRUE, FALSE) # arbitrary threshold
  really_big_site <- ifelse(area > 10^12, TRUE, FALSE)
  
  # Determining buffer based on area (arbitrary numbers that looked good with example sites)
  buffer <- ifelse (big_site, 10000, 2000)
  
  # Expanding bounding box by arbitrary buffer, otherwise doesn't map to edge
  bb_basemap <- bb
  bb_basemap[1:2] <- bb[1:2] - buffer
  bb_basemap[3:4] <- bb[3:4] + buffer
  #
  ## Extract context data in bounding box (osmextract::) ####
  if(LC){ #Only including these context layers in LC map. 
    # osm data did not work, took too long. Simple queries with extract worked better. 
    # Change internet timeout option
    options(timeout = 1000) # for loading a new North America .osm.pbf, but takes forever
    # Set globally before calling oe_read
    Sys.setenv(GDAL_CACHEMAX = "1024") 
    
    # For function show status update
    cat("\nReading context layers from Open Street Map data (", osm_pbf, "). This will take longer for larger sites.")
    if(osm_pbf == paste0(googleDrive, "/1. Source Datasets/OSM/canada-260308.osm.pbf")){
      cat("\nTo save processing time, consider copying /1. Source Datasets/OSM/canada-260308.osm.pbf to your C: drive and specifying the osm_pbf argument with the new location.
          This reduced runtime for site 2443 from 6.9 to 1.7 minutes.")
    } 
    
    # Create a subset of the OSM data for the bounding box using osmconvert (much faster than loading the whole dataset and then subsetting in R, which takes forever)
    bb_coords <- st_transform(bb_poly, 4326) %>% st_bbox() %>% as.vector() %>% paste(collapse = ",")
    osmconvert_path <- file.path(system.file(package = "OSMtools"), "osmconvert.exe") # Get the path to the osmconvert tool hidden inside the R package OSMTools (needs to be installed)
    osm_pbf_subset <- gsub(".osm.pbf", "_subset.pbf", osm_pbf) # creates a new filename for the subset, in the same location as the original
    cmd <- sprintf('"%s" "%s" -b=%s -o="%s"', osmconvert_path, osm_pbf, bb_coords, osm_pbf_subset) # Build the command. Format: osmconvert input -b=coords -o=output
    status_check <- system(cmd)# Run the command through the application to reduce the pbf to the bounding box of the site. 
    
    # If issues, stop and print message
    if (status_check == 0) {
      cat("\nOSM data subset created successfully.")
    } else {
      stop("\nError in creating OSM subset. Make sure OSMtools is installed. Check osmconvert tool and paths.")
    }
    
    # Determine query for lines layer based on options for OSM
    if(OSM_Water & OSM_Roads) {
      osm_lines_query <- "SELECT * FROM lines WHERE 
            highway IN ('motorway','trunk','primary','secondary','tertiary','residential','service','unclassified') OR 
            railway IN ('rail') OR
            waterway IS NOT NULL"
    }
    
    if(OSM_Water & !OSM_Roads) {
      osm_lines_query <- "SELECT * FROM lines WHERE
            waterway IS NOT NULL"
    }
    
    if(!OSM_Water & OSM_Roads) {
      osm_lines_query <- "SELECT * FROM lines WHERE 
            highway IN ('motorway','trunk','primary','secondary','tertiary','residential','service','unclassified') OR 
            railway IN ('rail')"
    }
    
    # Lines layer
    if(OSM_Water | OSM_Roads){ 
      osm_lines <- 
        oe_read(
          file_path = osm_pbf_subset,
          boundary = st_transform(bb_poly, 4326),  # <- polygon in WGS84 ensures server-side clip, otherwise loads whole dataset (which takes forever)
          extra_tags = "ref", 
          layer = "lines",
          query = osm_lines_query
          ,
          force_download = FALSE
        ) 
      
      #Fix for UTF encoding issue in the "ref" column (only an occassional issue sometimes?)
      #osm_lines <- osm_lines %>% mutate(across(where(is.character), ~ iconv(.x, from = "", to = "UTF-8", sub = NA)))
      
      # Sanity Check 
      #if(any(is.na(iconv(osm_lines$waterway, "", "UTF-8")))){print("UTF-8 encoding problem!")} 
      
      # modify (separated because of long processing time in first chunk)
      osm_lines <- osm_lines %>%
        st_transform(.,crs(bb_poly)) %>%
        mutate(
          type = ifelse (highway %in% c('motorway','trunk','primary','secondary'), "Major Road", NA),
          type = ifelse (highway %in% c('tertiary','residential','service','unclassified'), "Minor Road", type),
          type = ifelse (railway %in% c('rail'), "Railway", type),
          type = ifelse (!is.na(waterway), "Watercourse", type),
          
          #Match colours from legend table
          col = legend_items$lcol[ match(type, legend_items$lcat) ],
          linetype = legend_items$lty[ match(type, legend_items$lcat) ],
          linewidth = legend_items$lwd[ match(type, legend_items$lcat) ]
        )
    }
    
    # Points Layer
    osm_points <- oe_read(
      file_path = osm_pbf_subset,
      boundary = st_transform(bb_poly, 4326),  # <- polygon in WGS84 ensures server-side clip, otherwise loads whole dataset (which takes forever)
      extra_tags = "ref", 
      layer = "points",
      query = "SELECT * FROM points WHERE place IN ('city','town','village')", #removed 'hamlet',
      boundary_type = "clipsrc", # ensures server-side clip, otherwise loads whole dataset (which takes forever)
      force_download = FALSE
    ) 
    
    # Modify (separated because of long processing time in first chunk)
    osm_points <- osm_points %>%
      st_transform(.,crs(bb_poly)) %>%
      mutate(
        x = st_coordinates(.)[, 1],
        y = st_coordinates(.)[, 2]
      )
    
    
    # For function print status update
    cat("OSM layer read complete.")
  }  
  #
  ## Extract land cover raster in bounding box ####
  # For function print status update
  if(LC){
    cat("\nProcessing land cover raster. This will take longer for larger sites.\n")
    
    # Load from Google Drive
    if(!exists("landCover")){
      assign(
        "landCover",
        rast(paste0(googleDrive,"/1. Source Datasets/Natural Resources Canada/LandCover2020/landcover-2020-classification-EPSG3857.tif")),
        envir = .GlobalEnv
      )
      cat("Adding 'landCover' in EPSG 3857 to Global Environment")
    }
    
    # Check landCover is in the proper projection (and not preloaded in env for some reason)
    if(!st_crs(landCover)$epsg == 3857){
      stop("landCover is loaded and is not in the proper projection (EPSG 3857), remove it from your environment and rerun the function.")
    }
    
    # Clip to site boundary
    landCoverSite <- terra::crop(landCover, bb_poly) 
    landCoverSite <- terra::mask(landCoverSite, bb_poly)
    
    # Generalize result after mask/crop
    landCoverSite <- as.factor(round(landCoverSite))
    
    # Aggregation to reduce raster size and speed up visualization for large sites
    if(big_site){
      agFact <- ifelse (really_big_site, 4, 2) # arbitrary numbers that looked good with example sites, tweak as needed. Only applies to very large sites.
      landCoverSite <- aggregate(landCoverSite, fact = agFact, fun = "modal")
    }
    
    # Change to st_stars for use in ggplot
    landCoverSite <- st_as_stars(landCoverSite)
    
    # Create out for sites with no landcover
    if(length(landCoverSite$Canada2020) == length(which(is.nan(landCoverSite$Canada2020))) |
       length(landCoverSite$Canada2020) == length(which(is.na(landCoverSite$Canada2020)))){
      plot_LC <- F
    } else {
      plot_LC <- T
    }
    
    # Create vector of colours to use when plotting raster
    LCcols <- c(legend_items$fill)
    names(LCcols) <- c(legend_items$nalcmsID)
    LCcols <- LCcols[-which(LCcols == "")]
    
    # For function print status update
    cat("\nLand cover raster processing complete.\n")
  }
  
  ## Extract CPCAD in bounding box ####
  if(CPCAD){
    
    # Update
    cat("\nProcessing CPCAD\n")
    
    # Load from Google Drive (if not already loaded) into main environment (preprocessed in 11.0.1)
    if(!exists("cpcad")){
      assign(
        "cpcad",
        st_read(paste0(googleDrive,"/1. Source Datasets/Environment and Climate Change Canada/ProtectedConservedArea_2025/ProtectedConservedArea_2025.gdb"), layer = "ProtectedConservedArea_2025_EPSG3857"),
        envir = .GlobalEnv
      )
      cat("Adding 'cpcad' in EPSG 3857 to Global Environment")
    }
    
    # Check cpcad is in the proper projection (and not preloaded in env for some reason)
    if(!st_crs(cpcad)$epsg == 3857){
      stop("cpcad is already loaded and is not in the proper projection (EPSG 3857), remove it from your environment and rerun the function.")
    }
    
    # Crop to site
    cpcad_site <- cpcad %>%
      st_crop(bb) %>%
      mutate(
        type = case_when(
          GOV_TYPE == 1 ~ "National",
          GOV_TYPE == 2 ~ "Provincial",
          !GOV_TYPE %in% c(1,2) ~ "Other"
        ), 
        type = factor(type, levels = c("National", "Provincial", "Other")),
        Shape_Area = as.numeric(st_area(.)),
        label = case_when(
          Shape_Area > 0.05*area ~ NAME_E,
          TRUE ~ NA
        ), 
        label = stringr::str_wrap(label, 20)
      ) %>%
      group_by(type, label) %>%
      summarize(n = n()) %>%
      st_cast("MULTIPOLYGON")
    
    
    # Status Update
    cat("\nCPCAD Processing Complete.\n")
    
  }
  
  # Make Panels For Final Map ####
  # For function print status update
  cat("\nProducing Map Components\n")
  
  ## KBASiteMap_BM (basemap, page 1) ####
  if(BM){ # only create BMagery map if desired. 
    KBASiteMap_BM <- DBS_KBASite %>%
      
      #start plot
      ggplot()+
      
      #basemap (from basemaps::)
      basemap_gglayer(
        ext = bb_basemap, 
        #map_service = "esri",
        #map_type = "world_street_map" ,
        map_service = "maptiler", 
        map_type = "outdoor", 
        map_token = maptiler_api,
        map_res = ifelse(big_site, 1, 1) 
      )+
      scale_fill_identity() +
      
      #use ggnewscale to reset fill (allows multiple fill type [identity and manual])
      ggnewscale::new_scale_fill()
    
    # Add cpcad (if set)
    if(CPCAD){
      KBASiteMap_BM <- KBASiteMap_BM + 
        geom_sf(data = cpcad_site, color = "darkgreen", fill = "lightgreen", linewidth = 0.5, alpha = 0.5)
      #geom_sf(data = cpcad_site, aes(fill = type), linewidth = 0.5, alpha = 0.5) +
      #geom_sf_text(data = cpcad_site, aes(label = label), color = "darkgreen", size = 3, check_overlap = TRUE)
    }
    
    KBASiteMap_BM <- KBASiteMap_BM +
      
      # adding osm layers (removed, only aerial)
      #geom_sf(data = osm_lines, aes(col = col, linewidth = linewidth, linetype = linetype), show.legend = FALSE, alpha = 0.5) + 
      #scale_colour_identity() +
      #scale_linewidth_identity()+
      #scale_linetype_identity()+
      
      # Adding KBA polygon
      # 1) Bottom: thick black solid
      geom_sf(data = DBS_KBASite, color = "black", fill = NA, linewidth = 1.6, lineend = "butt", linejoin = "round") +
      # 2) Top: thinner white dashed (tweak dash style to taste)
      geom_sf(data = DBS_KBASite, color = "white", fill = NA, linewidth = 0.7,
              linetype = "longdash", lineend = "butt", linejoin = "round") +
      
      # Major roads Label
      #with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]), sigma = 1, colour = "white", expand = 2) + #relative reference
      #geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]) + #relative reference
      
      # Cities label
      #with_outer_glow(geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2]), sigma = 1, colour = "white", expand = 2) + #relative reference
      #geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2])+ #relative reference
      
      # Streams and Rivers Label
      #with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Roads"), aes(label = name), vjust = 0, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[7]), sigma = 1, colour = "white", expand = 2) + 
      
      # Add provinces and territories
      #geom_sf(data = PROV, fill = NA, col = "grey", lwd = 0.5, lty = 2) +
      
      # Suppress Axis Labels
      xlab("") + ylab("") +
      
      # SiteCode/ KBASiteID Label
      #ggfx::with_outer_glow(
      geom_sf_text(aes(label = ifelse(!published_site, paste("KBASiteID", kbasiteid), sitecode)),
                   family = "opensans", fontface = "bold",
                   size = 6, colour = "black") +#,
      # colour = "white",  # halo color
      #sigma  = 0,        # blur (0–1 crisp; >1 softer)
      #expand = 2         # halo thickness
      #) +
      
      #annotations (scale and north arrow) (from ggspatial::)
      annotation_scale(aes(location = "bl"), pad_x = unit(2,"cm"), pad_y = unit(0.35,"cm"), style = "ticks", line_width = 2) + 
      annotation_north_arrow(aes(location = "bl"), pad_x = unit(0.5,"cm"), pad_y = unit(0.2,"cm"), height = unit(1, "cm"), width = unit(1, "cm"), style = north_arrow_orienteering) +
      
      #Theme
      theme_void() + #removes graticule lines
      theme(
        plot.title = element_blank(),
        axis.line = element_blank(), 
        panel.background = element_rect(colour = "black", linewidth = 1),
        panel.ontop = TRUE,  # <-- ensures the border (and grid) are drawn on top of geoms
        plot.margin = grid::unit(c(0, 0, 0, 0), "mm")
      )
    
    # Final coordinate zoom (apply after adding all the labels and things to make sure the image is only the bb)
    quiet(KBASiteMap_BM <- KBASiteMap_BM + coord_sf(xlim = c(bb$xmin, bb$xmax), ylim = c(bb$ymin, bb$ymax), expand = F))
    
    #Display 
    KBASiteMap_BM
  }
  #
  
  ## KBASiteMap_IM (imagery, page 2) ####
  if(IM){ # only create Imagery map if desired. 
    KBASiteMap_IM <- DBS_KBASite %>%
      
      #start plot
      ggplot()+
      
      #basemap (from basemaps::)
      basemap_gglayer(
        ext = bb_basemap, 
        #map_service = "esri",
        #map_type = "world_imagery",
        map_service = "maptiler", 
        map_type = "satellite", 
        map_token = maptiler_api,
        map_res = ifelse(big_site, 1, 1)
      )+
      scale_fill_identity() +
      
      #use ggnewscale to reset fill (allows multiple fill type [identity and manual])
      ggnewscale::new_scale_fill() +
      
      # adding osm layers (removed, only aerial)
      #geom_sf(data = osm_lines, aes(col = col, linewidth = linewidth, linetype = linetype), show.legend = FALSE, alpha = 0.5) + 
      #scale_colour_identity() +
      #scale_linewidth_identity()+
      #scale_linetype_identity()+
      
      # Adding KBA polygon
      # 1) Bottom: thick black solid
      geom_sf(data = DBS_KBASite, color = "black", fill = NA, linewidth = 1.6, lineend = "butt", linejoin = "round") +
      # 2) Top: thinner white dashed (tweak dash style to taste)
      geom_sf(data = DBS_KBASite, color = "white", fill = NA, linewidth = 0.7,
              linetype = "longdash", lineend = "butt", linejoin = "round") +
      
      # Major roads Label
      #with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]), sigma = 1, colour = "white", expand = 2) + #relative reference
      #geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]) + #relative reference
      
      # Cities label
      #with_outer_glow(geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2]), sigma = 1, colour = "white", expand = 2) + #relative reference
      #geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2])+ #relative reference
      
      # Streams and Rivers Label
      #with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Roads"), aes(label = name), vjust = 0, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[7]), sigma = 1, colour = "white", expand = 2) + 
      
      # Add provinces and territories
      #geom_sf(data = PROV, fill = NA, col = "grey", lwd = 0.5, lty = 2) +
      
      # Suppress Axis Labels
      xlab("") + ylab("") +
      
      # SiteCode/ KBASiteID Label
      #ggfx::with_outer_glow(
      geom_sf_text(aes(label = ifelse(!published_site, paste("KBASiteID", kbasiteid), sitecode)),
                   family = "opensans", fontface = "bold",
                   size = 6, colour = "white") +#,
      # colour = "white",  # halo color
      #sigma  = 0,        # blur (0–1 crisp; >1 softer)
      #expand = 2         # halo thickness
      #) +
      
      #annotations (scale and north arrow) (from ggspatial::)
      annotation_scale(aes(location = "bl"), pad_x = unit(2,"cm"), pad_y = unit(0.35,"cm"), style = "ticks", line_width = 2, line_col = "white", text_col = "white") + 
      annotation_north_arrow(aes(location = "bl"), pad_x = unit(0.5,"cm"), pad_y = unit(0.2,"cm"), height = unit(1, "cm"), width = unit(1, "cm"), style = north_arrow_orienteering(fill = c("white", "white"), line_col = "black", text_col = "white")) +
      
      #Theme
      theme_void() + #removes graticule lines
      theme(
        plot.title = element_blank(),
        axis.line = element_blank(), 
        panel.background = element_rect(colour = "black", linewidth = 1),
        panel.ontop = TRUE,  # <-- ensures the border (and grid) are drawn on top of geoms
        plot.margin = grid::unit(c(0, 0, 0, 0), "mm")
      )
    
    # Final coordinate zoom (apply after adding all the labels and things to make sure the image is only the bb)
    quiet(KBASiteMap_IM <- KBASiteMap_IM + coord_sf(xlim = c(bb$xmin, bb$xmax), ylim = c(bb$ymin, bb$ymax), expand = F))
    
    #Display 
    KBASiteMap_IM
  }
  #
  ## KBASiteMap_LC (land cover, page 3)####
  if(LC) {
    KBASiteMap_LC <- DBS_KBASite %>%
      
      #start plot
      ggplot()+
      
      #basemap (from basemaps::get_types())
      basemap_gglayer(
        ext = bb_basemap, 
        #map_service = "esri",
        #map_type = "world_street_map",
        map_service = "maptiler", 
        map_type = "basic",
        map_token = maptiler_api,
        map_res = ifelse(big_site, 1, 1)
      )+
      scale_fill_identity() +
      
      #use ggnewscale to reset fill (allows multiple fill type [identity and manual])
      ggnewscale::new_scale_fill()
    
    #plot land cover (can't use this package, tidyterra with R 4.1.1)
    #geom_spatraster(data = landCoverSite, show.legend = FALSE) +
    #scale_fill_manual(values = LCcols, na.value = NA, drop = FALSE) +
    #coord_equal() +
    
    # plot land cover using stars
    if(plot_LC){
      KBASiteMap_LC <- KBASiteMap_LC +
        geom_stars(data = landCoverSite, na.rm = TRUE, show.legend = FALSE) +
        scale_fill_manual(values = LCcols, na.value = NA, drop = FALSE) +
        coord_equal() 
    }
    
    # adding osm layers (if desired)
    if(OSM_Water | OSM_Roads){
      KBASiteMap_LC <- KBASiteMap_LC + geom_sf(data = osm_lines, aes(col = col, linewidth = linewidth, linetype = linetype), show.legend = FALSE) +
        scale_colour_identity() +
        scale_linewidth_identity()+
        scale_linetype_identity()
    }
    
    # Adding other layers
    KBASiteMap_LC <- KBASiteMap_LC +
      
      # Adding KBA polygon
      # 1) Bottom: thick black solid
      geom_sf(data = DBS_KBASite, color = "black", fill = NA, linewidth = 1.6, lineend = "butt", linejoin = "round") +
      # 2) Top: thinner white dashed (tweak dash style to taste)
      geom_sf(data = DBS_KBASite, color = "white", fill = NA, linewidth = 0.7,
              linetype = "longdash", lineend = "butt", linejoin = "round") +
      
      # Major roads Label
      #ggfx::with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]), sigma = 0.4, colour = "white", expand = 0.8) + #relative reference
      #geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = -0.25, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[2]) + #relative reference
      #geom_textsf(data = filter(osm_lines, type == "Major Road"), aes(label = ref), vjust = 0.5, size = 3, lwd = 0, straight = TRUE, color = "white") + #relative reference
      
      # Cities label
      #ggfx::with_outer_glow(geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2]), sigma = 0.4, colour = "white", expand = 0.8) + #relative reference
      #geom_sf_text(data = filter(osm_points), aes(label = name), vjust = -1, size = 3, color = legend_items$labelcol[2])+ #relative reference
      geom_point(
        data = osm_points,
        aes(x = x, y = y),
        size = 1,
        color = "#000000ff" 
      ) + 
      geom_text_repel(
        data = osm_points,
        aes(x = x, y = y, label = name),
        size = 3,
        family = "montserrat",
        color = "#000000ff",
        box.padding = 0.2,
        point.padding = 0.1,
        min.segment.length = Inf   # always draw leader lines
      ) +
      
      #geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = "white")+ #relative reference
      #ggfx::with_shadow(geom_sf_text(data = filter(osm_points), aes(label = name), vjust = 0.5, size = 3, color = legend_items$labelcol[2]), sigma = 0.4, colour = "white") + #relative reference
      
      # Streams and Rivers Label
      #with_outer_glow(geom_textsf(data = filter(osm_lines, type == "Major Roads"), aes(label = name), vjust = 0, size = 3, lwd = 0, straight = TRUE, color = legend_items$labelcol[7]), sigma = 1, colour = "white", expand = 2) + 
      
      # Add provinces and territories
      #geom_sf(data = PROV, fill = NA, col = "grey", lwd = 0.5, lty = 2) +
      
      # Suppress Axis Labels
      xlab("") + ylab("") +
      
      # SiteCode/ KBASiteID Label
      #ggfx::with_outer_glow(
      geom_sf_text(aes(label = ifelse(published_site, sitecode, paste("KBASiteID", kbasiteid))),
                   family = "opensans", fontface = "bold",
                   size = 6, colour = "black") + #,
      #colour = "white",  # halo color
      #sigma  = 0,        # blur (0–1 crisp; >1 softer)
      #expand = 0.8         # halo thickness
      #) +
      
      #annotations (scale and north arrow) (from ggspatial::)
      annotation_scale(aes(location = "bl"), pad_x = unit(2,"cm"), pad_y = unit(0.35,"cm"), style = "ticks", line_width = 2) + 
      annotation_north_arrow(aes(location = "bl"), pad_x = unit(0.5,"cm"), pad_y = unit(0.2,"cm"), height = unit(1, "cm"), width = unit(1, "cm"), style = north_arrow_orienteering) +
      
      #Theme
      theme_void() + #removes graticule lines
      theme(
        plot.title = element_blank(),
        axis.line = element_blank(), 
        panel.background = element_rect(colour = "black", linewidth = 1),
        panel.ontop = TRUE,  # <-- ensures the border (and grid) are drawn on top of geoms
        plot.margin = grid::unit(c(0, 0, 0, 0), "mm")
      )
    
    # Final coordinate zoom (apply after adding all the labels and things to make sure the image is only the bb)
    quiet(KBASiteMap_LC <- KBASiteMap_LC + coord_sf(xlim = c(bb$xmin, bb$xmax), ylim = c(bb$ymin, bb$ymax), expand = F))
    
    #Display 
    KBASiteMap_LC
    
  }
  #
  ## KBASiteMap_BD (biodiversity element dist, page 4)####
  #Only make the map if there are biodiv element dist entries
  if(nrow(DBS_BiodivElementDistribution) > 0 & BD){
    
    # Create palate based on the number of biodiv elements in site. Just uses landcover colours for simplicity
    #bd_fills <- filter(legend_items, !is.na(nalcmsID))$fill # max 15
    #bd_fills <- bd_fills[1:nrow(DBS_BiodivElementDistribution)] # filter for number of total bd entries
    #DBS_BiodivElementDistribution$fill <- bd_fills 
    # Determined by Birds Canada staff to be not 'visually appealing.'
    
    # Create colour blind friendly palatte, filter for number of total bd entries
    bd_fills <- viridisLite::turbo(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::inferno(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::plasma(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::rocket(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::magma(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::mako(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    #bd_fills <- viridisLite::cividis(nrow(DBS_BiodivElementDistribution)) # filter for number of total bd entries
    
    DBS_BiodivElementDistribution$fill <- bd_fills
    
    # Create map
    KBASiteMap_BD <- DBS_KBASite %>%
      
      #start plot
      ggplot()+
      
      #basemap (from basemaps::get_types())
      basemap_gglayer(
        ext = bb_basemap, 
        map_service = "maptiler", 
        map_type = "basic",
        map_token = maptiler_api,
        map_res = ifelse(big_site, 1, 1)
      )+
      
      # add biodiv element dist polygons
      geom_sf(data = DBS_BiodivElementDistribution, aes(fill = fill), col = "black", show.legend = FALSE, alpha = 0.5) + 
      
      # both above use identity for fill
      scale_fill_identity() +
      
      # Adding KBA polygon
      # 1) Bottom: thick black solid
      geom_sf(data = DBS_KBASite, color = "black", fill = NA, linewidth = 1.6, lineend = "butt", linejoin = "round") +
      # 2) Top: thinner white dashed (tweak dash style to taste)
      geom_sf(data = DBS_KBASite, color = "white", fill = NA, linewidth = 0.7,
              linetype = "longdash", lineend = "butt", linejoin = "round") +
      
      # Suppress Axis Labels
      xlab("") + ylab("") +
      
      # SiteCode/ KBASiteID Label
      #ggfx::with_outer_glow(
      geom_sf_text(aes(label = ifelse(!published_site, paste("KBASiteID", kbasiteid), sitecode)),
                   family = "opensans", fontface = "bold",
                   size = 6, colour = "black", alpha = 1) +#,
      #colour = "white",  # halo color
      #sigma  = 0,        # blur (0–1 crisp; >1 softer)
      #expand = 2         # halo thickness
      #) +
      
      #annotations (scale and north arrow) (from ggspatial::)
      annotation_scale(aes(location = "bl"), pad_x = unit(2,"cm"), pad_y = unit(0.35,"cm"), style = "ticks", line_width = 2) + 
      annotation_north_arrow(aes(location = "bl"), pad_x = unit(0.5,"cm"), pad_y = unit(0.2,"cm"), height = unit(1, "cm"), width = unit(1, "cm"), style = north_arrow_orienteering) +
      
      #Theme
      theme_void() + #removes graticule lines
      theme(
        plot.title = element_blank(),
        axis.line = element_blank(), 
        panel.background = element_rect(colour = "black", linewidth = 1),
        panel.ontop = TRUE,  # <-- ensures the border (and grid) are drawn on top of geoms
        plot.margin = grid::unit(c(0, 0, 0, 0), "mm")
      )
    
    # Final coordinate zoom (apply after adding all the labels and things to make sure the image is only the bb)
    quiet(KBASiteMap_BD <- KBASiteMap_BD + coord_sf(xlim = c(bb$xmin, bb$xmax), ylim = c(bb$ymin, bb$ymax), expand = F))
    
    #Display 
    KBASiteMap_BD
  }
  #
  ## Inset map (MiddleTop LH) ####
  
  # filter for KBA province
  prov_sub <- 
    filter(PROV, PRENAME == DBS_KBASite$jurisdiction_en) %>% # FIX may also need adjoining for sites overlapping multiple jurisdictions
    st_transform(.,3857) %>% # Set CRS to match 
    st_make_valid()
  
  # Extract bounding box for inset map
  bb_prov <- st_bbox(prov_sub)
  bb_prov_poly <- st_as_sfc(bb_prov) %>% st_as_sf()
  
  # Expanding bounding box for white mask by arbitrary buffer to deal with rasters over bounding box edge
  buffer <- 100000 # trial and error
  bb_mask <- bb_prov
  bb_mask[1:2] <- bb_mask[1:2] - buffer
  bb_mask[3:4] <- bb_mask[3:4] + buffer
  bb_mask_poly <- st_as_sfc(bb_mask) %>% st_as_sf() %>% st_make_valid()
  
  # Create the outside_mask: prov bounding box + arbitrary buffer - rectangle minus province
  quiet(outside_mask <- st_difference(bb_mask_poly, prov_sub))
  
  # Inset Map
  KBAInset <- 
    
    #Start Plot with KBA bounding box
    ggplot(bb_poly) +
    
    # Add basemap (not an option in function selection, but could be)
    #basemap_gglayer(ext = bb_prov, map_service = "osm", map_type = "topographic") +
    #basemap (from basemaps::get_types())
    basemap_gglayer(
      ext = bb_prov, 
      map_service = "esri", 
      map_type = "world_shaded_relief",
      #map_token = maptiler_api, 
      map_res = 1
    )+
    
    scale_fill_identity() + # needed for basemap. 
    
    # Draw mask to hide basemap tiles outside the province
    geom_sf(data = outside_mask, inherit.aes = FALSE, color = NA, fill = "white") +
    
    # Province data (data from PROV in canadianmaps::) 
    #geom_prov(col = "black", fill = "white")+
    geom_sf(data = prov_sub, fill = NA, col = "#686868ff", lwd = 0.1) +
    
    # Make sure coord_sf uses the same CRS as your working layers and set limits
    coord_sf(crs = 3857, xlim = c(bb_prov$xmin, bb_prov$xmax), ylim = c(bb_prov$ymin, bb_prov$ymax), expand = F) +
    
    # Visualize bb_poly
    geom_sf(linewidth=1.5, show.legend = FALSE, col="#2F4858",fill=NA) + 
    
    #Labels
    xlab("") + ylab("") +
    
    # Set theme elements
    theme(
      panel.border = element_blank(),
      panel.background = element_blank(),
      panel.grid = element_blank(),
      axis.ticks = element_blank(),
      axis.text = element_blank()
    )
  
  #Display
  KBAInset
  #
  ## Logo and Name (MiddleTop RH) ####
  # KBA Logo 
  KBALogo <- paste0(googleDrive,"/7. Communications/0_KBA logos and brand guidelines/Canada KBA logos/plural (default)/Canada_KBA_EN_plural_darktext_crop.png")
  
  # Plot KBA Logo from Source Path 
  t_eng <- "Key Biodiversity Areas"
  t_fr <- "Zones cl\u00E9s pour la \nbiodiversit\u00E9"
  
  # Read image and wrap as a grob
  KBALogoR <- rasterGrob(readPNG(KBALogo), interpolate = TRUE)
  
  # Make a 3-row, 1-column gtable
  # heights: image row gets more space; tweak as needed
  KBACorner <- gtable(
    widths  = unit(1, "null"),
    heights = unit.c(unit(3, "cm"), unit(1, "cm"), unit(0.5, "cm"))
  )
  
  # Build text grobs (centered; change gp for styling)
  tg_eng <- textGrob(t_eng, x = 0.5, y = 0.5, just = "centre",
                     gp = gpar(fontface = "bold", cex = 1))
  tg_fr <- textGrob(t_fr, x = 0.5, y = 0.5, just = "centre",
                    gp = gpar(fontface = "italic", family = "rift", col = "#666666", fontsize = 12))
  
  
  # Place grobs: image in row 1, text lines in rows 2 and 3
  KBACorner <- gtable_add_grob(KBACorner, list(KBALogoR), t = 1, l = 1)
  #KBACorner <- gtable_add_grob(KBACorner, list(tg_eng), t = 2, l = 1)
  KBACorner <- gtable_add_grob(KBACorner, list(tg_fr), t = 2, l = 1)
  
  # (Optional) add padding around the whole table
  KBACorner <- gtable_add_padding(KBACorner, padding = unit(c(3, 3, 3, 3), "mm"))
  
  # Draw
  grid.newpage(); grid.draw(KBACorner)
  
  
  #
  ## Legend grob functions for symbols and vectors ####
  
  # Fill swatch grob function (for boxes in legend)
  grob_fill <- function(fill, border = "black", alpha = 1) {
    grobTree(
      rectGrob(
        gp = gpar(fill = fill, col = border, lwd = 1.2, alpha = alpha), 
        width = unit(1, "npc"), 
        height = unit(1, "npc")
      )
    )
  }
  
  # Line symbol grob function (For lines)
  grob_line <- function(col = "black", lty = 1, lwd = 2) {
    grobTree(
      linesGrob(x = unit(c(0.1, 0.9), "npc"),
                y = unit(c(0.5, 0.5), "npc"),
                gp = gpar(col = col, lwd = lwd, lty = lty))
    )
  }
  
  # Number-in-box grob (for class codes like <1, 1, 3, 89 in boxes) function
  grob_boxnum <- function(text, fill = "white", col = "black") {
    grobTree(
      rectGrob(
        gp = gpar(fill = fill, col = col, lwd = 1.2), 
        width = unit(1, "npc"), 
        height = unit(1, "npc")
      ),
      textGrob(
        label = text, 
        gp = gpar(col = "black", fontsize = ts_boxnum, fontface = "bold", family = "opensans")
      )
    )
  }
  
  # KBA alternating black and white bands
  grob_kba <- function() {
    grobTree(
      rectGrob(gp = gpar(fill = "white", col = "black", lwd = 5, lty = 1)),
      rectGrob(height = unit(0.9, "npc"), width = unit(0.9, "npc"), gp = gpar(fill = NA, col = "white", lwd = 1, lty = "longdash", lineend = "butt", linejoin = "round"))
    )
  }
  
  # Vectors for column widths
  cw_text <- 6.5 # width of text columns in cm, needs to align with wrap_length above
  cw_middle <- 0.75 # width of middle column in cm, adjusted manually
  
  # Vectors for row heights
  rh <- 0.5 # height of column in cm, halved for line entries in legend to save space
  
  # Vector for gap
  gap <- 0.01 # used below in the text drawing section, units in 'npc'
  
  # Text sizes
  ts_title <- 10 # For 'Legend' 
  ts_text <- 6 # for each line in the legend
  ts_boxnum <- 8 # for box number entry
  ts_boxnumExp <- 5 # for explanation under LC legend
  
  #
  ### Legend_BM gtable ####
  
  if (BM) {#only create imagery map if requested
    # Vectors for number of elements in legend table
    legend_items_BM <- filter(legend_items, BM == 1) # filtering out Land Cover
    n <- nrow(legend_items_BM)
    
    # Set column and row heights and widths
    
    #Set column widths and heights
    col_widths <- unit.c(unit(cw_text, "cm"), unit(cw_middle, "cm"), unit(cw_text, "cm")) # 3 columns
    row_heights <- unit(rep(rh, n),"cm") # regular row heights because less entries in this table
    
    # Specify vectors for table 
    Legend_BM <- gtable(widths = col_widths, heights = row_heights)
    
    # Add a title row at the top (bilingual headers)
    Legend_BM <- gtable_add_rows(Legend_BM, heights = unit(rh/2, "cm"), pos = 0) # empty row over titles to space from Top grobs
    Legend_BM <- gtable_add_rows(Legend_BM, heights = unit(rh*2, "cm"), pos = 1) # row for legend titles
    Legend_BM <- gtable_add_rows(Legend_BM, heights = unit(rh/2, "cm"), pos = 2) # empty row under titles
    row_offset <- 3 # Offset for looping through the actual data rows (titles + spacer occupy 3 rows in gtable indexing)
    
    # English title
    Legend_BM <- gtable_add_grob( 
      Legend_BM,
      grobs = textGrob(
        "Legend", 
        x = unit((1-gap), "npc"), just = "right",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")),
      t = 2, l = 1
    )
    
    # French title
    Legend_BM <- gtable_add_grob( 
      Legend_BM,
      grobs = textGrob(
        "L\u00E9gende", # Escaping french character
        x = unit(gap, "npc"), just = "left",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")), 
      t = 2, l = 3
    )
    
    # Loop through each row and produce the English Text, Symbol, and French Text for each row
    for (i in seq_len(n)) {
      # Vector r for r in the gtable intended, i will be row in legend_items_BM
      r <- i + row_offset
      
      # Left text (English)
      Legend_BM <- gtable_add_grob(
        Legend_BM,
        grobs = textGrob(
          legend_items_BM$eng[i],
          x = unit((1-gap), "npc"), just = "right",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 1
      )
      
      # Middle symbol grob 
      sym <- switch(
        legend_items_BM$type[i],
        "fill"   = grob_fill(legend_items_BM$fill[i], legend_items_BM$border[i]),
        "line"   = if (i == 1) grob_kba() else grob_line(legend_items_BM$lcol[i], legend_items_BM$lty[i], legend_items_BM$lwd[i]),
        "boxnum" = grob_boxnum(legend_items_BM$perc[i], fill = legend_items_BM$fill[i], col = legend_items_BM$border[i])
      )
      Legend_BM <- gtable_add_grob(Legend_BM, grobs = sym, t = r, l = 2)
      
      # Right text (French)
      Legend_BM <- gtable_add_grob(
        Legend_BM,
        grobs = textGrob(
          legend_items_BM$fr[i],
          x = unit(gap, "npc"), just = "left",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 3
      )
    }
    
    # Display
    grid.newpage(); grid.draw(Legend_BM)
  }
  #
  ### Legend_IM gtable ####
  
  if (IM) {#only create imagery map if requested
    # Vectors for number of elements in legend table
    legend_items_IM <- filter(legend_items, IM == 1) # filtering out Land Cover
    n <- nrow(legend_items_IM)
    
    # Set column and row heights and widths
    
    #Set column widths and heights
    col_widths <- unit.c(unit(cw_text, "cm"), unit(cw_middle, "cm"), unit(cw_text, "cm")) # 3 columns
    row_heights <- unit(rep(rh, n),"cm") # regular row heights because less entries in this table
    
    # Specify vectors for table 
    Legend_IM <- gtable(widths = col_widths, heights = row_heights)
    
    # Add a title row at the top (bilingual headers)
    Legend_IM <- gtable_add_rows(Legend_IM, heights = unit(rh/2, "cm"), pos = 0) # empty row over titles to space from Top grobs
    Legend_IM <- gtable_add_rows(Legend_IM, heights = unit(rh*2, "cm"), pos = 1) # row for legend titles
    Legend_IM <- gtable_add_rows(Legend_IM, heights = unit(rh/2, "cm"), pos = 2) # empty row under titles
    row_offset <- 3 # Offset for looping through the actual data rows (titles + spacer occupy 3 rows in gtable indexing)
    
    # English title
    Legend_IM <- gtable_add_grob( 
      Legend_IM,
      grobs = textGrob(
        "Legend", 
        x = unit((1-gap), "npc"), just = "right",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")),
      t = 2, l = 1
    )
    
    # French title
    Legend_IM <- gtable_add_grob( 
      Legend_IM,
      grobs = textGrob(
        "L\u00E9gende", # Escaping french character
        x = unit(gap, "npc"), just = "left",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")), 
      t = 2, l = 3
    )
    
    # Loop through each row and produce the English Text, Symbol, and French Text for each row
    for (i in seq_len(n)) {
      # Vector r for r in the gtable intended, i will be row in legend_items
      r <- i + row_offset
      
      # Left text (English)
      Legend_IM <- gtable_add_grob(
        Legend_IM,
        grobs = textGrob(
          legend_items$eng[i],
          x = unit((1-gap), "npc"), just = "right",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 1
      )
      
      # Middle symbol grob 
      sym <- switch(
        legend_items$type[i],
        "fill"   = grob_fill(legend_items$fill[i], legend_items$border[i]),
        "line"   = if (i == 1) grob_kba() else grob_line(legend_items$lcol[i], legend_items$lty[i], legend_items$lwd[i]),
        "boxnum" = grob_boxnum(legend_items$perc[i], fill = legend_items$fill[i], col = legend_items$border[i])
      )
      Legend_IM <- gtable_add_grob(Legend_IM, grobs = sym, t = r, l = 2)
      
      # Right text (French)
      Legend_IM <- gtable_add_grob(
        Legend_IM,
        grobs = textGrob(
          legend_items$fr[i],
          x = unit(gap, "npc"), just = "left",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 3
      )
    }
    
    # Display
    grid.newpage(); grid.draw(Legend_IM)
  }
  #
  ### Legend_LC gtable ####
  
  if (LC) { #only make a land cover map if requested
    # Vectors for number of elements in legend table
    legend_items_LC <- filter(legend_items, LC == 1) # filtering out Land Cover
    n <- nrow(legend_items_LC)
    
    #Set column widths and heights
    col_widths <- unit.c(unit(cw_text, "cm"), unit(cw_middle, "cm"), unit(cw_text, "cm")) # 3 columns
    row_heights <- unit.c(unit(ifelse(!is.na(legend_items_LC$nalcmsID)|legend_items_LC$eng == "Unclassified", rh, rh/2), "cm"))
    row_heights[1] <- unit(rh, "cm") # Setting the first entry, KBA, to regular height. 
    
    # Specify vectors for table 
    Legend_LC <- gtable(widths = col_widths, heights = row_heights)
    
    # Add a title row at the top (bilingual headers)
    Legend_LC <- gtable_add_rows(Legend_LC, heights = unit(rh/2, "cm"), pos = 0) # empty row over titles to space from Top grobs
    Legend_LC <- gtable_add_rows(Legend_LC, heights = unit(rh*2, "cm"), pos = 1) # row for legend titles
    Legend_LC <- gtable_add_rows(Legend_LC, heights = unit(rh/2, "cm"), pos = 2) # empty row under titles
    row_offset <- 3 # Offset for looping through the actual data rows (titles + spacer occupy 3 rows in gtable indexing)
    
    # English title
    Legend_LC <- gtable_add_grob( 
      Legend_LC,
      grobs = textGrob(
        "Legend", 
        x = unit((1-gap), "npc"), just = "right",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")),
      t = 2, l = 1
    )
    
    # French title
    Legend_LC <- gtable_add_grob( 
      Legend_LC,
      grobs = textGrob(
        "L\u00E9gende", # Escaping french character
        x = unit(gap, "npc"), just = "left",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")), 
      t = 2, l = 3
    )
    
    # Loop through each row and produce the English Text, Symbol, and French Text for each row
    for (i in seq_len(n)) {
      # Vector r for r in the gtable intended, i will be row in legend_items_LC
      r <- i + row_offset
      
      # Left text (English)
      Legend_LC <- gtable_add_grob(
        Legend_LC,
        grobs = textGrob(
          legend_items_LC$eng[i],
          x = unit((1-gap), "npc"), just = "right",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 1
      )
      
      # Middle symbol grob 
      sym <- switch(
        legend_items_LC$type[i],
        "fill"   = grob_fill(legend_items_LC$fill[i], legend_items_LC$border[i]),
        "line"   = if (i == 1) grob_kba() else grob_line(legend_items_LC$lcol[i], legend_items_LC$lty[i], legend_items_LC$lwd[i]),
        "boxnum" = grob_boxnum(legend_items_LC$perc[i], fill = legend_items_LC$fill[i], col = legend_items_LC$border[i])
      )
      Legend_LC <- gtable_add_grob(Legend_LC, grobs = sym, t = r, l = 2)
      
      # Right text (French)
      Legend_LC <- gtable_add_grob(
        Legend_LC,
        grobs = textGrob(
          legend_items_LC$fr[i],
          x = unit(gap, "npc"), just = "left",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 3
      )
    }
    
    # Create grob with descriptions at bottom of table (if KBA is published and numbers are in KBALandCover)
    if(published_site){
      g_legend     <- textGrob(
        "Numbers in the legend boxes indicate the approximate percentage of the KBA covered by each land cover type\n
        Les chiffres dans les rectangles de la l\u00E9gende indiquent le pourcentage approximatif de la KBA \nqui est concern\u00E9 par chacun des types de couverture terrestre",
        gp = gpar(family = "opensans", fontsize = ts_boxnumExp)
      )
      
      # Add to gtable
      Legend_LC <- gtable_add_rows(Legend_LC, heights = unit(1, "cm"), pos = r) # r should be equal to the max of the table
      Legend_LC <- gtable_add_grob(Legend_LC, g_legend,  t = (r+1), l = 1, b = (r+1), r = 3) # starting in new row, spreading across all three columns. 
    }
    
    # Display
    grid.newpage(); grid.draw(Legend_LC)
  }
  #
  ### Legend_BD gtable ####
  
  #Only make map if there are BiodivElementDist entries
  if (nrow(DBS_BiodivElementDistribution) > 0 & BD){
    
    # Vectors for number of elements in legend table
    legend_items_BD <- filter(legend_items, BD == 1) %>% # filtering only desired entries (just KBA so far)
      select(type, eng, fr, fill) #select columns to match with bd_data below
    
    # Add Biodiversity entries
    bd_data <- 
      left_join(
        select(DBS_BiodivElementDistribution, biodivelementdistributionid, fill), 
        select(DBS_SpeciesAtSite, biodivelementdistributionid, speciesatsiteid, speciesid),
        by = "biodivelementdistributionid"
      ) %>%
      left_join(
        .,
        select(DBS_BIOTICS_ELEMENT_NATIONAL, speciesid, national_scientific_name, national_engl_name, national_fr_name),
        by = "speciesid"
      ) %>%
      mutate(
        type = "fill",
        eng = paste(national_engl_name, paste0("(", national_scientific_name, ")"), sep = "\n"),
        fr = paste(national_fr_name, paste0("(", national_scientific_name, ")"), sep = "\n")
      ) %>%
      as.data.frame() %>%
      select(type, eng, fr, fill)
    
    # Bind tables
    legend_items_BD <- rbind(legend_items_BD, bd_data)
    
    # Wrap text in french and english columns (JIC long species name)
    wrap_length <- 50 # num char chosen based on Trial and Error, dependent on width of legend columns in Part 3
    wrap_text <- function(x) {x |> strwrap(width = wrap_length) |> paste(collapse = "\n")} 
    legend_items_BD$fr <- sapply(legend_items_BD$fr, wrap_text)
    legend_items_BD$eng <- sapply(legend_items_BD$eng, wrap_text)
    
    #Get number of rows
    n <- nrow(legend_items_BD)
    
    # Set column widths and heights
    col_widths <- unit.c(unit(cw_text, "cm"), unit(cw_middle, "cm"), unit(cw_text, "cm")) # 3 columns
    row_heights <- unit(rep(1.25*rh, n),"cm") # regular row heights because less entries in this table, need large enough for three rows of text
    
    # Specify vectors for table 
    Legend_BD <- gtable(widths = col_widths, heights = row_heights)
    
    # Add a title row at the top (bilingual headers)
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(rh/2, "cm"), pos = 0) # empty row over titles to space from Top grobs
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(rh*2, "cm"), pos = 1) # row for legend titles
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(rh/2, "cm"), pos = 2) # empty row under titles
    row_offset <- 3 # Offset for looping through the actual data rows (titles + spacer occupy 3 rows in gtable indexing)
    
    # English title
    Legend_BD <- gtable_add_grob( 
      Legend_BD,
      grobs = textGrob(
        "Legend", 
        x = unit((1-gap), "npc"), just = "right",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")),
      t = 2, l = 1
    )
    
    # French title
    Legend_BD <- gtable_add_grob( 
      Legend_BD,
      grobs = textGrob(
        "L\u00E9gende", # Escaping french character
        x = unit(gap, "npc"), just = "left",
        gp = gpar(fontsize = ts_title, family = "opensans", fontface = "bold")), 
      t = 2, l = 3
    )
    
    # Loop through each row and produce the English Text, Symbol, and French Text for each row
    for (i in seq_len(n)) {
      # Vector r for r in the gtable intended, i will be row in legend_items
      r <- i + row_offset
      
      # Left text (English)
      Legend_BD <- gtable_add_grob(
        Legend_BD,
        grobs = textGrob(
          legend_items_BD$eng[i],
          x = unit((1-gap), "npc"), just = "right",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 1
      )
      
      # Middle symbol grob 
      sym <- switch(
        legend_items_BD$type[i],
        "fill"   = grob_fill(legend_items_BD$fill[i], legend_items_BD$border[i], alpha = 0.5), 
        "line"   = grob_kba() #if (i == 1) grob_kba() else grob_line(legend_items_BD$lcol[i], legend_items_BD$lty[i], legend_items_BD$lwd[i]), #won't work without edits, not needed
        #"boxnum" = grob_boxnum(legend_items_BD$perc[i], fill = legend_items_BD$fill[i], col = legend_items_BD$border[i]) #won't work without edits, not needed
      )
      Legend_BD <- gtable_add_grob(Legend_BD, grobs = sym, t = r, l = 2)
      
      # Right text (French)
      Legend_BD <- gtable_add_grob(
        Legend_BD,
        grobs = textGrob(
          legend_items_BD$fr[i],
          x = unit(gap, "npc"), just = "left",
          gp = gpar(lineheight = 0.75, fontsize = ts_text, family = "opensans")
        ),
        t = r, l = 3
      )
    } # End Loop
    
    # Create grob with descriptions at bottom of table (if KBA is published and numbers are in KBALandCover)
    g_legend     <- textGrob(
      paste("Biodiversity element distributions represent the area where specific 
            biodiversity elements are found within a KBA.\n
            Les distributions d'\u00E9l\u00E9ments de biodiversit\u00E9 repr\u00E9sentent les zones o\u00F9 se situent 
            les diff\u00E9rents \u00E9l\u00E9ments de biodiversit\u00E9 au sein d'une KBA."),
      gp = gpar(family = "opensans", fontsize = ts_boxnumExp)
    )
    
    # Add to gtable
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(4, "cm"), pos = r) # r should be equal to the max of the table
    Legend_BD <- gtable_add_grob(Legend_BD, g_legend,  t = (r+1), l = 1, b = (r+1), r = 3) # starting in new row, spreading across all three columns. 
    
    # Add final gap with BD element dist labels
    g_legendEng <- textGrob("Biodiversity Element \nDistribution", x = unit((1-gap), "npc"), just = "right", gp = gpar(family = "opensans", fontsize = ts_text+2, fontface = "bold"))
    g_legendFr <- textGrob("Distributions des \u00E9l\u00E9ments de \nbiodiversit\u00E9", x = unit(gap, "npc"), just = "left", gp = gpar(family = "opensans", fontsize = ts_text+2, fontface = "bold"))
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(rh/2, "cm"), pos = 4) # empty row to create gap
    Legend_BD <- gtable_add_rows(Legend_BD, heights = unit(rh*1.5, "cm"), pos = 5) # large header for titles
    Legend_BD <- gtable_add_grob(Legend_BD, g_legendEng,  t = 6, l = 1) # starting in new row
    Legend_BD <- gtable_add_grob(Legend_BD, g_legendFr,  t = 6, l = 3) # starting in new row
    
    
    # Display
    grid.newpage(); grid.draw(Legend_BD)
    
  } # End IF
  #
  ## Text Blocks (bottom) ####
  # Text Wrapper
  wrap_length2 <- 67 # num char chosen based on Trial and Error, dependent on width of legend columns
  wrap_text2 <- function(x) {x |> strwrap(width = wrap_length2) |> paste(collapse = "\n")} 
  
  #Create grobs
  g_block1   <- textGrob(
    wrap_text2(paste0(
      "Map generated by KBA Canada (kba@kbacanada.org) on ", Sys.Date(),". Spatial Reference: EPGS:3857 (WGS 84/ Pseudo-Mercator). ", 
      "Data Sources: ",
      "KBA Canada (", ifelse(!published_site, "https://kbacanada.org/", paste0("https://kbacanada.org/site/?SiteCode=", DBS_KBASite$sitecode)), "); ", 
      "Government of Canada (2020 Land Cover of Canada, https://search.open.canada.ca/opendata/); ",
      "Open Street Maps (https://download.geofabrik.de/north-america/ canada-latest.osm.pbf); ",
      "ESRI Basemaps (https://www.esri.com/); ",
      "Maptiler Basemaps (https://cloud.maptiler.com/)."
      
    )
    ),
    x = 0.5, y = 0.5, gp = gpar(family = "opensans", fontsize = 5)
  )
  
  #Create grobs
  g_block2   <- textGrob(
    wrap_text2(paste0(
      "Carte produite par KBA Canada (kba@kbacanada.org) ", Sys.Date(),". R\u00E9f\u00E9rence spatiale: EPGS:3857 (WGS 84/ Pseudo-Mercator). ", 
      "Sources de donn\u00E9es: ",
      "KBA Canada (", ifelse(!published_site, "https://kbacanada.org/", paste0("https://kbacanada.org/site/?SiteCode=", DBS_KBASite$sitecode)), "); ", 
      "Gouvernement du Canada (2020 Couverture terrestre du Canada, https://search.open.canada.ca/opendata/); ",
      "Open Street Maps (https://download.geofabrik.de/north-america/ canada-latest.osm.pbf); ",
      "ESRI fond cartographique (https://www.esri.com/); ",
      "Maptiler fond cartographique (https://cloud.maptiler.com/)."
    )
    ),
    x = 0.5, y = 0.5, gp = gpar(family = "opensans", fontsize = 5)
  )
  
  # Create gtable
  TextBlocks <- gtable(
    widths  = unit.c(rep(unit((0.90*cw_text), "cm"),2)), # 1 column, uses width from legend section to match width
    #widths  = unit.c(unit((2*cw_text + cw_middle), "cm")), # 1 column, uses width from legend section to match width
    heights = unit.c(rep(unit(4, "cm"), 1)) # Creating each x cm row, 
  )
  
  #Add grobs to gtable
  TextBlocks <- gtable_add_grob(TextBlocks, g_block1, t = 1, l = 1)
  TextBlocks <- gtable_add_grob(TextBlocks, g_block2, t = 1, l = 2)
  
  # Use in plot_grid: wrap the gtable as a grob-drawing plot
  ggdraw() + draw_grob(TextBlocks)
  
  ## Edge (Right Edge with Names/Codes) ####
  # Text
  lab_top    <- DBS_KBASite$jurisdiction_en
  lab_middle <- ifelse(published_site, DBS_KBASite$nationalname, paste("DRAFT BOUNDARY -", DBS_KBASite$nationalname))
  lab_bottom <- ifelse(published_site, DBS_KBASite$sitecode, paste("KBASiteID", DBS_KBASite$kbasiteid))
  
  # Make a 3-row, 1-column gtable.
  # Row heights are 'null' so they share space evenly.
  Edge <- gtable(
    widths  = unit(0.75, "cm"), #strip width
    heights = unit(c(1.2, 4, 1.2), "null") # separations between terms
  )
  
  # Add a light background strip
  bg <- rectGrob(gp = gpar(fill = "#F4F0E8", col = NA))
  Edge <- gtable_add_grob(Edge, bg, t = 1, l = 1, b = 3, r = 1, z = -Inf)
  
  # Build the three rotated text grobs
  g_top <- textGrob(
    lab_top,
    rot = 90,
    x = 0.5, y = 0.5,
    just = "centre",
    #gp = gpar(fontface = "bold", family = "opensans")
    gp = gpar(fontface = "bold", fontfamily = "montserrat")
  )
  
  g_mid <- textGrob(
    lab_middle,
    rot = 90,
    x = 0.5, y = 0.5,
    just = "centre",
    #gp = gpar(fontface = "bold", family = "opensans")
    gp = gpar(fontface = "bold", fontfamily = "montserrat")
  )
  
  g_bot <- textGrob(
    lab_bottom,
    rot = 90,
    x = 0.5, y = 0.5,
    just = "centre",
    #gp = gpar(fontface = "bold", family = "opensans")
    gp = gpar(fontface = "bold", fontfamily = "montserrat")
  )
  
  # Place each grob into its row
  Edge <- gtable_add_grob(Edge, g_top, t = 1, l = 1)
  Edge <- gtable_add_grob(Edge, g_mid, t = 2, l = 1)
  Edge <- gtable_add_grob(Edge, g_bot, t = 3, l = 1)
  
  # Optional padding to give breathing room
  Edge <- gtable::gtable_add_padding(Edge, padding = unit(c(2, 6, 2, 6), "mm"))  # top, right, bottom, left
  
  # Draw
  grid.newpage();grid.draw(Edge)
  
  ## Left edge (Spaces KBASiteMap_LC from the edge) ####
  
  #Create gtable for empty rectangle
  EdgeL <- gtable(
    widths  = unit(0.75, "cm"), #strip width
    heights = unit(c(1), "null")
  )
  
  # Add an empty background strip
  bg <- rectGrob(gp = gpar(fill = NA, col = NA))
  EdgeL <- gtable_add_grob(EdgeL, bg, t = 1, l = 1, b = 1, r = 1, z = -Inf)
  
  # Draw
  grid.newpage();grid.draw(EdgeL)
  
  #
  # Format Final Map ####
  # Status update in Function
  cat("\nAll map components complete. Producing final map.\n")
  
  ## Create grobs for each vertical strip ####
  
  # Middle Top Section (With inset and KBA Logo side by side)
  MiddleTop <- plot_grid(KBAInset, KBACorner, ncol = 2, rel_widths = c(1,1))
  
  # Create vector for relative heights
  rel_heights <- c(1,2,0.75)
  
  # Middle Strip (With Middle Top, Legend_LC, Disclaimer Text, and Middle Bottom)
  if(BM) {Middle_BM <- plot_grid(MiddleTop, Legend_BM, TextBlocks, nrow = 3, rel_heights = rel_heights)}
  if(IM) {Middle_IM <- plot_grid(MiddleTop, Legend_IM, TextBlocks, nrow = 3, rel_heights = rel_heights)}
  if(LC) {Middle_LC <- plot_grid(MiddleTop, Legend_LC, TextBlocks, nrow = 3, rel_heights = rel_heights)}
  if(nrow(DBS_BiodivElementDistribution)>0 & BD){Middle_BD <- plot_grid(MiddleTop, Legend_BD, TextBlocks, nrow = 3, rel_heights = rel_heights)}
  
  #
  ## Map processing Function #### 
  # Vector for PDF file based on pathway entered into function
  pathway_filename <- paste0(pathway, ifelse(!published_site, paste0("ID", KBASiteID, ".pdf"), paste0(DBS_KBASite$sitecode, ".pdf")))
  
  # Create function to allow looping and altering of dpi based on file size  
  map_processing <- function(dpi) {
    
    # this is used instead of ggsave to maintain the right proportions based on page size
    # the vectorized pdf was too large for the website, so write pdf, load and convert to png, then rewrite as pdf).
    # Not a lot of processing time, especially compared to osm_extract. 
    
    ## Initial PDF Printing ##
    cairo_pdf(pathway_filename, height = 8.5, width = 14)
    
    rel_widths <- c(0.1, 1.9, 1, 0.1)
    
    # Page 1 - Basemap
    if (BM){
      print(
        plot_grid(
          EdgeL,
          KBASiteMap_BM,
          Middle_BM,
          Edge,
          ncol = 4,
          rel_widths = rel_widths
        )
      ) 
    }
    
    # Page 2 - Imagery
    if (IM){
      print(
        plot_grid(
          EdgeL,
          KBASiteMap_IM,
          Middle_IM,
          Edge,
          ncol = 4,
          rel_widths = rel_widths
        )
      ) 
    }
    # Page 3 - Land Cover
    if (LC) {
      print(
        plot_grid(
          EdgeL,
          KBASiteMap_LC,
          Middle_LC,
          Edge,
          ncol = 4,
          rel_widths = rel_widths
        )
      )
    }
    
    # Page 4 - BD Element Dists (if present)
    if(nrow(DBS_BiodivElementDistribution) > 0 & BD){
      print(
        plot_grid(
          EdgeL,
          KBASiteMap_BD,
          Middle_BD,
          Edge,
          ncol = 4,
          rel_widths = rel_widths
        )
      )
    }
    
    #End pdf
    dev.off()
    
    ## Load and Create pngs to reduce file size ##
    graphics.off()
    gc()
    Sys.sleep(0.5)
    
    # Start recreating pages as png files using the pdftools package
    pages <- pdftools::pdf_info(pathway_filename)$pages
    
    # While loop for creating png of each page 
    p <- 1
    while (p <= pages){
      # Determine filename for page
      pathway_page <- paste0(gsub(".pdf", "", pathway_filename), "_", p, ".png")
      
      # Convert PDF
      quiet(pdf_convert(pdf = pathway_filename, page = p, dpi = dpi, filenames = pathway_page))
      
      # Read png in R and assign to object
      assign(paste0("p", p), rasterGrob(readPNG(pathway_page), interpolate = TRUE))
      
      # Delete files
      file.remove(pathway_page)
      
      # Add counter at end of loop
      p <- p + 1
    }
    
    ## PDF Rewrite ##
    cairo_pdf(pathway_filename, height = 8.5, width = 14)
    
    # Loop through pages
    for(p in 1:pages){
      print(plot_grid(get(paste0("p", p))))
    }
    
    #End pdf
    dev.off()
    
    # Pause before calculation
    graphics.off()
    gc()
    Sys.sleep(0.5)
    
    # Calculate file size
    pathway_filesize <- file.size(pathway_filename)/1000000 # megabytes 
    
    # Set return to the filesize 
    return(pathway_filesize)
    
  }
  
  ## Apply Map Processing Function ####
  # Initial processing at standard dpi
  pathway_filesize <- map_processing(dpi)
  
  # Create modifiable DPI object
  dpi_mod <- dpi
  
  # Only limit filesize if set as parameter
  
  if (limit_filesize) {
    
    # Loop through map_processing function with different dpi values
    while(pathway_filesize > max_filesize) {
      cat("Reducing DPI to decrease filesize.\n")
      
      # If the reduced dpi is greater than 0...
      if(dpi_mod - dpi_reduction_factor >= dpi_min) { # arbitrary minimum, but text not really visible at 100
        
        #Decrease dpi by reduction factor
        dpi_mod <- dpi_mod - dpi_reduction_factor
        
        #Process Map at new dpi
        pathway_filesize <- map_processing(dpi_mod)
        
        # Return Status Check
        cat("DPI =", dpi_mod,"File size =", pathway_filesize, "MB\n")
        
      } else {
        # Print message
        warning("DPI reduction to decrease file size below ", max_filesize, " MB failed. File size is ", pathway_filesize, 
                " MB with a DPI of ", dpi_mod, 
                ". Try increasing the maximum file size or removing the filesize limit and running the function again.")
        
        # Exit while loop
        break
      }
    }
  }
  
  ## Final PDF ####
  
  # Open PDF (could turn off if annoying, but helpful for testing)
  if(open_pdf) {browseURL(pathway_filename)}
  
  # Calculate file size
  Sys.sleep(0.5)
  pathway_filesize <- file.size(pathway_filename)/1000000
  
  #Final status check
  message("\nPDF map successfully saved as:", pathway_filename, "\n",
          ifelse(open_pdf, "Map opened via browseURL function.\n", ""), 
          "File Size =", pathway_filesize, "mb.\n", 
          "DPI =", dpi_mod, "\n")
  
  ## Return table of info ####
  # Create list of column names
  tmp_names <- c("KBASiteID", "nationalname", "runtime", "filename", "filesize", "DPI")
  
  # Bind to table
  tmp_times <- data.frame(matrix(
    NA,
    nrow = 1,
    ncol = length(tmp_names),
    dimnames = list(NULL, tmp_names)
  ))
  
  # Extract Runtime
  runtime <- round(Sys.time()- start, 2)
  
  # Populate temporary table
  tmp_times$KBASiteID <- KBASiteID
  tmp_times$nationalname <- DBS_KBASite$nationalname
  tmp_times$runtime <- runtime
  tmp_times$filename <- pathway_filename
  tmp_times$filesize <- pathway_filesize # mb
  tmp_times$DPI <- dpi_mod
  
  # Return table
  return(tmp_times)
  
} # End function


#
#### Get taxonomic level (species or infraspecies) based on scientific name ####
nameLevel <- function(name){
  
  # Analyse text string
  nSpaces <- gregexpr(" ", name, fixed=T)[[1]] %>%
    as.integer() %>%
    {!. == -1} %>%
    sum()
  
  nCf <- gregexpr(" cf. ", name, fixed=T)[[1]] %>%
    as.integer() %>%
    {!. == -1} %>%
    sum()
  
  nSp <- gregexpr(" sp. ", name, fixed=T)[[1]] %>%
    as.integer() %>%
    {!. == -1} %>%
    sum()
  
  nX <- gregexpr(" x ", name, fixed=T)[[1]] %>%
    as.integer() %>%
    {!. == -1} %>%
    sum()
  
  nSpecial <- nCf + nSp + nX
  
  # Compute level
  level <- ifelse((nSpaces == 1) | (nSpaces == 1+nSpecial) | ((nX > 0) & (nSpaces == 2+nSpecial)),
                  "Species",
                  "Infraspecies")
  
  return(level)
}

#### Get single language from translated text string ####
extractLanguage <- function(string, language){
  
  if(grepl("ENGLISH", string, ignore.case = F)){
    en <- as.vector(str_locate(string, "ENGLISH")[1,1])
    
  }else{
    en <- NA
  }
  
  if(grepl("FRANCAIS|FRANÇAIS", string, ignore.case = F)){
    fr <- as.vector(str_locate(string, "FRANCAIS|FRANÇAIS")[1,1])
    
  }else{
    fr <- NA
  }
  
  if(language == "EN"){
    
    if(is.na(en)){
      
      if(is.na(fr)){
        result <- string
        
      }else{
        stop("No English translation provided")
      }
      
    }else{
      
      if(is.na(fr)){
        result <- sub("ENGLISH -|ENGLISH-", "", string) %>%
          trimws()
        
      }else{
        
        if(fr < en){
          result <- sub(".*ENGLISH", "", string) %>%
            sub("-", "", .) %>%
            trimws()
          
        }else{
          result <- sub("FRANCAIS.*|FRANÇAIS.*", "", string) %>%
            sub("ENGLISH -|ENGLISH-", "", .) %>%
            trimws()
        }
      }
    }
    
  }else{
    stop(paste("Method for language", language, "not yet implemented."))
  }
  
  return(result)
}
