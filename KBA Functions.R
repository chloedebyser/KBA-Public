#### Identification of Canadian KBAs
#### Wildlife Conservation Society - 2021-2023
#### Script by Chloé Debyser

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
    {ifelse(grepl("g", ., T), ifelse(grepl("n", ., T), "Global and National", "Global"), "National")}
  
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
      mutate(Field = replace(Field, Field == "Name", "Name of proposal development lead"),
             Field = replace(Field, Field == "Email", "Email of proposal development lead"),
             Field = replace(Field, Field == "Organization", "Organization of proposal development lead"),
             Field = replace(Field, Field == "I agree to the data in this form being stored in the World Database of KBAs and used for the purposes of KBA identification and conservation.", "I agree to the data in this form being stored in the Canadian KBA Registry and in the World Database of KBAs, and used for the purposes of KBA identification and conservation."))
    
    proposer %<>%
      add_row(Field = "Name(s) to be displayed publicly", .after = proposer %>% with(which(Field == "Organization of proposal development lead")))
  }
  
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
    
    siteVersion <- site %>%
      filter(Field == "Purpose") %>%
      pull(GENERAL) %>%
      trimws() %>%
      {ifelse(. == "1. Propose a new KBA that does not intersect any existing KBAs", 1, NA)} %>%
      as.integer()
    
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
    mutate(`Common name` = trimws(`Common name`),
           `Scientific name` = trimws(`Scientific name`),
           `Derivation of best estimate` = ifelse(`Derivation of best estimate` == "Other (please add further details in column AA)", "Other", `Derivation of best estimate`)) %>%
    arrange(`Scientific name`)
  
        # Handle differing form versions
  if(formVersion < 1.2){
    
    species %<>%
      rename(`RU source` = `RU Source`,
             `Seasonal distribution` = `Seasonal Distribution`) %>%
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
  threats %<>%
    left_join(., DB_BIOTICS_ELEMENT_NATIONAL[which(!is.na(DB_BIOTICS_ELEMENT_NATIONAL$national_engl_name)), c("speciesid", "national_engl_name")], by=c("Specific biodiversity element" = "national_engl_name")) %>%
    relocate(., speciesid, .after = 'Specific biodiversity element')
  
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
           `Short citation` = trimws(`Short citation`))
  
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
}

#### KBA-EBAR Database - Load data ####
read_KBAEBARDatabase <- function(datasetNames, type, environmentPath, account){

  # Load password and CRS
  if(!missing(environmentPath)){
    load(environmentPath)
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
                     c("KBAInputPolygon", "KBA_View/FeatureServer/6", T),
                     c("KBAAcceptedSite", "KBA_Accepted_Sites/FeatureServer/0", T),
                     c("DatasetSource", "Restricted/FeatureServer/5", F),
                     c("InputDataset", "Restricted/FeatureServer/7", F),
                     c("ECCCRangeMap", "Restricted/FeatureServer/2", T),
                     c("RangeMap", "Restricted/FeatureServer/10", F),
                     c("EcoshapeOverviewRangeMap", "EcoshapeRangeMap/FeatureServer/1", T))
  
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
      
      query <- DB_DatasetSource %>%
        filter(datasetsourcename == "ECCC Range Maps") %>%
        pull(datasetsourceid) %>%
        {DB_InputDataset[which(DB_InputDataset$datasetsourceid %in% .), "inputdatasetid"]} %>%
        {paste0("inputdatasetid IN (", paste(., collapse=","), ")")}
      
    }else{
      
      query <- "OBJECTID >= 0"
    }
    
    # Get GeoJSON
    url <- parse_url("https://gis.natureserve.ca/arcgis/rest/services")
    url$path <- paste(url$path, paste0("EBAR-KBA/", address, "/query"), sep = "/")
    url$query <- list(where = query,
                      outFields = "*",
                      returnGeometry = "true",
                      f = "geojson")
    request <- build_url(url)
    response <- VERB(verb = "GET",
                     url = request,
                     add_headers(`Authorization` = paste("Bearer ", token)))
    data <- content(response, as="text") %>%
      geojson_sf()
    
    # If non-spatial, drop geometry
    if(!spatial){
      data %<>% st_drop_geometry()
    
    # If spatial, transform to the database CRS
    }else{
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
    
    if("n_rank_review_date" %in% colnames(data)){
      data %<>% mutate(n_rank_review_date = as.POSIXct(as.numeric(n_rank_review_date)/1000, origin = "1970-01-01", tz = "GMT"))
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
                     c("KBAAcceptedSite", "KBA_Accepted_Sites/FeatureServer/0", T))
  
  # Only retain datasets that are desired
  if(!missing(datasetNames)){
    DBdatasets <- sapply(DBdatasets, function(x) x[1] %in% datasetNames) %>%
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
      
      if(name %in% c("KBASite", "SpeciesAtSite", "EcosystemAtSite", "KBACitation", "KBAThreat", "KBAAction", "KBALandCover", "KBAProtectedArea", "OriginalDelineation", "KBAAcceptedSite")){
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
    
    newDataset %<>%
      left_join(., oldDataset, by=cols[which(!cols == id)]) %>%
      mutate({{id}} := get(paste0(id, ".y"))) %>%
      select(all_of(cols))
  }
  
  # Detect deletions
  if(sum(!oldDataset[[id]] %in% newDataset[[id]]) > 0){
    deletions %<>% add_row(Dataset = dataset,
                           ID = oldDataset[which(!oldDataset[[id]] %in% newDataset[[id]]), id])
  }
  
  # Add primary key, where missing
        # Get maximum ID already assigned
  maxID <- ifelse(exists(paste0(dataset, "_final")),
                  max(get(paste0(dataset, "_final"))[[id]]),
                  ifelse(nrow(DBDataset) > 0,
                         max(DBDataset[[id]]),
                         0))
  
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
# TO DO: Add check that meetscriteria is never NA
# TO DO: Add checks on conservation status being used (i.e. that there needs to be a status at all - i.e. only A1 - and that it is the correct one)
check_KBADataValidity <- function(){
  
  # Starting parameters
  error <- F
  message <- c()
  
  # KBA CANADA PROPOSAL FORM
        # Check that proposer email is provided twice
  if(PF_formVersion < 1.2){
    if(!PF_proposer$Entry[which(PF_proposer$Field == "Email of proposal development lead")] == PF_proposer$Entry[which(PF_proposer$Field == "Email (please re-enter)")]){
      error <- T
      message <- c(message, "The proposer email is not correctly entered in one of the required fields.")
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
  
  # KBA-EBAR DATABASE
        # Boundary generalization
  if(DBS_KBASite$boundarygeneralization == "2"){
    error <- T
    message <- c(message, "The site boundary needs to be generalized!")
  }
  
        # PresentAtSite
  presentAtSite <- DBS_SpeciesAtSite %>%
    filter(meetscriteria == "Yes") %>%
    pull(presentatsite) %>%
    unique()
  
  if(sum(!presentAtSite == "Y") > 0){
    error <- T
    message <- c(message, "Some SpeciesAtSite records that meet criteria are not flagged as PresentAtSite = Yes")
  }
  
  # CONCURRENCE BETWEEN KBA CANADA PROPOSAL FORM AND KBA-EBAR DATABASE
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
  
  Jurisdiction_PF <- PF_site %>%
    filter(Field == "Province or Territory") %>%
    pull(GENERAL) %>%
    trimws()
  
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
  
  if((!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=2)) & ((str_sub(formatC(round(SiteArea_DB, 2), format='f', digits=2), start=-1) == "0") & (!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=1))) & ((str_sub(formatC(round(SiteArea_DB, 2), format='f', digits=2), start=-2) == "00") & (!SiteArea_PF == formatC(round(SiteArea_DB, 2), format='f', digits=0)))){
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
  
  if(!sum(LatLon_DB == LatLon_PF) == 2){
    error <- T
    message <- c(message, "There is a mismatch between the latitude and longitude provided in the proposal form and that provided in the database.")
  }
  
        # Species
              # Based on NATIONAL_SCIENTIFIC_NAME
                    # All species
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
  
                    # Species meeting KBA criteria
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
  SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
    filter(national_engl_name %in% PF_species$`Common name`) %>%
    pull(speciesid) %>%
    unique()
  
  if(length(SpeciesIDs_PF) > nrow(PF_species)){
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
  
                    # Species meeting KBA criteria
  SpeciesIDs_PF <- DB_BIOTICS_ELEMENT_NATIONAL %>%
    filter(national_engl_name %in% PF_species[which(!is.na(PF_species$`Criteria met`)), "Common name"]) %>%
    pull(speciesid) %>%
    unique()
  
  if(length(SpeciesIDs_PF) > nrow(PF_species)){
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
  
  # Return end parameters
  return(list(error, message))
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
  
        # Add informal taxonomic group
  speciesatsite %<>%
    left_join(., speciesBiotics[, c("speciesid", "kba_group")], by="speciesid") %>%
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
  
        # Add informal classification group
  ecosystematsite %<>%
    left_join(., ecosystemBiotics[, c("ecosystemid", "subclass_name")], by="ecosystemid") %>%
    rename(kba_group = subclass_name) %>%
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
  
  # Get prepositions, for FR only
  if(language == "FR"){
    biodivelements %<>% mutate(preposition = sapply(kba_group_translated, function(x) ifelse(substr(x, start=1, stop=1) %in% c("A", "E", "I", "O", "U", "Y"), "d'", "de ")))
    
  }else{
    biodivelements %<>% mutate(preposition = NA)
  }
  
  # Sort groups by alphabetical order, for consistency
  biodivelements %<>% arrange(level, criteriamet, kba_group_translated)
  
  # Initialize final text
  finalText <- c()
  finalText_level <- c()
  
  # Create text
  for(global_national in unique(biodivelements$level)){
    
    for(criterion in biodivelements %>% filter(level == global_national) %>% pull(criteriamet) %>% unique()){
      
      # Get the biodiversity elements that meet the criterion at that level
      biodivelements_criterion <- biodivelements %>%
        filter((level == global_national) & (criteriamet == criterion))
      
      # Get the gender of the concatenated group
      if(language %in% c("FR", "ES")){
        
        gender <- KBAgroups %>%
          filter(KBA_Group_EN %in% biodivelements_criterion$kba_group) %>%
          pull(paste0("Gender_", language))
        
        gender <- ifelse("M" %in% gender, "M", "F") # If there is a masculine noun then the group is masculine
      }
      
      # Get header
      headerLabel <- paste0("Header_", language, ifelse(language %in% c("FR", "ES"), paste0("_", gender), ""))
      header <- criterionHeader %>%
        filter(Criterion == criterion) %>%
        pull(headerLabel)
      
      # Concatenate taxonomic groups
      if(grepl("#", header, fixed=T)){
        
        groups <- biodivelements_criterion %>%
          mutate(withpreposition = paste0(preposition, kba_group_translated)) %>%
          pull(withpreposition) %>%
          pasteEnumeration(string=.)
        
        header <- gsub("#", "", header, fixed=T)
        
      }else{
        groups <- pasteEnumeration(string=biodivelements_criterion$kba_group_translated)
      }
      
      # Insert groups into header
      header <- gsub("*", groups, header, fixed=T)
      
      # Final text
      finalText <- c(finalText, paste0(str_to_sentence(header), " (", criterion, ")"))
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
    
    finalText <- paste(paste(string[1:(length(string)-1)], collapse=", "), string[length(string)], collapse=" & ")
  }
  
  # Return final text
  return(finalText)
}

#### Miscellaneous - Convert m2 to km2 ####
m2tokm2 <- function(x){
  y <- x/1000000
  return(y)
}