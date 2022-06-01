##### Project Name and other inputs #########
server <- "https://central.server"
project <- "https://central.server/v1/projects/128/forms/22A10818.svc"  ## OData Link
username <- "username"
password <- "password"
timezone <- "Asia/Calcutta"
project_name <- "ProjectName"   ## Needed for output folder creation
workspace_folder <- "/path/to/work/folder/"  ## 
excel_qre_file <- "xls-qre.xlsx"

########### Create Project Folder ############
if (file.exists(project_name)){
  setwd(file.path(workspace_folder, project_name))
} else {
  dir.create(file.path(workspace_folder, project_name), showWarnings = FALSE)
  setwd(file.path(workspace_folder, project_name))
  
} 

########### Function to load/install libraries ############
loadpkg <- function(toLoad){
  for(lib in toLoad){
    if(! lib %in% installed.packages()[,1]) {
      install.packages(lib, repos='http://cran.rstudio.com/')
    }
    suppressMessages( library(lib, character.only=TRUE) )
  }
}

loadpkg(c("dplyr", "ReporteRs", "readxl", "pollster", "rio" , "knitr", 
          "openxlsx", "ruODK", "gridExtra", "stringr", "ggplot2", 
          "plotly", "htmlwidgets", "plyr", "tableHTML", "sjPlot", "writexl"))


########### Read Survey File & Find Single Select and Multi Select list ############
survey <- read_excel(paste0(workspace_folder,excel_qre_file), sheet = "survey")
choices <- read_excel(paste0(workspace_folder,excel_qre_file), sheet = "choices")

# single_select list
selectone <- dplyr::filter(survey, grepl('select_one', type))
single_select <- selectone$name

# multi_select list
selectmulti <- dplyr::filter(survey, grepl('select_multiple', type))
multi_select <- selectmulti$name

rm(selectone,selectmulti)



########### Set Project and Get Data & Schema #############
# Set project
project <- project

# `ruODK` users only need default settings to their ODK Central:
ru_setup(url = server, 
         un = username, 
         pw = password, 
         svc = project,
         tz = timezone)

# File attachment download location
loc <- fs::path("media")

Sys.sleep(5)  ## 5 Sec wait

# GET data table
fq_data <- ruODK::odata_submission_get(
  table = fq_svc$name[1], 
  local_dir = loc,
  download = FALSE,  ### TRUE for downloading attachments
  wkt=TRUE)

# Get schema
df_schema <- ruODK::form_schema_ext()




########### Function "odk_remove_group_names" with removed group-names  ############
odk_remove_group_names <- function(df_schema, fq_data){
  dfvarname <- c()
  for (x in 1: nrow(df_schema)) {
    splt <- strsplit(df_schema$path[x], split = "/")
    splt <- splt[[1]]
    element <- gsub("-","_", splt[length(splt)])
    dfvarname[x] <- element
  }
  
  df_schema_mod <- df_schema
  df_schema_mod$ruodk_name <- dfvarname
  
  df2 <- df_schema_mod[!(df_schema_mod$type=="structure" ),]
  varname <- df2$ruodk_name
  
  df3 = fq_data[,(2:(ncol(fq_data)-11) )]
  colnames(df3) <- varname
  
  # seperate server added data
  dmeta = fq_data[,(ncol(fq_data)-12+2):ncol(fq_data)]
  df <- cbind(df3,dmeta)
  
  ##### Remove all columns with d_ 
  cnames <- colnames(df3)
  dlist <- which(startsWith(cnames, "d_"))
  
  if (length(dlist) > 0){
    dlist_1 <- dlist[1]
    dlist_f <- length(dlist) + dlist_1 - 1
    d1 <- df3[-c(dlist_1:dlist_f)] ## all without "d_"
    d2 <- df3[c(dlist_1:dlist_f)]  ## all "d_" cols
    d2start <- df3[1:4]
    d2 <- cbind(d2start,d2,dmeta)
    d1 <- cbind(d1, dmeta)
    
    ## List all outputs
    listin <- list(data = df, data_r = d1, data_d = d2, scema = df_schema_mod)
  } else {
    listin <- list(data = df, scema = df_schema_mod)
  }
  
  listin  
}




########### Multiple Select function ###########

func_pulse_multisplit <- function(multiselectVariable, data, schema){
  
  # multiselectVariable <- multi_select[2]
  
  multisVector <- dplyr::pull(data, multiselectVariable)
  
  ## Title
  
  title_text <- schema$label_english[schema$ruodk_name == multiselectVariable]
  
  
  ## Options List
  options <- schema$choices_english[schema$ruodk_name == multiselectVariable][[1]]$labels
  values <- schema$choices_english[schema$ruodk_name == multiselectVariable][[1]]$values
  countOptions <- length(options)
  # options_serial <- c(1:countOptions)
  # names_df <- data.frame(serial = options_serial, Options = options)
  
  
  
  
  ## Row Build
  
  multiselectDF <- data.frame(matrix(NA, nrow = 0, ncol = countOptions))
  
  y <- 1
  while (y <= length(multisVector)) {
    cellVector <- unlist(strsplit(as.character(multisVector[y]), split=" "))
    
    ## Column Build
    rowVector <- NULL
    n <- 1
    while (n <= countOptions) {
      rowVector <- append(rowVector, (values[n]  %in% cellVector))
      n = n + 1
    }
    
    
    ## Vector to DF
    multiselectDF <- rbind(multiselectDF, rowVector)
    y = y + 1
  }
  ## name the columns 
  multiselectDFname <- multiselectDF
  colnames(multiselectDF) <- values
  # colnames(multiselectDF) <- c(1:countOptions)
  colnames(multiselectDFname) <- paste(options, sep = " ", collapse = NULL)
  
  figs <- list()
  tabls <- list()
  ## Create visuals
  i <- 1
  while (i <= countOptions) {
    truename <- paste0("true_c_",i)
    falsename <- paste0("false_c_",i)
    
    colvector <- multiselectDF[,i]
    truevalue <- length(which(colvector == "TRUE"))
    falsevalue <- length(which(colvector == "FALSE"))
    
    labels = c("Yes", "No")
    values = c(truevalue, falsevalue)
    tabl <- data.frame(labels, values)
    colnames(tabl) <- c("","Count")
    
    tabls[[i]] <- tabl
    i = i + 1
  }
  
  ## Return list
  list("raw" = multiselectDF, "raw_named" = multiselectDFname, "options" = options, "tables" = tabls, "title" = title_text)
}




########### Function strip HTML ########
cleanHTML <- function(htmlString) {
  t <- gsub("<.*?>", "", htmlString)
  t <- gsub("[\r\n]", " ", t)
  t <- gsub("'", "", t)
  t <- gsub("\\**", "", t)
  t <- str_squish(t)
  t <- str_to_title(t, locale = "en")
  t <- str_trunc(t, 200, side = c("center"), ellipsis = "...")
  return(t)
}

########### Start processing ##############

# remove group names data + schema
t <- odk_remove_group_names(df_schema, fq_data)

data <- t$data
schema <- t$scema
rm(t)


# Spilt select multiple questions
data_o <- data
x <- 1
while (x <= length(multi_select)) {
  # z dataframe with multiselect split
  y <- func_pulse_multisplit(multi_select[x], data_o, schema)
  z <- y$raw
  z[z=="TRUE"]<- 1
  z[z=="FALSE"]<- 0
  st <- paste0(multi_select[x],"/",names(z), sep="")   ##
  colnames(z) <- st
  indexvar <- which(colnames(data_o) == multi_select[x])
  lefttable <- data_o[1:indexvar]
  righttable <- data_o[(indexvar+1):ncol(data_o)]
  data_o <- cbind(lefttable,z,righttable)
  x = x + 1
}
rm(x,y,z,lefttable,righttable,indexvar,st)


########### Prepare CSV output with labels as values +  Add Variable Names to col heads #######

# Data frame for labeled values and variables dala_OL
data_OL <- data_o
varlabels <- schema$label_english
varnames <- schema$ruodk_name

## replace NA by 0 in varlabels vector
varlabels[is.na(varlabels)] <- 0

## Name variables as per lables in the data_OL df
x <- 1
while (x <= length(varlabels)) {
  if ((varlabels[x]) != 0){names(data_OL)[names(data_OL) == varnames[x]] <- trimws(cleanHTML(varlabels[x]))}
  x = x + 1
}

# Multiselect - add option text in column head

x <- 1
while (x <= length(multi_select)) {
  s <- schema %>% filter(ruodk_name == multi_select[x])
  sx <- s$choices_english
  sxv <- sx[[1]]$values
  sxl <- sx[[1]]$labels
  
  y <- 1
  while (y <= length(sxv)) {
    num <- which(colnames(data_OL) == paste0(multi_select[x],"/",sxv[y])) ##
    if (sxl[y] == "") {
      colnames(data_OL)[num] <- paste0(multi_select[x],"/",sxv[y]) ##
    } else {
      colnames(data_OL)[num] <- paste0(multi_select[x],"/",sxl[y]) ##
    }
    
    # print(paste("X Loop: ",x, " | Y loop: ", y))
    y = y + 1
    
  }
  x = x + 1
}


# Single select option text on data_OL

x <- 1
while (x <= length(single_select)) {
  s <- schema %>% filter(ruodk_name == single_select[x])
  sx <- s$choices_english
  
  sxv <- sx[[1]]$values
  sxl <- sx[[1]]$labels
  
  sx <- as.data.frame(cbind(sxv,sxl))
  colnames(sx) <- c("val","label") 
  sx <- sx[!duplicated(sx), ]
  
  index <- which(colnames(data_o) == single_select[x])
  colvector <- data_OL[index]
  coln <- colnames(colvector)
  colnames(colvector) <- c("val")
  dft <- left_join(colvector, sx, by="val")
  
  outvec <- dft$label
  data_OL[, index] <- outvec
  x = x + 1
}

########### SPSS Script Build ################

varnames <- schema$ruodk_name
varlables_spss <- cleanHTML(schema$label_english) 
varlables_spss[is.na(varlables_spss)] <- ""
df <- data.frame(varnames, varlables_spss)

df <- df %>% 
  mutate(spsslables = ifelse(varlables_spss == "", df$varnames, df$varlables_spss))

spsslables <- df$spsslables

dfspsslables <- data.frame(varname = varnames, lables = spsslables)
rm(varnames, varlables_spss, df)

outputvars <- colnames(data_o)
outputvars <- data.frame(varname = outputvars)

xtemp <- left_join(outputvars, dfspsslables, by = "varname")
xtemp <- cbind(spss = "VARIABLE LABELS ", xtemp)
colnames(xtemp) <- c("spss","varname","lables")


# Replacing multi-select option variable names

veclabels <- xtemp$lables

vecmulti <- colnames(data_o)
v <- vecmulti  %>% str_subset(pattern = "/") ##

x <- 1
while (x <= length(v)) {
  strs <- str_split(v[x], "/") ##
  strs <- strs[[1]]
  indexinvecmulti <-  which(vecmulti == paste0(v[x]))
  
  s <- schema %>% filter(ruodk_name == strs[1])
  sx <- s$choices_english
  
  sxv <- sx[[1]]$values
  sxl <- sx[[1]]$labels
  
  indexsxv <-  which(sxv == strs[2])
  
  veclabels[indexinvecmulti] <- paste0(gsub("/",".",v[x]), " ", sxl[as.integer(indexsxv)])
  
  
  x = x + 1
}

xtemp <- cbind(xtemp[1:2], veclabels)

# Write SPSS Syntax File VARIABLE LABELS

text <- paste0(xtemp$spss, gsub("/",".",xtemp$varname), " '", cleanHTML(xtemp$veclabels),  "'.")
cat(text, file = paste0(workspace_folder, project_name, "/spss_syntax.sps"), sep = "\n")



# Append SPSS Syntax File VALUE LABELS

t <- ""

x <- 1

while (x <= length(single_select)) {
  s <- schema %>% filter(ruodk_name == single_select[x])
  sx <- s$choices_english
  sxv <- sx[[1]]$values
  sxl <- sx[[1]]$labels
  sxl <- cleanHTML(sxl)
  
  
  t1 <- paste0("VALUE LABELS ",single_select[x])
  t2 <- paste0("\t", sxv, " '", sxl, "'", collapse = "\n" )
  text2 <- paste0(t1, "\n", t2, sep = "\n" )
  
  cat("\n", file = paste0(workspace_folder, project_name, "/spss_syntax.sps"), sep = "\n", append=TRUE)
  cat(text2, file = paste0(workspace_folder, project_name, "/spss_syntax.sps"), append=TRUE)
  cat(". \n",  file = paste0(workspace_folder, project_name, "/spss_syntax.sps"), append=TRUE)
  
  x = x + 1
}

## SPSS Data-Frame | Align variable names with syntax (replace / with . for var names in)

data_spss <- data_o
names <- colnames(data_spss)
names <- gsub("/",".",names)
colnames(data_spss) <- names

## Write data files
write_xlsx(data_o, path = paste0(workspace_folder, project_name, "/data_out.xlsx"))  ## Standard ODK output format
write_xlsx(data_OL, path = paste0(workspace_folder, project_name, "/data_out_L.xlsx")) ## File with variable and value labels 
write_xlsx(data_spss, path = paste0(workspace_folder, project_name, "/spss_data.xlsx")) ## File for SPSS importing


