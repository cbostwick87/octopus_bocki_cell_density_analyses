# five_replicates
setwd("C:/Users/cbost/Desktop/BISC/Gabby_paper/7-5-24/")

######################
######################
### Load Libraries ###
######################
######################

library(tidyverse)
library(readxl)
library(openxlsx)
library(lme4)
library(broomExtra)
library(glmmTMB)

#######################################
###########   5-11-24   ##############
#####################################
# Import data from excel file
# Path to the Excel file
file_path <- "TotalonlyNormalizedafteraggregating_SQUAREMM.xlsx"
#
#############################
#############################
# base vs tip for each probe each is replicate

### future cell size comp
# tip vs base XL
# tip L vs base L
# tip M vs base M
# tip S vs base S
# Read the names of the sheets
# Load necessary libraries

# Load the workbook and get the sheet names
wb <- loadWorkbook(file_path)
sheet_names <- names(wb)

# Function to process each sheet and extract data
process_sheet <- function(sheet_name) {
  # Read the data from the sheet
  data <- read.xlsx(file_path, sheet = sheet_name)
  
  # Determine the type based on the sheet name (Tip or Base)
  type <- ifelse(grepl("Tip", sheet_name), "Tip", "Base")
  
  # Determine the probe based on the sheet name
  probe <- strsplit(sheet_name, " ")[[1]][1]
  
  # Convert data to long format and add Type, Replicate, and Probe columns
  data_long <- data %>%
    pivot_longer(cols = starts_with(c("Tip_", "Base_")), 
                 names_to = "Replicate", values_to = "Count") %>%
    mutate(Type = type, Probe = probe) %>% 
    select(Probe, Type, Replicate, Count)
  
  return(data_long)
}

# Process each sheet and combine the data
data_list <- lapply(sheet_names, process_sheet)
combined_data <- bind_rows(data_list)

# View the combined data
print(combined_data)

# Function to fit the model and tidy results for each probe
fit_and_tidy_model <- function(probe_data) {
  model <- glmmTMB(Count ~ Type + (1 | Replicate), data = probe_data, family = nbinom2)
  tidy(model)
}

# Split data by probe, fit the model, and tidy the results for each probe
probe_results <- combined_data %>%
  group_by(Probe) %>%
  group_map(~ fit_and_tidy_model(.x))

# Combine the results into a named list
names(probe_results) <- unique(combined_data$Probe)

# Create a new workbook
wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
for(probe_name in names(probe_results)) {
  addWorksheet(wb, probe_name)
  writeData(wb, probe_name, probe_results[[probe_name]])
}

# Save the workbook
output_file_path <- "C:/Users/cbost/Desktop/BISC/Gabby_paper/7-5-24/GLMM_neg_binom_total_cell_area_aggregated_first_7-5-24.xlsx"
saveWorkbook(wb, output_file_path, overwrite = TRUE)
###################################################################################
################   Ipsilateral vs contralateral model  ##########################
################################################################################

#############################
# excel_file <- "IConlyNormalizedafteraggregating_SQUAREMM.xlsx"
# ipsi tip vs contra tip row 2 vs row 3 # GLMM_neg_binom_IC_internal_area_aggregated_first_7-5-24
# ipsi tip vs ipsi base # GLMM_neg_binom_IC_external_ipsi_area_aggregated_first_7-5-24
# contra tip vs contra base # GLMM_neg_binom_IC_external_contra_area_aggregated_first_7-5-24

# Define the path to the new Excel workbook
file_path <- "IConlyNormalizedafteraggregating_SQUAREMM.xlsx"

# Load the workbook and get the sheet names
wb <- loadWorkbook(file_path)
sheet_names <- names(wb)

# Function to process each sheet and extract data
process_sheet <- function(sheet_name) {
  # Read the data from the sheet
  data <- read.xlsx(file_path, sheet = sheet_name, colNames = TRUE)
  
  # Determine the type based on the sheet name (Tip or Base)
  type <- ifelse(grepl("Tip", sheet_name), "Tip", "Base")
  
  # Determine the probe based on the sheet name
  probe <- strsplit(sheet_name, " ")[[1]][1]
  
  # Convert data to long format and add Type, Replicate, Probe, and Side columns
  data_long <- data %>%
    pivot_longer(cols = starts_with(c("Tip_", "Base_")), 
                 names_to = "Replicate", values_to = "Count") %>%
    mutate(Type = type, Probe = probe, Side = rep(c("I", "C"), each = n() / 2)) %>% 
    select(Probe, Type, Replicate, Side, Count)
  
  return(data_long)
}

# Process each sheet and combine the data
data_list <- lapply(sheet_names, process_sheet)
combined_data <- bind_rows(data_list)

# View the combined data
print(combined_data)

# Function to fit the model and tidy results for each comparison
fit_and_tidy_model <- function(data, comparison) {
  if (comparison == "I_vs_C" && length(unique(data$Side)) > 1) {
    model <- glmmTMB(Count ~ Side + (1 | Replicate), 
                     data = data, 
                     family = nbinom2, 
                     control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
  } else if (comparison == "I_Tip_vs_Base" && length(unique(data$Type)) > 1) {
    data <- data %>% filter(Side == "I")
    if (length(unique(data$Type)) > 1) {
      model <- glmmTMB(Count ~ Type + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "C_Tip_vs_Base" && length(unique(data$Type)) > 1) {
    data <- data %>% filter(Side == "C")
    if (length(unique(data$Type)) > 1) {
      model <- glmmTMB(Count ~ Type + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else {
    return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
  }
  tidy(model)
}

# Prepare a list to store all model results
model_results_list <- list()

# Fit and store models for each probe
for (probe in unique(combined_data$Probe)) {
  for (type in unique(combined_data$Type)) {
    # Data for the current probe and type
    current_data <- combined_data %>%
      filter(Probe == probe & Type == type)
    
    # Fit models for Ipsilateral vs Contralateral
    model_results_list[[paste(probe, type, "I_vs_C", sep = "_")]] <- fit_and_tidy_model(current_data, "I_vs_C")
  }
  
  # Fit models for Ipsilateral Tip vs Base
  current_data_i <- combined_data %>%
    filter(Probe == probe & Side == "I")
  model_results_list[[paste(probe, "I_Tip_vs_Base", sep = "_")]] <- fit_and_tidy_model(current_data_i, "I_Tip_vs_Base")
  
  # Fit models for Contralateral Tip vs Base
  current_data_c <- combined_data %>%
    filter(Probe == probe & Side == "C")
  model_results_list[[paste(probe, "C_Tip_vs_Base", sep = "_")]] <- fit_and_tidy_model(current_data_c, "C_Tip_vs_Base")
}

# Create a new workbook
wb <- createWorkbook()

# Loop through model results and write each to a new sheet
for (model_name in names(model_results_list)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, model_results_list[[model_name]])
}

# Save the workbook
output_file_path <- "C:/Users/cbost/Desktop/BISC/Gabby_paper/7-5-24/GLMM_neg_binom_IC_internal_area_aggregated_first_7-5-24.xlsx"
saveWorkbook(wb, output_file_path, overwrite = TRUE)


###################################################################################
##########   Top vs middle vs bottom model (oral/aboral)   ######################
################################################################################

# top vs middle, top vs bottom, middle vs bottom # GLMM_neg_binom_TMB_internal_area_aggregated_first_7-5-24
# tip top vs base top # GLMM_neg_binom_TMB_external_Top_area_aggregated_first_7-5-24
# tip m vs base m # GLMM_neg_binom_TMB_external_Middle_area_aggregated_first_7-5-24
# tip b vs base b # GLMM_neg_binom_TMB_external_Bottom_area_aggregated_first_7-5-24

file_path <- "TMBonlyNormalizedafteraggregating_SQUAREMM.xlsx"

# Load the workbook and get the sheet names
wb <- loadWorkbook(file_path)
sheet_names <- names(wb)

# Function to process each sheet and extract data
process_sheet <- function(sheet_name) {
  # Read the data from the sheet
  data <- read.xlsx(file_path, sheet = sheet_name, colNames = TRUE)
  
  # Determine the type based on the sheet name (Tip or Base)
  type <- ifelse(grepl("Tip", sheet_name), "Tip", "Base")
  
  # Determine the probe based on the sheet name
  probe <- strsplit(sheet_name, " ")[[1]][1]
  
  # Convert data to long format and add Type, Replicate, Probe, and Level columns
  data_long <- data %>%
    pivot_longer(cols = starts_with(c("Tip_", "Base_")), 
                 names_to = "Replicate", values_to = "Count") %>%
    mutate(Type = type, Probe = probe, Level = rep(c("T", "M", "B"), each = n() / 3)) %>% 
    select(Probe, Type, Replicate, Level, Count)
  
  return(data_long)
}

# Process each sheet and combine the data
data_list <- lapply(sheet_names, process_sheet)
combined_data <- bind_rows(data_list)

# View the combined data
print(combined_data)

# Function to fit the model and tidy results for each comparison
fit_and_tidy_model <- function(data, comparison) {
  if (comparison == "T_vs_M" && length(unique(data$Level)) > 1) {
    data <- data %>% filter(Level %in% c("T", "M"))
    if (length(unique(data$Level)) > 1) {
      model <- glmmTMB(Count ~ Level + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "T_vs_B" && length(unique(data$Level)) > 1) {
    data <- data %>% filter(Level %in% c("T", "B"))
    if (length(unique(data$Level)) > 1) {
      model <- glmmTMB(Count ~ Level + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "M_vs_B" && length(unique(data$Level)) > 1) {
    data <- data %>% filter(Level %in% c("M", "B"))
    if (length(unique(data$Level)) > 1) {
      model <- glmmTMB(Count ~ Level + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "Tip_vs_Base_T" && length(unique(data$Type)) > 1) {
    data <- data %>% filter(Level == "T")
    if (length(unique(data$Type)) > 1) {
      model <- glmmTMB(Count ~ Type + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "Tip_vs_Base_M" && length(unique(data$Type)) > 1) {
    data <- data %>% filter(Level == "M")
    if (length(unique(data$Type)) > 1) {
      model <- glmmTMB(Count ~ Type + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else if (comparison == "Tip_vs_Base_B" && length(unique(data$Type)) > 1) {
    data <- data %>% filter(Level == "B")
    if (length(unique(data$Type)) > 1) {
      model <- glmmTMB(Count ~ Type + (1 | Replicate), 
                       data = data, 
                       family = nbinom2, 
                       control = glmmTMBControl(optCtrl = list(iter.max = 10000, eval.max = 10000)))
    } else {
      return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
    }
  } else {
    return(tibble(term = NA, estimate = NA, std.error = NA, statistic = NA, p.value = NA))
  }
  tidy(model)
}

# Prepare a list to store all model results
model_results_list <- list()

# Fit and store models for each probe and type
for (probe in unique(combined_data$Probe)) {
  for (type in unique(combined_data$Type)) {
    # Data for the current probe and type
    current_data <- combined_data %>%
      filter(Probe == probe & Type == type)
    
    # Fit models for Top vs Middle
    model_results_list[[paste(probe, type, "T_vs_M", sep = "_")]] <- fit_and_tidy_model(current_data, "T_vs_M")
    
    # Fit models for Top vs Bottom
    model_results_list[[paste(probe, type, "T_vs_B", sep = "_")]] <- fit_and_tidy_model(current_data, "T_vs_B")
    
    # Fit models for Middle vs Bottom
    model_results_list[[paste(probe, type, "M_vs_B", sep = "_")]] <- fit_and_tidy_model(current_data, "M_vs_B")
  }
  
  # Fit models for Tip Top vs Base Top
  current_data_top <- combined_data %>%
    filter(Probe == probe & Level == "T")
  model_results_list[[paste(probe, "Tip_vs_Base_T", sep = "_")]] <- fit_and_tidy_model(current_data_top, "Tip_vs_Base_T")
  
  # Fit models for Tip Middle vs Base Middle
  current_data_middle <- combined_data %>%
    filter(Probe == probe & Level == "M")
  model_results_list[[paste(probe, "Tip_vs_Base_M", sep = "_")]] <- fit_and_tidy_model(current_data_middle, "Tip_vs_Base_M")
  
  # Fit models for Tip Bottom vs Base Bottom
  current_data_bottom <- combined_data %>%
    filter(Probe == probe & Level == "B")
  model_results_list[[paste(probe, "Tip_vs_Base_B", sep = "_")]] <- fit_and_tidy_model(current_data_bottom, "Tip_vs_Base_B")
}

# Create a new workbook
wb <- createWorkbook()

# Loop through model results and write each to a new sheet
for (model_name in names(model_results_list)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, model_results_list[[model_name]])
}

# Save the workbook
output_file_path <- "C:/Users/cbost/Desktop/BISC/Gabby_paper/7-5-24/GLMM_neg_binom_TMB_internal_area_aggregated_first_7-5-24.xlsx"
saveWorkbook(wb, output_file_path, overwrite = TRUE)


#######################################
######################################
########### By Cell Size ############
######################################
#######################################

# Aggregating function: Summarizes counts across locations for each Cell Size and Type
aggregate_counts <- function(df, type) {
  df %>%
    group_by(Cell_Size = `Cell Size`) %>%
    summarise_at(vars(starts_with(c("Tip", "Base"))), sum) %>%
    mutate(Type = type) %>%
    ungroup()
}

# Apply the function to each dataset and combine Tips and Bases
datasets <- list(
  TbH = bind_rows(aggregate_counts(data_list$`TbH Tips`, "Tip"), aggregate_counts(data_list$`TbH Base`, "Base")),
  TyH = bind_rows(aggregate_counts(data_list$`TyH Tips`, "Tip"), aggregate_counts(data_list$`TyH Base`, "Base")),
  Brady = bind_rows(aggregate_counts(data_list$`Brady Tips`, "Tip"), aggregate_counts(data_list$`Brady Base`, "Base")),
  TpH = bind_rows(aggregate_counts(data_list$`TpH Tips`, "Tip"), aggregate_counts(data_list$`TpH Base`, "Base")),
  FLRI = bind_rows(aggregate_counts(data_list$`FLRI Tips`, "Tip"), aggregate_counts(data_list$`FLRI Base`, "Base"))
)

# Combine all into a single dataframe
combined_data <- bind_rows(datasets, .id = "Group")

# Melt the dataframe to long format for GLMM
long_data <- combined_data %>%
  pivot_longer(cols = starts_with(c("Tip", "Base")), names_to = "Replicate", values_to = "Count") %>%
  mutate(Replicate = factor(Replicate))

# Fit GLMM (large being the default (alphabetically first) reference level)
glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = long_data, family = nbinom2(link = "log"))

# View the summary of the model
summary(glmm_size_model)

##########################################
#####################################
tbh_glmm_size <- long_data %>% filter(Group == "TbH")
tbh_glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tbh_glmm_size, family = nbinom2(link = "log"))
summary(tbh_glmm_size_model)
#####################################
#####################################
### model fails to converge since all XL counts are 0, if you remove then warnings disappear, but model results are the same
tyh_glmm_size <- long_data %>% filter(Group == "TyH")# %>% filter(Cell_Size  != "XL")
tyh_glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tyh_glmm_size, family = nbinom2(link = "log"))
summary(tyh_glmm_size_model)
#####################################
#####################################
tph_glmm_size <- long_data %>% filter(Group == "TpH")
tph_glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tph_glmm_size, family = nbinom2(link = "log"))
summary(tph_glmm_size_model)
#####################################
#####################################
brady_glmm_size <- long_data %>% filter(Group == "Brady")
brady_glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = brady_glmm_size, family = nbinom2(link = "log"))
summary(brady_glmm_size_model)
#####################################
#####################################
flri_glmm_size <- long_data %>% filter(Group == "FLRI")
flri_glmm_size_model <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = flri_glmm_size, family = nbinom2(link = "log"))
summary(flri_glmm_size_model)
#####################################
#####################################
size_models_list <- list(
  tbh_cell_size = tbh_glmm_size_model,
  tyh_cell_size = tyh_glmm_size_model,
  tph_cell_size = tph_glmm_size_model,
  brady_cell_size = brady_glmm_size_model,
  flri_cell_size = flri_glmm_size_model
)

# Tidy each model
size_models_tidied <- lapply(size_models_list, tidy)

# Create a new workbook
wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
#names(size_models_tidied) <- sapply(names(size_models_tidied), function(name) gsub(" ", "_", name))
for(model_name in names(size_models_tidied)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, size_models_tidied[[model_name]])
}

# Save the workbook
saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_cell_size_five_replicates_results_neg_binom_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)

##########################################
#########################################
########################################

# refit with small as the reference:
long_data_small_ref <- long_data %>%
  mutate(Cell_Size = factor(Cell_Size, levels = c("S", "M", "L", "XL"))) %>%
  mutate(Cell_Size = relevel(Cell_Size, ref = "S"))
#####################################
#####################################
tbh_glmm_size_small_ref <- long_data_small_ref %>% filter(Group == "TbH")
tbh_glmm_size_model_small_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tbh_glmm_size_small_ref, family = nbinom2(link = "log"))
summary(tbh_glmm_size_model_small_ref)
#####################################
#####################################
### model fails to converge since all XL counts are 0, if you remove then warnings disappear, but model results are the same
tyh_glmm_size_small_ref <- long_data_small_ref %>% filter(Group == "TyH")
tyh_glmm_size_model_small_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tyh_glmm_size_small_ref, family = nbinom2(link = "log"))
summary(tyh_glmm_size_model_small_ref)
#####################################
#####################################
tph_glmm_size_small_ref <- long_data_small_ref %>% filter(Group == "TpH")
tph_glmm_size_model_small_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tph_glmm_size_small_ref, family = nbinom2(link = "log"))
summary(tph_glmm_size_model_small_ref)
#####################################
#####################################
brady_glmm_size_small_ref <- long_data_small_ref %>% filter(Group == "Brady")
brady_glmm_size_model_small_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = brady_glmm_size_small_ref, family = nbinom2(link = "log"))
summary(brady_glmm_size_model_small_ref)
#####################################
#####################################
flri_glmm_size_small_ref <- long_data_small_ref %>% filter(Group == "FLRI")
flri_glmm_size_model_small_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = flri_glmm_size_small_ref, family = nbinom2(link = "log"))
summary(flri_glmm_size_model_small_ref)
#####################################
#####################################
size_models_list_small_ref <- list(
  tbh_cell_size = tbh_glmm_size_model_small_ref,
  tyh_cell_size = tyh_glmm_size_model_small_ref,
  tph_cell_size = tph_glmm_size_model_small_ref,
  brady_cell_size = brady_glmm_size_model_small_ref,
  flri_cell_size = flri_glmm_size_model_small_ref
)

# Tidy each model
size_models_tidied_small_ref <- lapply(size_models_list_small_ref, tidy)

# Create a new workbook
wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
#names(size_models_tidied_small_ref) <- sapply(names(size_models_tidied_small_ref), function(name) gsub(" ", "_", name))
for(model_name in names(size_models_tidied_small_ref)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, size_models_tidied_small_ref[[model_name]])
}

# Save the workbook
saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_cell_size_small_ref_five_replicates_results_neg_binom_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)
#####################################
#####################################
# refit with medium as the reference:
long_data_medium_ref <- long_data %>%
  mutate(Cell_Size = factor(Cell_Size, levels = c("S", "M", "L", "XL"))) %>%
  mutate(Cell_Size = relevel(Cell_Size, ref = "M"))
#####################################
#####################################
tbh_glmm_size_medium_ref <- long_data_medium_ref %>% filter(Group == "TbH")
tbh_glmm_size_model_medium_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tbh_glmm_size_medium_ref, family = nbinom2(link = "log"))
summary(tbh_glmm_size_model_medium_ref)
#####################################
#####################################
### model fails to converge since all XL counts are 0, if you remove then warnings disappear, but model results are the same
tyh_glmm_size_medium_ref <- long_data_medium_ref %>% filter(Group == "TyH")
tyh_glmm_size_model_medium_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tyh_glmm_size_medium_ref, family = nbinom2(link = "log"))
summary(tyh_glmm_size_model_medium_ref)
#####################################
#####################################
tph_glmm_size_medium_ref <- long_data_medium_ref %>% filter(Group == "TpH")
tph_glmm_size_model_medium_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tph_glmm_size_medium_ref, family = nbinom2(link = "log"))
summary(tph_glmm_size_model_medium_ref)
#####################################
#####################################
brady_glmm_size_medium_ref <- long_data_medium_ref %>% filter(Group == "Brady")
brady_glmm_size_model_medium_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = brady_glmm_size_medium_ref, family = nbinom2(link = "log"))
summary(brady_glmm_size_model_medium_ref)
#####################################
#####################################
flri_glmm_size_medium_ref <- long_data_medium_ref %>% filter(Group == "FLRI")
flri_glmm_size_model_medium_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = flri_glmm_size_medium_ref, family = nbinom2(link = "log"))
summary(flri_glmm_size_model_medium_ref)
#####################################
#####################################
size_models_list_medium_ref <- list(
  tbh_cell_size = tbh_glmm_size_model_medium_ref,
  tyh_cell_size = tyh_glmm_size_model_medium_ref,
  tph_cell_size = tph_glmm_size_model_medium_ref,
  brady_cell_size = brady_glmm_size_model_medium_ref,
  flri_cell_size = flri_glmm_size_model_medium_ref
)

# Tidy each model
size_models_tidied_medium_ref <- lapply(size_models_list_medium_ref, tidy)

# Create a new workbook
wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
#names(size_models_tidied_medium_ref) <- sapply(names(size_models_tidied_medium_ref), function(name) gsub(" ", "_", name))
for(model_name in names(size_models_tidied_medium_ref)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, size_models_tidied_medium_ref[[model_name]])
}

# Save the workbook
saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_cell_size_medium_ref_five_replicates_results_neg_binom_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)
#####################################
#####################################
# refit with extra large as the reference:
long_data_XL_ref <- long_data %>%
  mutate(Cell_Size = factor(Cell_Size, levels = c("S", "M", "L", "XL"))) %>%
  mutate(Cell_Size = relevel(Cell_Size, ref = "XL"))
#####################################
#####################################
tbh_glmm_size_XL_ref <- long_data_XL_ref %>% filter(Group == "TbH")
tbh_glmm_size_model_XL_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tbh_glmm_size_XL_ref, family = nbinom2(link = "log"))
summary(tbh_glmm_size_model_XL_ref)
#####################################
#####################################
### model fails to converge since all XL counts are 0, if you remove then warnings disappear, but model results are the same
tyh_glmm_size_XL_ref <- long_data_XL_ref %>% filter(Group == "TyH")
tyh_glmm_size_model_XL_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tyh_glmm_size_XL_ref, family = nbinom2(link = "log"))
summary(tyh_glmm_size_model_XL_ref)
#####################################
#####################################
tph_glmm_size_XL_ref <- long_data_XL_ref %>% filter(Group == "TpH")
tph_glmm_size_model_XL_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = tph_glmm_size_XL_ref, family = nbinom2(link = "log"))
summary(tph_glmm_size_model_XL_ref)
#####################################
#####################################
brady_glmm_size_XL_ref <- long_data_XL_ref %>% filter(Group == "Brady")
brady_glmm_size_model_XL_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = brady_glmm_size_XL_ref, family = nbinom2(link = "log"))
summary(brady_glmm_size_model_XL_ref)
#####################################
#####################################
flri_glmm_size_XL_ref <- long_data_XL_ref %>% filter(Group == "FLRI")
flri_glmm_size_model_XL_ref <- glmmTMB(Count ~ Cell_Size * Type + (1 | Replicate), data = flri_glmm_size_XL_ref, family = nbinom2(link = "log"))
summary(flri_glmm_size_model_XL_ref)
#####################################
#####################################
size_models_list_XL_ref <- list(
  tbh_cell_size = tbh_glmm_size_model_XL_ref,
  tyh_cell_size = tyh_glmm_size_model_XL_ref,
  tph_cell_size = tph_glmm_size_model_XL_ref,
  brady_cell_size = brady_glmm_size_model_XL_ref,
  flri_cell_size = flri_glmm_size_model_XL_ref
)

# Tidy each model
size_models_tidied_XL_ref <- lapply(size_models_list_XL_ref, tidy)

# Create a new workbook
wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
#names(size_models_tidied_XL_ref) <- sapply(names(size_models_tidied_XL_ref), function(name) gsub(" ", "_", name))
for(model_name in names(size_models_tidied_XL_ref)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, size_models_tidied_XL_ref[[model_name]])
}

# Save the workbook
saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_cell_size_XL_ref_five_replicates_results_neg_binom_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)
#####################################
#####################################
#####################################
#####################################

# Function to aggregate data within each probe
aggregate_data <- function(data, probe_name) {
  # Identify the location column dynamically
  location_col <- names(data)[grepl("Location$", names(data))]
  
  # Reshape data from wide to long
  data_long <- data %>%
    pivot_longer(
      cols = -c(location_col, `Cell Size`),  # Exclude location and cell size columns
      names_to = "Replicate",
      values_to = "Count"
    ) %>%
    mutate(Type = ifelse(grepl("^Tip_", Replicate), "Tip", "Base"),
           Probe = probe_name) %>%
    group_by(Probe, Type, Replicate) %>%
    summarise(Total_Count = sum(Count), .groups = "drop")
  
  return(data_long)
}

# List to store aggregated data for each probe
aggregated_data_list <- list()

# Names of the probes
probe_names <- c("TbH", "TyH", "TpH", "Brady", "FLRI")

# Loop through each probe and prepare data
for (i in seq_along(probe_names)) {
  probe_name <- probe_names[i]
  probe_data <- data_list[[paste0(probe_name, " Tips")]]
  base_data <- data_list[[paste0(probe_name, " Base")]]
  
  # Aggregate Tip and Base data
  tip_data_agg <- aggregate_data(probe_data, probe_name)
  base_data_agg <- aggregate_data(base_data, probe_name)
  
  # Combine Tip and Base data for the probe
  combined_data <- rbind(tip_data_agg, base_data_agg)
  
  # Store in list
  aggregated_data_list[[probe_name]] <- combined_data
}

# Combine all data into a single data frame
all_data <- bind_rows(aggregated_data_list)

# List to store models
models <- list()

# Fit a model for each probe
for (probe_name in probe_names) {
  # Subset data for the current probe
  data_subset <- subset(all_data, Probe == probe_name)
  
  # Fit the GLMM
  glmm_model <- glmmTMB(Total_Count ~ Type + (1 | Replicate),
                        data = data_subset,
                        family = nbinom2(link = "log"))
  # Store the model in the list using the probe name as the key
  models[[probe_name]] <- glmm_model
}

# Output the summary of each model
model_summaries <- lapply(names(models), function(probe_name) {
  summary(models[[probe_name]])
})

model_summaries

# Tidy each model
models_list_tidied <- lapply(models, tidy)

wb <- createWorkbook()

# Loop through tidied models and write each to a new sheet
for(model_name in names(models_list_tidied)) {
  addWorksheet(wb, model_name)
  writeData(wb, model_name, models_list_tidied[[model_name]])
}

# Save the workbook
saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_Probe_Type_five_replicates_results_neg_binom_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)
# saveWorkbook(wb, "C:/Users/cbost/Desktop/BISC/Gabby_paper/region_normalized/GLMM_Probe_Type_five_replicates_results_neg_binom_WholeCLNormperSQUAREmm_regionalnormArea_7-3-24.xlsx", overwrite = TRUE)
