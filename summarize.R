################################################################################
#
#  This script merges several Excel worksheets from the GRACEnet database on
#  the variables specified in the worksheet titled DataCompilation-Soil
#
################################################################################

# Working copy of GRACEnet data file
xlsPath <- 'W:/GRACEnet/data summary project/'
#xlsPath <- 'C:/Users/Robert/Documents/R/GRACEnet/'
xlsInFile <- paste(xlsPath, 'GRACEnet_working_copy.xlsx', sep = '')

# Use openxlsx for reading and writing large xlsx files.
library(openxlsx)
# summary <- openxlsx::read.xlsx(xlsInFile, sheet = 'DataCompilation-Soil')

# Read five different worksheets
#
expUnits <- openxlsx::read.xlsx(xlsInFile, sheet = 'ExperimentalUnits')
#expUnits <- expUnits[, names(expUnits) %in% names(summary)]
mgtAmends <- openxlsx::read.xlsx(xlsInFile, sheet = 'MgtAmendments')
#mgtAmends <- mgtAmends[, names(mgtAmends) %in% names(summary)]
measSoilChem <- openxlsx::read.xlsx(xlsInFile, sheet = 'MeasSoilChem')
#measSoilChem <- measSoilChem[, names(measSoilChem) %in% names(summary)]
measSoilPhys <- openxlsx::read.xlsx(xlsInFile, sheet = 'MeasSoilPhys')
#measSoilPhys <- measSoilPhys[, names(measSoilPhys) %in% names(summary)]
treatments <- openxlsx::read.xlsx(xlsInFile, sheet = 'Treatments')


#-------------- Debugging block 1 start -----------------------------------
#
# Find the indices of duplicate rows within each worksheet
# expUnitsDupInd <- which(duplicated(expUnits))
# mgtAmendsDupInd <- which(duplicated(mgtAmends))
# measSoilChemDupInd <- which(duplicated(measSoilChem))
# measSoilPhysDupInd <- which(duplicated(measSoilPhys))
# treatmentsDupInd <- which(duplicated(treatments))

# For the current project, only mgtAmends and measSoilPhys have duplicate rows
#
# Sys.setenv(R_ZIPCMD = "C:/Rtools/bin/zip.exe") - run this if zip error occurs
#
# Compile a DF of unique observations representing mgtAmends duplicate rows
# mgtAmendsDupDF <- unique(mgtAmends[mgtAmendsDupInd, ])
# Write DF to an Excel file
# openxlsx::write.xlsx(mgtAmendsDupDF,
#                      file = paste(xlsPath, 'duplicates_original_mgmt.xlsx',
#                                   sep = ''), sheetName = 'MgtAmendsDups')
# Compile a DF of unique observations representing mgtAmends duplicate rows
# measSoilPhysDupDF <- unique(measSoilPhys[measSoilPhysDupInd, ])
# Write DF to an Excel file
# openxlsx::write.xlsx(measSoilPhysDupDF,
#                      file = paste(xlsPath, 'duplicates_original_phys.xlsx',
#                                   sep = ''), sheetName = 'SoilPhysDups')
#
#-------------- Debugging block 1 end -----------------------------------

expUnits <- unique(expUnits)
mgtAmends <- unique(mgtAmends)
measSoilChem <- unique(measSoilChem)
measSoilPhys <- unique(measSoilPhys)
treatments <- unique(treatments)

# Perform a series of full outer joins on the five pertinent DFs
mDF1 <- merge(x = expUnits, y = mgtAmends,
              by = c('Experimental.Unit.ID', 'Treatment.ID'), all = TRUE)

mDF2 <- merge(x = mDF1, y = measSoilChem,
              by = c('Experimental.Unit.ID', 'Treatment.ID'), all = TRUE)

mDF3 <- merge(x = mDF2, y = measSoilPhys,
              by = c('Experimental.Unit.ID', 'Treatment.ID'), all = TRUE)

mDF4 <- merge(x = mDF3, y = treatments, by = 'Treatment.ID', all = TRUE)

# Add a column for state abbreviation
mDF4$State <- NA_character_

# For each row in mDF4...
for(i in 1:nrow(mDF4)) {
  
  # Replace a Treatment ID of 0 with NA
  if(!is.na(mDF4$Treatment.ID[i]) &
       mDF4$Treatment.ID[i] == 0) mDF4$Treatment.ID[i] <- NA
  # Replace an Experimental Unit ID of 0 with NA
  if(!is.na(mDF4$Experimental.Unit.ID[i]) &
       mDF4$Experimental.Unit.ID[i] == 0) mDF4$Experimental.Unit.ID[i] <- NA
  
  # If Treatment ID exists, subset the state abbreviation
  if(!is.na(mDF4$Treatment.ID[i])) {
    mDF4$State[i] <- substr(mDF4$Treatment.ID[i], 1, 2)
    # Else if Experimental Unit ID exists, subset the state abbreviation
  }else if(!is.na(mDF4$Experimental.Unit.ID[i])) {
    mDF4$State[i] <- substr(mDF4$Experimental.Unit.ID[i], 1, 2)
  }
  
  # Check for valid state abbreviation
  if(is.na(mDF4$State[i] %in% state.abb)) {
    msg <- paste('Invalid state abbreviation in row', i)
    stop(msg)
  }
  
}  # End for-loop

# Write the final merged DF to a csv file
write.csv(mDF4, file = paste(xlsPath, 'mDF4.csv', sep = ''))

# openxlsx::write.xlsx(mDF4, file = paste(xlsPath, 'mDF4.xlsx', sep = ''))


#------- SUBSETTING FOR CARBON DATA --------------------------------------------

# Create a subset for TSC and BD only.  If SIC value is missing, we will assume
# that SIC was negligibly small, implying that SOC ~ TSC.

# ++++++  code will go here



# Subset only those rows in which total soil carbon, inorganic soil carbon
# and bulk density all exist
carbonDF <- mDF4[!is.na(mDF4[, 15]) & !is.na(mDF4[, 17]) & !is.na(mDF4[, 61]), ]
# Write the final merged DF to a csv file
write.csv(carbonDF, file = paste(xlsPath, 'mDF4_TSC_ISC_present.csv',
                                  sep = ''))

# Now subset those carbonDF rows in which soil particulate carbon exists
soilPartCarbon <- carbonDF[!is.na(carbonDF[, 18]), ]
# Write the final merged DF to a csv file
write.csv(soilPartCarbon, file = paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv',
                                 sep = ''))

#------ SEARCH FOR MULTIPLE DATES -----------------------------------------
#
# This section of code looks for multiple dates within each Treatment ID because
# the earliest date will be considered as a baseline for carbon content.

# Define default path and input filename
# xlsPath <- 'W:/GRACEnet/data summary project/'
# xlsInFile <- paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv', sep = '')

# Read in csv format to avoid memory errors on a 4GB system
# carbonSubset <- read.csv(xlsInFile)

carbonSubset <- carbonDF  # Subset of mDF4 with values to calculate soil C

# Rename carbonSubset depth columns for easier coding
#
upperLayerOriginal <- 'Upper.soil.layer,.soil,.centimeters'
lowerLayerOriginal <- 'Lower.soil.layer,.soil,.centimeters'
names(carbonSubset)[names(carbonSubset) == upperLayerOriginal] <- 'Depth.upper'
names(carbonSubset)[names(carbonSubset) == lowerLayerOriginal] <- 'Depth.lower'


# Define a DF 'baselineSub' whose rows will consist of Exp Unit ID +
# Treatment ID + upper soil depth + lower soil depth combos from carbonSubset
# that have at least two sampling dates (so that delta C can be calculated)
#
# Replicate column names from carbonSubset
baselineSub <- carbonSubset[FALSE, ]

# The dplyr package provides filter()
library(dplyr)

# Remove rows where date = NA
carbonSubset <- filter(carbonSubset, !is.na(Date))

# Form a list of unique Treatment IDs in carbonSubset
trtIdList <- unique(carbonSubset$Treatment.ID)
# Remove NA values from trtIdList
trtIdList <- trtIdList[!is.na(trtIdList)]

# For each Treatment ID...
for(trt in trtIdList) {
  # Subset rows by current trt
  trtSub <- filter(carbonSubset, Treatment.ID == trt)
  # Create a list of unique exp unit ids for current trt
  expTrtList <- unique(trtSub$Experimental.Unit.ID)
  
  # For each exp value within current trt
  for(exp in expTrtList) {
    # Subset rows by current exp
    expTrtSub <- filter(trtSub, Experimental.Unit.ID == exp)
    # Form a list of unique depth.upper values
    depthUpExpTrtList <- unique(expTrtSub$Depth.upper)

    # For each depth.upper in current trt + exp combo
    for(du in depthUpExpTrtList) {
      # Subset rows by current du
      duExpTrtSub <- filter(expTrtSub, Depth.upper == du)
      # Form a list of unique depth.lower values
      depthLoDepthUpExpTrtList <- unique(duExpTrtSub$Depth.lower)
      
      # For each depth.lower in current trt + exp + du combo
      for(dl in depthLoDepthUpExpTrtList) {
        dlDuExpTrtSub <- filter(duExpTrtSub, Depth.lower == dl)
        # If earliest and latest dates are not equal
        if(min(dlDuExpTrtSub$Date) != max(dlDuExpTrtSub$Date)) {
          # Sort df by date
          dlDuExpTrtSub <-  arrange(dlDuExpTrtSub, Date)
          # Add first and last rows to baselineSub
          baselineSub <- rbind(baselineSub, dlDuExpTrtSub[1, ],
                               dlDuExpTrtSub[nrow(dlDuExpTrtSub), ])
        }  # End if
      }  # End dl for-loop
    }  # End du for-loop
  }  # End exp for-loop
}  # End trt for-loop


# Calculate SOC stocks
#
# Rename key columns for conciseness
names(baselineSub)[15] <- 'Total.soil.carbon'
names(baselineSub)[16] <- 'Total.soil.nitrogen'
names(baselineSub)[17] <- 'Inorganic.soil.carbon'
names(baselineSub)[61] <- 'Bulk.density'

# Coerce pertinent quantities to numeric
baselineSub$Total.soil.carbon <- as.numeric(baselineSub$Total.soil.carbon)
baselineSub$Inorganic.soil.carbon <-
  as.numeric(baselineSub$Inorganic.soil.carbon)
baselineSub$Bulk.density <- as.numeric(baselineSub$Bulk.density)
baselineSub$Total.soil.nitrogen <- as.numeric(baselineSub$Total.soil.nitrogen)

# Calculate soil organic carbon
baselineSub$Soil.organic.carbon <- baselineSub$Total.soil.carbon -
  baselineSub$Inorganic.soil.carbon

# Calculate soil organic carbon stocks
baselineSub$Soil.organic.carbon.stocks <- baselineSub$Soil.organic.carbon *
  baselineSub$Bulk.density * (baselineSub$Depth.lower - baselineSub$Depth.upper)

# Calculate soil N stocks
baselineSub$Soil.nitrogen.stocks <- baselineSub$Total.soil.nitrogen *
  baselineSub$Bulk.density * (baselineSub$Depth.lower - baselineSub$Depth.upper)

# Provides functions for date arithmetic
library(lubridate)

# Coerce text dates to 'Date' class
baselineSub$Date <- as.Date(baselineSub$Date, format = '%m/%d/%Y')

# Create columns for delta C values
baselineSub$Delta.soil.organic.carbon.stocks <- NA_real_
baselineSub$Yearly.delta.soil.organic.carbon.stocks <- NA_real_

# Calculate change in soil organic carbon stocks
for(i in seq(2, nrow(baselineSub), by =2)) {
  # Absolute change
  baselineSub$Delta.soil.organic.carbon.stocks[i] <-
    baselineSub$Soil.organic.carbon.stocks[i] -
    baselineSub$Soil.organic.carbon.stocks[i - 1]
  # Number of years in study (fractional)
  numberOfYears <- (baselineSub$Date[i] - baselineSub$Date[i - 1]) / eyears(1)
  # Yearly change
  baselineSub$Yearly.delta.soil.organic.carbon.stocks[i] <-
    baselineSub$Delta.soil.organic.carbon.stocks[i] / numberOfYears
}

# Restore original column names to baselineSub
names(baselineSub)[names(baselineSub) == 'Depth.upper'] <- upperLayerOriginal
names(baselineSub)[names(baselineSub) == 'Depth.lower'] <- lowerLayerOriginal
names(baselineSub)[15] <- names(carbonSubset)[15]
names(baselineSub)[16] <- names(carbonSubset)[16]
names(baselineSub)[17] <- names(carbonSubset)[17]
names(baselineSub)[61] <- names(carbonSubset)[61]


write.csv(baselineSub, file = paste(xlsPath, 'baselineSub.csv'))




#  ----------------------------------------------
#
# Troubleshooting steps below are not a part of this script
  
df1 <- data.frame(col1 = c(1, 3, 3), col2 = c(4, 5, NA))
df2 <- data.frame(col1 = c(3, 3, 3), col3 = c(6, 7, 8))

mDF1 <- merge(x = df1, y = df2, by = 'col1', all = TRUE)

# Create a sorted table of the three most frequent exp unit IDs in mDF1
sort(table(mDF1$Experimental.Unit.ID),decreasing=TRUE)[1:3]

# Subset rows containing the most frequent exp unit ID
subDF <- mDF1[mDF1$Experimental.Unit.ID == 'COARDEC3_ST-N6', ]

# Sort that subset by exp unit ID (not necessary: only one value)
subDFsort <- mDF1[with(mDF1, order(Experimental.Unit.ID)), ]


