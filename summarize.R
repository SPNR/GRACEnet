################################################################################
#
#  This script merges several Excel worksheets from the GRACEnet database on
#  the variables specified in the worksheet titled DataCompilation-Soil
#
################################################################################

# Working copy of GRACEnet data file
xlsPath <- 'W:/GRACEnet/data summary project/'
xlsInFile <- paste(xlsPath, 'GRACEnet.xlsx', sep = '')

# Use openxlsx for reading and writing large xlsx files.
library(openxlsx)
summary <- openxlsx::read.xlsx(xlsInFile, sheet = 'DataCompilation-Soil')

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

# Find the indices of duplicate rows within each worksheet
expUnitsDupInd <- which(duplicated(expUnits))
mgtAmendsDupInd <- which(duplicated(mgtAmends))
measSoilChemDupInd <- which(duplicated(measSoilChem))
measSoilPhysDupInd <- which(duplicated(measSoilPhys))
treatmentsDupInd <- which(duplicated(treatments))

# Compile a DF of unique observations representing mgtAmends duplicate rows
mgtAmendsDupDF <- unique(mgtAmends[mgtAmendsDupInd, ])
# Write DF to an Excel file
openxlsx::write.xlsx(mgtAmendsDupDF,
                     file = paste(xlsPath, 'duplicates_original_mgmt.xlsx', sep = ''),
                     sheetName = 'MgtAmendsDups')

# Compile a DF of unique observations representing mgtAmends duplicate rows
measSoilPhysDupDF <- unique(measSoilPhys[measSoilPhysDupInd, ])
# Write DF to an Excel file
openxlsx::write.xlsx(measSoilPhysDupDF,
                     file = paste(xlsPath, 'duplicates_original_phys.xlsx', sep = ''),
                     sheetName = 'SoilPhysDups')

# Remove duplicate rows
expUnits <- unique(expUnits)
mgtAmends <- unique(mgtAmends)
measSoilChem <- unique(measSoilChem)
measSoilPhys <- unique(measSoilPhys)
treatments <- unique(treatments)

#------- Join on date when possible !!!!!!!!

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
  
}

# Write the final merged DF to a csv file
write.csv(mDF4, file = paste(xlsPath, 'mDF4.csv', sep = ''))


#------- SUBSETTING FOR CARBON DATA --------------------------------------------

# Subset only those rows in which total total soil carbon, inorganic soil carbon
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
xlsPath <- 'W:/GRACEnet/data summary project/'
xlsInFile <- paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv', sep = '')

# Read in csv format to avoid memory errors on a 4GB system
mdf4Subset <- read.csv(xlsInFile)

# Rename mdf4Subset depth columns for easier coding
names(mdf4Subset)[names(mdf4Subset) ==
                    'Upper.soil.layer..soil..centimeters'] <- 'depth.upper'
names(mdf4Subset)[names(mdf4Subset) ==
                    'Lower.soil.layer..soil..centimeters'] <- 'depth.lower'

# Define a DF 'baselineSub' whose rows will consist of Treatment ID + soil depth
# combos from mdf4Subset that have at least two sampling dates (so that delta C
# can be calculated)
#
baselineSub <- mdf4Subset[1, ]  # This replicates column names from mdf4Subset
baselineSub <- baselineSub[-c(1), ]  # Now delete the first (and only) row

# Remove rows where date = NA
mdf4Subset <- mdf4Subset[!is.na(mdf4Subset$Date), ]
# Form a list of unique Treatment IDs in mdf4Subset
trtIdList <- unique(mdf4Subset$Treatment.ID)
# Remove NA values from trtIdList
trtIdList <- trtIdList[!is.na(trtIdList)]

# For each Treatment ID...
for(trt in trtIdList) {
  
  # Subset rows by current trt
  trtSub <- mdf4Subset[mdf4Subset$Treatment.ID == trt, ]
  # Create a list of unique upper depths for current trt
  depthList <- unique(trtSub$depth.upper)
  # For each depth.upper value
  for(du in depthList) {
    
    # Form a list of unique text dates
    textDateList <- unique(trtSub$Date)
    # Identify earliest and latest dates
    firstDate <- min(as.Date(textDateList, format = '%m/%d/%Y'))
    lastDate <- max(as.Date(textDateList, format = '%m/%d/%Y'))
    # If first and last dates are the same...
    if(firstDate == lastDate) {
      # Find the row number for this combo of trt and du in mdf4Subset
      mdf4CurrentRow <- which(mdf4Subset$Treatment.ID == trt &
                                mdf4Subset$depth.upper == du)
      # Delete this row from mdf4Subset
      mdf4Subset <- mdf4Subset[-c(mdf4CurrentRow), ]
    # Else rbind to DF baselineSub
    }else {
      baselineSub <- rbind(baselineSub, trtSub[trtSub$depth.upper == du, ])
    }
    
  }  # End du for-loop
  
}  # End trt for-loop

write.csv(baselineSub, file =
            paste(xlsPath, 'multiple_date_trts.csv', sep = ''))


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


