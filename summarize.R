################################################################################
#
#  This script merges several Excel worksheets from the GRACEnet database on
#  the variables specified in the worksheet titled DataCompilation-Soil
#
################################################################################

# Working copy of GRACEnet data file
xlsPath <- 'W:/GRACEnet/data summary project files/'
xlsInFile <- paste(xlsPath, 'GRACEnet_working_copy.xlsx', sep = '')

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

# Find the indices of duplicate rows
expUnitsDupInd <- which(duplicated(expUnits))
mgtAmendsDupInd <- which(duplicated(mgtAmends))
measSoilChemDupInd <- which(duplicated(measSoilChem))
measSoilPhysDupInd <- which(duplicated(measSoilPhys))
treatmentsDupInd <- which(duplicated(treatments))

# Compile a DF of unique observations representing mgtAmends duplicate rows
mgtAmendsDupDF <- unique(mgtAmends[mgtAmendsDupInd, ])
# Write DF to an Excel file
openxlsx::write.xlsx(mgtAmendsDupDF,
                     file = paste(xlsPath, 'duplicates.xlsx', sep = ''),
                     sheetName = 'MgtAmendsDups')
# Compile a DF of unique observations representing mgtAmends duplicate rows
measSoilPhysDupDF <- unique(measSoilPhys[measSoilPhysDupInd, ])
# Write DF to an Excel file
openxlsx::write.xlsx(measSoilPhysDupDF,
                     file = paste(xlsPath, 'duplicates.xlsx', sep = ''),
                     sheetName = 'SoilPhysDups')

# Remove duplicate rows
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
  
}

# Write the final merged DF to a csv file
write.csv(mDF4, file = paste(xlsPath, 'mDF4.csv', sep = ''))


#------- SUBSETTING ---------------------------------------------------

# Subset only those rows in which total soil carbon, inorganic soil carbon and
# bulk density all exist
carbonDF <- mDF4[!is.na(mDF4[, 15]) & !is.na(mDF4[, 17]) & !is.na(mDF4[, 61]), ]
# Write the final merged DF to a csv file
write.csv(carbonDF, file = paste(xlsPath, 'mDF4_TSC_ISC_present.csv',
                                  sep = ''))

# Now subset those carbonDF rows in which mineral associated carbon exists
soilPartCarbon <- carbonDF[!is.na(carbonDF[, 18]), ]
# Write the final merged DF to a csv file
write.csv(soilPartCarbon, file = paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv',
                                 sep = ''))

#-------------------------------------------------------------------------


xlsPath <- 'W:/GRACEnet/data summary project files/'
xlsInFile <- paste(xlsPath, 'mDF4_TSC_ISC_present.csv', sep = '')

# Use openxlsx for reading and writing large xlsx files.
library(openxlsx)
subset1 <- openxlsx::read.xlsx(xlsInFile)

subset1 <- read.csv(xlsInFile)
trtIdList <- unique(subset1$Treatment.ID)
for(trt in trtIdList) {
  trtSub <- subset1[subset1$Treatment.ID == trt, ]
  dateCount <- length(unique(** extracted date (or year) from datetime **))
}



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


