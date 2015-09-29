################################################################################
#
#  This script merges several Excel worksheets from the GRACEnet database on
#  the variables specified in the worksheet titled DataCompilation-Soil
#
################################################################################

# Working copy of GRACEnet data file
xlsPath <- 'W:/GRACEnet/soil carbon project/'
#xlsPath <- 'C:/Users/Robert/Documents/R/GRACEnet/'
xlsInFile <- paste(xlsPath, 'GRACEnet_working_copy_lite.xlsx', sep = '')

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

# Keep only unique rows
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
# for(i in 1:nrow(mDF4)) {
#   #------- Not sure why I'm checking for zero values here; class is character!
#   # Replace a Treatment ID of 0 with NA
#   if(!is.na(mDF4$Treatment.ID[i]) &
#        mDF4$Treatment.ID[i] == 0) mDF4$Treatment.ID[i] <- NA
#   # Replace an Experimental Unit ID of 0 with NA
#   if(!is.na(mDF4$Experimental.Unit.ID[i]) &
#        mDF4$Experimental.Unit.ID[i] == 0) mDF4$Experimental.Unit.ID[i] <- NA


# Create progress bar for following for-loop
pb <- winProgressBar(title = "progress bar", min = 0,
                     max = nrow(mDF4), width = 300)

for(i in 1:nrow(mDF4)) {  
    mDF4$State[i] <- substr(mDF4$Treatment.ID[i], 1, 2)
    # Else if Experimental Unit ID exists, subset the state abbreviation

  # Check for valid state abbreviation
  if(is.na(mDF4$State[i] %in% state.abb)) {
    msg <- paste('Invalid state abbreviation in row', i)
    stop(msg)
  }
    Sys.sleep(0.1)
    setWinProgressBar(pb, i, title=paste( round(i/nrow(mDF4)*100, 0),
                                          "% done"))
}
close(pb)  # Close progress bar


# Write the final merged DF to a csv file
write.csv(mDF4, file = paste(xlsPath, 'mDF4.csv', sep = ''))

tempPath <- 'E:/'
write.csv(mDF4, file = paste(tempPath, 'mDF4.csv', sep = ''))
# openxlsx::write.xlsx(mDF4, file = paste(xlsPath, 'mDF4.xlsx', sep = ''))


#------- SUBSETTING FOR CARBON DATA --------------------------------------------
#
# Choose one of the following three subsets:

# Create a subset for TSC and BD only.  If SIC value is missing, we will assume
# that SIC was measured to be negligibly small, implying that SOC ~ TSC.
mDF4_TSC_BD <- mDF4[!is.na(mDF4[, 17]) & !is.na(mDF4[, 63]), ]
tempPath <- 'E:/'
write.csv(mDF4_TSC_BD, file = paste(tempPath, 'mDF4_TSC_BD_present.csv',
                                 sep = ''))

# # Subset only those rows in which total soil carbon, inorganic soil carbon
# # and bulk density all exist
# carbonDF <- mDF4[!is.na(mDF4[, 15]) & !is.na(mDF4[, 17]) & !is.na(mDF4[, 61]), ]
# # Write the final merged DF to a csv file
# write.csv(carbonDF, file = paste(xlsPath, 'mDF4_TSC_ISC_present.csv',
#                                   sep = ''))
# 
# # Subset only those rows in which total soil carbon, inorganic soil carbon,
# # soil particulate carbon and bulk density all exist
# soilPartCarbon <- carbonDF[!is.na(carbonDF[, 18]), ]
# # Write the final merged DF to a csv file
# write.csv(soilPartCarbon, file = paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv',
#                                  sep = ''))


#------ SEARCH FOR MULTIPLE DATES -----------------------------------------
#
# This section of code looks for multiple dates within each Treatment ID because
# the earliest date will be considered as a baseline for carbon content.

# Define default path and input filename
# xlsPath <- 'W:/GRACEnet/data summary project/'
# xlsInFile <- paste(xlsPath, 'mDF4_TSC_ISC_SPC_present.csv', sep = '')

# Read in csv format to avoid memory errors on a 4GB system
# carbonSubset <- read.csv(xlsInFile)

carbonSubset <- mDF4_TSC_BD  # Subset of mDF4 with values to calculate soil C

# Rename carbonSubset depth columns for easier coding
#
# Depth descriptor from measSoilChem
upperLayerOriginal <- 'Upper.soil.layer,.soil,.centimeters'
lowerLayerOriginal <- 'Lower.soil.layer,.soil,.centimeters'
names(carbonSubset)[names(carbonSubset) ==
                      upperLayerOriginal] <- 'Depth.upper.chem'
names(carbonSubset)[names(carbonSubset) ==
                      lowerLayerOriginal] <- 'Depth.lower.chem'

# Depth descriptor from measSoilPhys
upperHorizonOriginal <- 'Soil.horizon.depth,.upper,.centimeter'
lowerHorizonOriginal <- 'Soil.horizon.depth,.lower,.centimeter'
names(carbonSubset)[names(carbonSubset) ==
                      upperHorizonOriginal] <- 'Depth.upper.phys'
names(carbonSubset)[names(carbonSubset) ==
                      lowerHorizonOriginal] <- 'Depth.lower.phys'





# Define a DF 'baselineSub' whose rows will consist of Exp Unit ID +
# Treatment ID + upper soil depth + lower soil depth combos from carbonSubset
# that have at least two sampling dates (so that delta C can be calculated)
#
# Replicate column names from carbonSubset
baselineSub <- carbonSubset[FALSE, ]

# The dplyr package provides filter() and other functions
library(dplyr)

# Keep only rows where date is not NA, and TSC and BD are both > 0
carbonSubset <- filter(carbonSubset, !is.na(Date) & carbonSubset[, 15] > 0 &
                         carbonSubset[, 61] > 0)

# Form a list of unique Treatment IDs in carbonSubset
trtIdList <- unique(carbonSubset$Treatment.ID)
# Remove NA values from trtIdList
trtIdList <- trtIdList[!is.na(trtIdList)]
# Coerce text dates to 'Date' class
carbonSubset$Date <- as.Date(carbonSubset$Date, format = '%m/%d/%Y')

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


# Calculate organic soil carbon stocks
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
#
# If inorganic soil carbon is NA or nonpositive, assume OSC = TSC
if(is.na(baselineSub$Inorganic.soil.carbon) |
   baselineSub$Inorganic.soil.carbon <= 0) {
  baselineSub$Organic.soil.carbon <- baselineSub$Total.soil.carbon

# Else organic soil carbon = total soil carbon - inorganic soil carbon
} else {
  baselineSub$Organic.soil.carbon <- baselineSub$Total.soil.carbon -
    baselineSub$Inorganic.soil.carbon
}

# Calculate soil organic carbon stocks
baselineSub$Organic.soil.carbon.stocks <- baselineSub$Organic.soil.carbon *
  baselineSub$Bulk.density * (baselineSub$Depth.lower -
                                baselineSub$Depth.upper) * 100

# Calculate soil N stocks
baselineSub$Soil.nitrogen.stocks <- baselineSub$Total.soil.nitrogen *
  baselineSub$Bulk.density * (baselineSub$Depth.lower -
                                baselineSub$Depth.upper) * 100

# Create columns for delta C values
baselineSub$Delta.organic.soil.carbon.stocks <- NA_real_
baselineSub$Yearly.delta.organic.soil.carbon.stocks <- NA_real_

# Create columns for start date and end date
baselineSub$Start.date <- structure(NA, class = 'Date')
baselineSub$End.date <- structure(NA, class = 'Date')

# Provides functions for date arithmetic
library(lubridate)

# Calculate change in soil organic carbon stocks
for(i in seq(2, nrow(baselineSub), by =2)) {
  # Absolute change
  baselineSub$Delta.organic.soil.carbon.stocks[i] <-
    baselineSub$Organic.soil.carbon.stocks[i] -
    baselineSub$Organic.soil.carbon.stocks[i - 1]
  # Number of years in study (fractional)
  numberOfYears <- (baselineSub$Date[i] - baselineSub$Date[i - 1]) / eyears(1)
  # Yearly change
  baselineSub$Yearly.delta.organic.soil.carbon.stocks[i] <-
    baselineSub$Delta.organic.soil.carbon.stocks[i] / numberOfYears
  # Populate start date and end date fields
  baselineSub$Start.date[i - 1] <- baselineSub$Date[i - 1]
  baselineSub$End.date[i - 1] <- baselineSub$Date[i]
  baselineSub$Start.date[i] <- baselineSub$Date[i - 1]
  baselineSub$End.date[i] <- baselineSub$Date[i]
}

# Restore original column names to baselineSub
names(baselineSub)[names(baselineSub) == 'Depth.upper'] <- upperLayerOriginal
names(baselineSub)[names(baselineSub) == 'Depth.lower'] <- lowerLayerOriginal
names(baselineSub)[15] <- names(carbonSubset)[15]
names(baselineSub)[16] <- names(carbonSubset)[16]
names(baselineSub)[17] <- names(carbonSubset)[17]
names(baselineSub)[61] <- names(carbonSubset)[61]

# Add units to new column headings
names(baselineSub)[names(baselineSub) == 'Organic.soil.carbon'] <-
  'Organic.soil.carbon,.grams.C.per.kilogram.soil'
names(baselineSub)[names(baselineSub) == 'Organic.soil.carbon.stocks'] <-
  'Organic.soil.carbon.stocks,.grams.C.per.kilogram.soil'
names(baselineSub)[names(baselineSub) == 'Soil.nitrogen.stocks'] <-
  'Soil.nitrogen.stocks,.grams.N.per.kilogram.soil'
names(baselineSub)[names(baselineSub) == 'Delta.organic.soil.carbon.stocks'] <-
  'Delta.organic.soil.carbon.stocks,.grams.C.per.kilogram.soil'
names(baselineSub)[names(baselineSub) ==
                     'Yearly.delta.organic.soil.carbon.stocks'] <-
  'Yearly.delta.organic.soil.carbon.stocks,.grams.C.per.kilogram.soil.per.year'

# Save analysis results
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


