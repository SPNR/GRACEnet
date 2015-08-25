xlsPath <- 'W:/GRACEnet/data summary project files/'
xlsInFile <- paste(xlsPath, 'GRACEnet_working_copy.xlsx', sep = '')

# Package openxlsx provides read/write functions for Excel files
# Use openxlsx for large xlsx files.
library(openxlsx)
summary <- openxlsx::read.xlsx(xlsInFile, sheet = 'DataCompilation-Soil')

expUnits <- openxlsx::read.xlsx(xlsInFile, sheet = 'ExperimentalUnits')
expUnits <- expUnits[, names(expUnits) %in% names(summary)]
expUnits <- unique(expUnits)

mgtAmends <- openxlsx::read.xlsx(xlsInFile, sheet = 'MgtAmendments')
mgtAmends <- mgtAmends[, names(mgtAmends) %in% names(summary)]
mgtAmends <- unique(mgtAmends)

measSoilChem <- openxlsx::read.xlsx(xlsInFile, sheet = 'MeasSoilChem')
measSoilChem <- measSoilChem[, names(measSoilChem) %in% names(summary)]
measSoilChem <- unique(measSoilChem)

measSoilPhys <- openxlsx::read.xlsx(xlsInFile, sheet = 'MeasSoilPhys')
measSoilPhys <- measSoilPhys[, names(measSoilPhys) %in% names(summary)]
measSoilPhys <- unique(measSoilPhys)


# Outer join
mDF1 <- merge(x = expUnits, y = mgtAmends, by = 'Experimental.Unit.ID',
              all = TRUE)
rm(summary, expUnits, mgtAmends)


mDF2 <- merge(x = mDF1, y = measSoilChem, by = 'Experimental.Unit.ID',
             all = TRUE)
rm(mDF1, measSoilChem)


mDF3 <- merge(x = mDF2, y = measSoilPhys, by = 'Experimental.Unit.ID',
              all = TRUE)



write.csv(mDF2, file = paste(xlsPath, 'mDF2a.csv', sep = ''))


----------------------------------------------

xlsOutFile <- paste(xlsPath, 'mDF2.xlsx', sep = '')
openxlsx::write.xlsx(mDF2, xlsOutFile)



