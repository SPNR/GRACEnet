xlsFile <- 'C:/Users/Robert/Downloads/GRACEnet.xlsx'

# Package openxlsx provides read/write functions for Excel files
# Use openxlsx for large xlsx files.
library(openxlsx)
summarySheet <- openxlsx::read.xlsx(xlsFile, sheet = 'DataCompilation-Soil',
                                    colNames = FALSE)

for(i in ncol(summarySheet)) {
  
}
# This list contains duplicates
sheetNames <- summarySheet[2, ]


# strategy:
#   1. Row 1 in the DataCompilation-Soil tab contains the sheet names for
# quantities we need
# 2. Row 2 contains the sheet name where each item in row 1 is located



#colClassVector <- templateSheet2[, 2]

names(colClassVector) <- templateSheet2[, 1]
# Now read worksheet 1, using colClassVector
# templateSheet <- read.xlsx2(filename, sheetIndex = 1,
#                             colClasses = colClassVector,
#                             stringsAsFactors = FALSE)

