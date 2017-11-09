#' Hello, world!
#'
#' This is an example function named 'hello'
#' which prints 'Hello, world!'.
#'
#' You can learn more about package authoring with RStudio at:
#'
#' \url{http://r-pkgs.had.co.nz/}
#'
#' Some useful keyboard shortcuts for package authoring:
#'
#' Build and Reload Package:  "Cmd + Shift + B"
#'
#' Check Package:             "Cmd + Shift + E"
#'
#' Test Package:              "Cmd + Shift + T"
formScanner2Excel <- function(fileName,
                              noItems,
                              courseCode,
                              saveName = "ICT_results") {
  ## Check whether a *.csv has been provided
  if (!(tolower(strsplit(fileName, "\\.")[[1]][2]) == "csv")) {
    stop("FormScanner responses not read, due to an invalid responses file extension (should be *.csv file)")
  }
  ## Convert noItems into an integer
  noItems <- as.integer(noItems)
  ## Read the student multiple-choice responses into a data.frame responses
  responses <- read.csv2(file = fileName)
  ## Order responses by responses$File.name
  responses <- responses[order(responses$File.name), ]
  ## Generate the student registration number (reg.number) from columns
  ## ID_NO_001:ID_NO_012
  responses <- tidyr::unite(data = responses,
                            col = "Reg Nummer",
                            X.IDStudent..IDNo001:X.IDStudent..IDNo012, # to join
                            sep = "",
                            remove = TRUE)
  ## Select the columns reg.number, Version and the student responses to test
  ## items
  responses <- dplyr::select(responses,
                             "Reg Nummer",
                             X.Exam..Version,
                             grep("^X\\.Q", colnames(responses)))
  ## Select columns "reg.number","Version" and item responses equal to noItems
  ## in the key.
  responses <- responses[, seq_len(2 + noItems)]
  ## Change all columns of the responses data.frame to character
  for (j in seq_len(ncol(responses))) {
    responses[, j] <- as.character(responses[, j])
  }
  ## Fill empty fields with "BLANK"
  responses[!is.na(responses) & responses == ""] <- "BLANK"
  ## Change column names
  colnames(responses) <- c("Reg Nummer",
                           "Versie",
                           paste0("Q", seq_len(noItems)))
  ## Create a Microsoft Office Excel workbook with the specified saveename,
  ## the default is ICT_results.xlsx.
  wb <- XLConnect::loadWorkbook(filename = paste(saveName, "xlsx", sep = "."),
                                create = TRUE)
  ## Create a worksheet named after the supplied courseCode in the workbook wb
  XLConnect::createSheet(object = wb,
                         name = courseCode)
  ## Write the student responses to the worksheet courseCode in the workbook wb
  XLConnect::writeWorksheet(object = wb,
                            data = responses,
                            sheet = courseCode,
                            startRow = 1,
                            startCol = 1,
                            header = TRUE)
  ## Save the workbook to an Excel file
  return(XLConnect::saveWorkbook(wb))
}
