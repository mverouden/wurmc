#' GUI: Conversion of FromScanner File into Microsoft Excel File.
#'
#' Graphical User Interface guided conversion of a FormScanner csv-file into a
#' Microsoft Excel xlsx-file.
#'
#' @return A Microsoft Excel workbook xlsx-file with the name specified in one
#'   of the GUI questions.
#'
#' @seealso \code{\link{formScanner2Excel}}
#'
#' @export
#' @import svDialogs
processScan <- function() {
  ## GUI: working directory selection
  setwd(svDialogs::dlgDir(default = getwd(), title = "Choose your working directory")$res)
  # GUI: selection of the FormScanner file with student responses
  fileName <- svDialogs::dlgOpen(title = "Select the FormScanner file containing the student responses")$res
  ## GUI: enter the number of items in the exam
  switch(tolower(Sys.info()[['sysname']]),
         windows = {noItems <- as.integer(svDialogs::dlgInput(message = "Enter the number of exam questions (max. 40 questions)",
                                                              default = c(25))$res)},
         linux   = {noItems <- as.integer(svDialogs::dlgInput(message  = "Enter the number of exam questions (ma. 40 questions)",
                                                              default = c(25))$res)},
         darwin  = {noItems <- as.integer(strsplit(x = svDialogs::dlgInput(message = "Enter the number of exam questions (max. 40 questions)",
                                                                           default = c(25))$res,
                                                   split = "button returned:OK, text returned:")[[1]][2])})
  ## GUI: enter a savename to which the student responses will be written
  switch(tolower(Sys.info()[['sysname']]),
         windows = {saveName <- svDialogs::dlgInput(message = "Enter a filename to save the processed student responses to an Excel file, e.g. 2014-2015_P6a_MAT15303_xm150608_ICT_results",
                                                    default = "ICT_results")$res},
         linux   = {saveName <- svDialogs::dlgInput(message = "Enter a filename to save the processed student responses to an Excel file, e.g. 2014-2015_P6a_MAT15303_xm150608_ICT_results",
                                                    default = "ICT_results")$res},
         darwin  = {saveName <- strsplit(x = svDialogs::dlgInput( message = "Enter a filename to save the processed student responses to an Excel file, e.g. 2014-2015_P6a_MAT15303_xm150608_ICT_results",
                                                              default = "ICT_results")$res,
                                         split = "button returned:OK, text returned:")[[1]][2]})
  ## GUI: enter a coursecode to identify from which course these student responses
  ## are. The coursecode will be used as identification for the worksheet in the
  ## created Microsoft Excel workbook.
  switch(tolower(Sys.info()[['sysname']]),
         windows = {courseCode <- svDialogs::dlgInput(message = "Enter a coursecode to identify from which course the student responses are, e.g. MAT15303",
                                                      default = "MATxxxxx")$res},
         linux   = {courseCode <- svDialogs::dlgInput(message = "Enter a coursecode to identify from which course the student responses are, e.g. MAT15303",
                                                      default = "MATxxxxx")$res},
         darwin  = {courseCode <- strsplit(x = svDialogs::dlgInput(message = "Enter a coursecode to identify from which course the student responses are e.g. MAT15303",
                                                                   default = "MATxxxxx")$res,
                                           split = "button returned:OK, text returned:")[[1]][2]})

  ## Write the processed FormScanner student responses to a Microsoft Excel file
  ## (*.xlsx file). The resulting file will be containing the columns:
  ##  the student registration number (Reg Nummer),
  ##  exam version (Versie), and
  ##  the responses to the multiple-choice test items (Q1 to maximally Q40)
  ##  depending on the number of test items in the exam (noItmes).
  formScanner2Excel(fileName, noItems, courseCode, saveName)
  ## Return nothing
  return(message(paste0("File saved as: ", getwd(), "/", saveName, ".xlsx")))
}
