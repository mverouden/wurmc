#' @title
#' Load Answer Key File
#'
#' @description
#' Load a file with the multiple choice exam answer key needed for grading.
#'
#' @param keyFile A character string (**REQUIRED**) specifying the location and
#'    name of the `.csv` or `.xls(x)` file containing the answer key to the exam
#'    items (questions).
#'
#' @details
#' The `loadKey()` function loads a file with the exam answer key for all versions
#' needed to grade the multiple choice answer forms handed in by students. This
#' file should be a comma or semicolon separated value file, having the extension
#' `.csv`, or a Microsoft Excel file, having the extension `.xls(x)`, with on its
#' rows the exam versions (e.g. "A", "B", "C", "D") and on its columns the
#' correct responses to the multiple choice test items (questions).
#'
#' The `examples` directory of this package contains a Microsoft Excel Template
#' named **`key_template.xltx`** to aid the creation of the answer key file.
#'
#' **NOTE: The answer key file should be cropped to the number of required**
#' **versions as well as to the required number of test items (questions).**
#'
#' The first rows of the file specifies the column names:
#'
#'   * Column 1 should be named __Version__
#'   * Subsequent columns denotes the multiple choice items example given
#'     __Q01__, __Q002__, _etc._
#'
#'  Rows following the first provide the answer key for each multiple choice
#'  exam required and starts with the version (e.g. "A", "B", "C", "D", _etc._).
#'
#' @return
#' A data.frame object with in the first column the version and the subsequent
#' columns giving the correct answers to the multiple choice exam test items
#' (questions). The dimensions of the data.frame are used in the grading process.
#'
#' @author Maikel Verouden
#'
#' @family load functions
#'
#' @examples
#' ## Load the answer key from a Microsoft Excel file created with the Microsoft
#' ## Excel Template key_template contained in the examples folder of this package
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/test_key.xlsx"))
#'
#' ## Load the answer key from a comma separated value file
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/test_key.csv"))
#'
#' @importFrom utils read.csv read.csv2 tail
#'
#' @export
loadKey <- function(keyFile) {
  validExt <- c("csv", "xls", "xlsx")
  ## Check whether a *.csv, *.xls or *.xlsx has been provided
  if (!(tail(x = tolower(strsplit(x = keyFile, split = "\\.")[[1]]), n = 1) %in% validExt)) {
    stop("Key not read, due to an invalid key file extension (should be either *.csv, *.xls or *.xlsx).")
  }
  ## Read the correct multiple-choice responses into a data.frame named key
  if (tail(tolower(strsplit(x = keyFile, split = "\\.")[[1]]), n = 1) == validExt[1]) {# a .csv file
    ## Read the *.csv key file into a data.frame named key
    firstLine <- readLines(con = keyFile, n = 1L)
    if (grepl(";", firstLine)) {# TRUE then semicolon separated
      key <- read.csv2(file = keyFile)
    } else if (grepl(",", firstLine)) {# TRUE then comma separated
      key <- read.csv(file = keyFile)
    } else {# FALSE then neither semicolon nor comma separated
      stop("Key not read, due to an invalid invalid separator (should be either a comma or a semicolon) in the *.csv file.")
    }
    ## Change all columns of the key data.frame to class character
    for (j in seq_len(ncol(key))) {
      key[, j] <- as.character(key[, j])
    }
  } else if (tail(tolower(strsplit(x = keyFile, split = "\\.")[[1]]), n = 1) %in% validExt[2:3]) { # a .xls(x) file
    ## Read the *.xls(x) workbook file into a structure named wb
    wb <- XLConnect::loadWorkbook(filename = keyFile,
                                  create = FALSE)
    ## Obtain the sheet names in the *.xls(x) workbook as represented by wb
    sheets <- XLConnect::getSheets(object = wb)
    ## Read the first sheet of wb into a data.frame named key
    key <- XLConnect::readWorksheet(object = wb,
                                    sheet = sheets[1],
                                    startRow = 0,
                                    endRow = 0,
                                    startCol = 0,
                                    endCol = 0)
  }
  ## Set row- and column-names in the key data.frame
  rownames(key) <- LETTERS[1:nrow(key)]
  colnames(key) <- c("Version", paste0("I", seq_len(ncol(key) - 1)))
  return(key)
} # Checked and Done!

#' @title
#' Load Student Responses File
#'
#' @description
#' Loads a file containing the student responses to the multiple choice exam
#' test items (questions).
#'
#' @param respFile  A character string (**REQUIRED**) specifying the location
#'    and name of the `.csv` or `.xls(x)` file containing the student responses
#'    to the multiple choice exam test items (questions).
#' @param noItems An integer (**OPTIONAL**) specifying the number of multiple
#'    choice exam test items (questions) on the answers sheet to be considered
#'    and matching the number of test items (questions) in the answer key file.
#'    If a `key` object has been created with the \code{\link{loadKey}()}
#'    function prior to using the `loadResponses()` function the `noItems`
#'    function argument requires no specification and will be filled
#'    automatically.
#'
#' @details
#' The `loadResponses()` function loads a file with the student responses to the
#' multiple choice exam test items (questions). This file should be a FormScanner
#' (\url{http://www.formscanner.org/}) generated comma or semicolon separated
#' value `.csv` file or a Microsoft Excel `.xls(x)` file generated by Wageningen
#' University & Research EDUsuport, or by means of the
#' \code{\link{formScanner2Excel}()} function in this package.
#'
#' The Microsoft Excel `.xls(x)` file has on each row for each student in the
#' following particular order:
#'
#'   1. the student registration number,
#'   2. the exam version,
#'   3. the responses to the multiple choice exam test items (questions).
#'
#' @return A data.frame object with in the first column the student registration
#'    numbers of students taking the exam. The second column specifies the exam
#'    version the student filled. Subsequent columns give the filled in answers
#'    to the multiple choice exam test items (questions) as filled by the
#'    students, with a maximum of 40 test items (questions).
#'
#' @author Maikel Verouden
#'
#' @family load functions
#'
#' @examples
#' ## Load Answer Key File
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/keyFile.xlsx"))
#'
#' ## Load the student responses from a Microsoft Excel file
#' responses <- loadResponses(respFile = paste0(.libPaths()[1], "/wurmc/examples/respFile.xlsx"))
#'
#' ## Load the student responses from a comma separated value file as created by
#' ## FormScanner
#' responses <- loadResponses(respFile = paste0(.libPaths()[1], "/wurmc/examples/respFile.csv"))
#'
#' @export
loadResponses <- function(respFile, noItems = ncol(key) - 1) {
  validExt <- c("csv", "xls", "xlsx")
  ## Check whether a *.csv, *.xls or *.xlsx has been provided
  if (!(tail(x = tolower(strsplit(x = respFile, split = "\\.")[[1]]), n = 1) %in% validExt)) {
    stop("Responses not read, due to an invalid responses file extension (should be either *.csv, *.xls or *.xlsx)")
  }
  ## Read the student multiple-choice responses into a data.frame named responses
  if (tail(tolower(strsplit(x = respFile, split = "\\.")[[1]]), n = 1) == validExt[1]) {# a .csv file
    ## Read the *.csv responses file into a data.frame named responses
    firstLine <- readLines(con = respFile, n = 1L)
    if (grepl(";", firstLine)) {# TRUE then semicolon separated
      responses <- read.csv2(file = respFile)
    } else if (grepl(",", firstLine)) {# TRUE then comma separated
      responses <- read.csv(file = respFile)
    } else {# FALSE then neither semicolon nor comma separated
      stop("Student responses not read, due to an invalid invalid separator (should be either a comma or a semicolon) in the *.csv file.")
    }
    ## Order by responses$File.name
    responses <- responses[order(responses[["File.name"]]), ]
    ## Generate the student registration number (Studentnumber) from columns ID_NO_001:ID_NO_007
    responses <- tidyr::unite(data = responses,
                              col = "Studentnumber",
                              "X.IDStudent..IDNo001":"X.IDStudent..IDNo007",
                              sep = "",
                              remove = TRUE)
    ## Select the columns Studentnumber, Version and the student responses to
    ## the test items (questions)
    responses <- dplyr::select(responses,
                               "Studentnumber",
                               "X.Exam..Version",
                               grep("^X\\.Q", colnames(responses)))
    ## Select the columns "Studentnumber", "Version" and item responses equal to
    ## noItems in the key.
    responses <- responses[, seq_len(2 + noItems)]
    ## Change all columns of the responses data.frame to character
    for (j in seq_len(ncol(responses))) {
      responses[, j] <- as.character(responses[, j])
    }
    ## Turn Studentnumber column into numeric
    responses[["Studentnumber"]] <- as.numeric(responses[["Studentnumber"]])
    ## Fill empty fields with "BLANK"
    responses[!is.na(responses) & responses == ""] <- "BLANK"
  } else if (tail(tolower(strsplit(x = respFile, split = "\\.")[[1]]), n = 1) %in% validExt[2:3]) {# a .xls(x) file
    ## Read the *.xls(x) workbook file into a structure named wb
    wb <- XLConnect::loadWorkbook(filename = respFile,
                                  create = FALSE)
    ## Obtain the sheet names in the *.xls(x) workbook as represented by wb
    sheets <- XLConnect::getSheets(object = wb)
    ## Read the first sheet of wb into a data.frame named responses
    responses <- XLConnect::readWorksheet(object = wb,
                                          sheet = sheets[1],
                                          startRow = 0,
                                          endRow = 0,
                                          startCol = 0,
                                          endCol = 0)
  }
  ## Change column names
  colnames(responses) <- c("Studentnumber", "Version", paste0("I", seq_len(noItems)))
  return(responses)
} # Checked and Done!
