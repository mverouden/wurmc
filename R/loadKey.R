#' Load the answer key file for grading.
#'
#' \code{loadKey()} loads a file with the exam answer key for all versions. This
#' file should be a comma or semicolon separated value file (csv) or a Microsoft
#' Excel file (xls or xlsx) with on its rows the exam versions (e.g. "A", "B",
#' "C", "D") and on its columns the correct responses to the multiple choice
#' test items (questions). The first column specifies the exam version (e.g.
#' "A", "B", "C", "D").
#'
#' @param keyFile A character string specifying the location and name of the csv
#'    or (x)lsx file containing the answer key to the exam items (questions).
#' @return A data.frame object with in the first column the exam version and
#'    the following columns giving the answers to the exam items (questions).
#'
#' @export
#'
#' @examples
#' \dontrun{
#' key <- loadKey(keyFile)}
loadKey <- function(keyFile) {
  validExt <- c("csv", "xls", "xlsx")
  ## Check whether a *.csv, *.xls or *.xlsx has been provided
  if (!(tolower(strsplit(x = keyFile, split = "\\.")[[1]][2]) %in% validExt)) {
    stop("Key not read, due to an invalid key file extension (should be either *.csv, *.xls or *.xlsx)")
  }
  ## Read the correct multiple-choice responses into a data.frame named key
  if (tolower(strsplit(x = keyFile,
                      split = "\\.")[[1]][2]) == validExt[1]) { # a csv-file
    ## Read the *.csv key file into a data.frame named key
    L <- readLines(con = keyFile, n = 1L)
    if (grepl(";", L)) {
      key <- read.csv2(file = keyFile)
    }else {
      key <- read.csv(file = keyFile)
    }
    ## Change all columns of the key data.frame to class character
    for (j in seq_len(ncol(key))) {
      key[, j] <- as.character(key[, j])
    }
  }else {
    if (tolower(strsplit(x = keyFile,
                            split = "\\.")[[1]][2]) %in% validExt[2:3]) { # a xls(x)-file
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
  }
  ## Set row- and column-names in the key data.frame
  rownames(key) <- LETTERS[1:nrow(key)]
  colnames(key) <- c("Version", paste0("I", seq_len(ncol(key) - 1)))
  return(key)
}
