#' @title
#' Determine Number of Correct Student Reponses
#'
#' @description
#' Determine the number of correct student responses to multiple choice exam
#' items (questions)
#'
#' @param responses A data.frame object (**REQUIRED**), as created with the
#'    \code{\link{loadResponses}()} function, with in the first column the
#'    student registration numbers of students taking the exam. The second column
#'    specifies the exam version the student filled. Subsequent columns give the
#'    filled in answers to the multiple choice exam test items (questions) as
#'    filled by the students, with a maximum of 40 test items (questions).
#' @param key A data.frame object (**REQUIRED**), as created with the
#'    \code{\link{loadKey}()} function, with in the first column the exam
#'    version and the subsequent columns giving the correct answers to the
#'    multiple choice exam test items (questions). The dimensions of the `key`
#'    data.frame object  are used in the grading process.
#'
#' @details
#' The\code{numberCorrect()} function checks the student's responses to the
#' multiple choice exam test items (questions) with the appropriate version
#' answer key and creates a column `No. Correct` indicating the number of test
#' items (questions) answered correctly by the particular student. The column
#' `No. Correct` is added after the last column in the `responses` data.frame
#' object.
#'
#' @return The input `responses` data.frame object extended with one column
#'    (`No. Correct`) indicating the number of multiple choice exam test items
#'    (questions) answered correctly by the students.
#'
#' @author Maikel Verouden
#'
#' @seealso \code{\link{loadResponses}()}, \code{\link{loadKey}()}
#'
#' @family grade functions
#'
#' @examples
#' ## Load the answer key
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/keyFile.xlsx"))
#'
#' ## Load the student responses from a Microsoft Excel
#' responses <- loadResponses(respFile = paste0(.libPaths()[1], "/wurmc/examples/respFile.xlsx"))
#'
#' ## Determine the number of correct student responses
#' responses <- numberCorrect(responses, key)
#'
#' @export
numberCorrect <- function(responses, key) {
  ## Check whether numberCorrect was already used
  if ("No. Correct" %in% colnames(responses)) {
    stop("Function terminated, because the numberCorrect() function was used before and, therefore, a column 'No. Correct' already exists in the responses data.frame object.")
  }
  ## define constants and variables
  noItems <- ncol(key) - 1
  noCorrect <- c(0)
  ## Count the number of correct responses to the multiple choice test items
  ## for each student
  for (i in seq_len(length.out = nrow(responses))) {
    noCorrect[i] <- sum(responses[i, seq(from = 3, to = 2 + noItems)] == key[responses[["Version"]][i], seq(from = 2, to = ncol(key))])
  }
  ## Add a new variable No.correct to the data.frame responses
  responses <- dplyr::mutate(responses, "No. Correct" = noCorrect)
  return(responses)
} # Checked and Done!

#' @title
#' Link Student Names
#'
#' @description
#' Link names of students and their study programme to the corrected studentnumbers.
#'
#' @param responses A data.frame object (**REQUIRED**), as created with the
#'    \code{\link{loadResponses}()} function and modified by the
#'    \code{\link{numberCorrect}()} function.
#' @param regFile A character string (**REQUIRED**) specifying the location and
#'    name of the `*.csv` or `.xls(x)` file containing the studentnumber, full
#'    name and study programme for all students, which registered for the exam
#'    via Osiris.
#'
#' @details
#' The \code{linkNames()} function compares the student registration numbers
#' scanned from the multiple choice answer sheets (filled by students during the
#' exam) with the student registration numbers from the list of students that registered
#' for the exam via Osiris. In case the registration number was incorrectly
#' filled on the answer sheet this function provides the opportunity to manually
#' correct this. From the Test participants per student group file exported from
#' Osiris the student's name and study programme are matched with the student
#' registration number and added to the responses object.
#'
#' @return The input \code{responses} data.frame object extended with three columns
#'    indicating the corrected studentnumber (`Corrected Studentnumber`),
#'    name (`Full Name`), and Study (`Programme`) of all students.
#'
#' @author Maikel Verouden
#'
#' @seealso \code{\link{loadKey}()}, \code{\link{loadResponses}()}
#'
#' @family grade functions
#'
#' @examples
#' ## Load the answer key
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/keyFile.xlsx"))
#'
#' ## Load the student responses from a Microsoft Excel
#' responses <- loadResponses(respFile = paste0(.libPaths()[1], "/wurmc/examples/respFile.xlsx"))
#'
#' ## Determine the number of correct student responses
#' responses <- numberCorrect(responses, key)
#'
#' ## Link Student Names
#' responses <- linkNames(responses, regFile = paste0(.libPaths()[1], "/wurmc/examples/regFile.xlsx"))
#'
#' @importFrom stats complete.cases
#'
#' @export
linkNames <- function(responses, regFile) {
  ## Check whether numberCorrect was already used
  if (!("No. Correct" %in% colnames(responses))) {
    stop("Function terminated, because the column 'No. Correct' does not exists in the responses object. Please use the numberCorrect() function before applying the linkNames() function.")
  }
  ## Check whether linkNames() was already used
  if (all(c("Corrected Studentnumber", "Full Name", "Programme") %in% colnames(responses))) {
    stop("Function terminated, because the linkNames function was used before and, therefore, the columns 'Corrected Studentnumber', 'Full Name', and 'Programme' already exist in the responses object.")
  }
  validExt <- c("csv", "xls", "xlsx")
  ## Check whether a *.csv, *.xls or *.xlsx has been provided
  if (!(tail(x = tolower(strsplit(x = regFile, split = "\\.")[[1]]), n = 1) %in% validExt)) {
    stop("List of students registered for the exam not read, due to an Test participants per student group file extension (should be either *.csv or *.xls(x)).")
  }
  ## Read the list of students that have registered for the exam
  if (tail(tolower(strsplit(x = regFile, split = "\\.")[[1]]), n = 1) == validExt[1]) {# a .csv file
    ## Read the *.csv file of registered students into data.frame registered
    firstLine <- readLines(con = regFile, n = 1L)
    if (grepl(";", firstLine)) {# TRUE then semicolon separated
      registered <- read.csv2(file = regFile)
    } else if (grepl(",", firstLine)) {# TRUE then comma separated
      registered <- read.csv(file = regFile)
    } else {# FALSE then neither semicolon nor comma separated
      stop("Test participants per student group file not read, due to an invalid invalid separator (should be either a comma or a semicolon) in the *.csv file.")
    }
  } else if (tail(tolower(strsplit(x = regFile, split = "\\.")[[1]]), n = 1) %in% validExt[2:3]) {# a .xls(x) file
    ## Read the *.xls(x) workbook file into a structure named wb
    wb <- XLConnect::loadWorkbook(filename = regFile,
                                    create = FALSE)
    ## Obtain the sheet names in the *.xls(x) workbook as represented by wb
    sheets <- XLConnect::getSheets(object = wb)
    ## Read the first sheet of wb into data.frame registered
    registered <- XLConnect::readWorksheet(object = wb,
                                           sheet = sheets[1],
                                           startRow = 8,
                                           endRow = 0,
                                           startCol = 0,
                                           endCol = 0)
    }
  ## Select only the "Studentnumber", "Full.Name" and "Programme"
  registered <- dplyr::select(registered, "Studentnumber", "Full.Name", "Programme")
  ## Change all columns of the registered data.frame to character
  for (j in seq_len(ncol(registered))) {
    registered[, j] <- as.character(registered[, j])
  }
  ## Turn Studentnumber column into numeric
  registered[["Studentnumber"]] <- as.numeric(registered[["Studentnumber"]])
  ## Delete rows with NA's
  registered <- registered[stats::complete.cases(registered), ]
  ## Rename the selected column names
  registered <- dplyr::rename(registered,
                              "Full Name" = "Full.Name")
  ## Order registered by "Full Name"
  registered <- registered[order(registered[["Full Name"]]),]
  rownames(registered) <- seq(from = 1, to = nrow(registered), by = 1)
  ## Create three new character empty variables
  correctedStudentnumber <- c("")
  name <- c("")
  programme <- c("")
  ## Check registration numbers of all students
  for (i in seq_len(nrow(responses))) {
    ## Find a match between the scanned studentnumber and the studentnumber in
    ## the Test participants per student group file (ordered by Full Name) as
    ## stored in the registered data.frame object and store the position in a
    ## selected object.
    selected <- which(registered[["Studentnumber"]] == responses[["Studentnumber"]][i])
    if (length(selected) == 0) { # no match found
      ## Display the Studentnumber for which no match in the registered
      ## data.frame object was found as well as the Full Name of the previous
      ## student, who has been matched. Correct the mistakes in the displayed
      ## Studentnumber manually.
      switch(tolower(Sys.info()[["sysname"]]),
             windows = {correctedStudentnumber[i] <- svDialogs::dlgInput(message = paste("Studentnumber (previous student: ",
                                                                                         ifelse(i - 1 == 0,
                                                                                                "first student",
                                                                                                name[i - 1]),
                                                                                         ")"),
                                                                         default = responses[["Studentnumber"]][i])$res},
             linux   = {correctedStudentnumber[i] <- svDialogs::dlgInput(message = paste("Studentnumber (previous student: ",
                                                                                         ifelse(i - 1 == 0,
                                                                                                "first student",
                                                                                                name[i - 1]),
                                                                                         ")"),
                                                                         default = responses[["Studentnumber"]][i])$res},
             darwin  = {correctedStudentnumber[i] <- strsplit(x = svDialogs::dlgInput(message = paste("Studentnumber (previous student: ",
                                                                                                      ifelse(i - 1 == 0,
                                                                                                             "first student",
                                                                                                             name[i - 1]),
                                                                                                      ")"),
                                                                                      default = responses[["Studentnumber"]][i])$res,
                                                              split = "button returned:OK, text returned:")[[1]][2]})
      ## Match the corrected student registration number with a student
      ## registration number in the list of students registered for the
      ## exam.
      selected <- which(registered[["Studentnumber"]] == correctedStudentnumber[i])
      if (length(selected) == 0) {
        ## No match found, either there was a mistake in the corrected student
        ## registration number or the student was not registered for the exam.
        name[i] <- "unknown"
        programme[i] <- "unknown"
      } else if (length(selected) == 1) {
       ## After correcting the student registration number a match with
       ## a student registered for the exam was found.
       name[i] <- registered[["Full Name"]][selected]
       programme[i] <- registered[["Programme"]][selected]
      }
    } else if (length(selected) == 1) {
      ## A match of the scanned student registration number with a student
      ## registered for the exam was found.
      correctedStudentnumber[i] <- registered[["Studentnumber"]][selected] # == responses[["Studentnumber"]][i]
      name[i] <- registered[["Full Name"]][selected]
      programme[i] <- registered[["Programme"]][selected]
    }
  }
  ## Add the corrected Studentnumber, Full Name and Programme to the responses
  ## data.frame as columns named Corrected.reg.number, Name and Study
  responses <- dplyr::mutate(responses, "Corrected Studentnumber" = correctedStudentnumber, "Full Name" = name, "Programme" = programme)
  return(responses)
} # Checked and Done!

#' @title
#' Grade Multiple Choice Exams
#'
#' @description
#' Grade Multiple Choice Exams with Correction for Guessing
#'
#' @param responses A data.frame object (**REQUIRED**), as created with the
#'    \code{\link{loadResponses}()} function modified by the
#'    \code{\link{numberCorrect}} and \code{\link{linkNames}} functions.
#' @param noItems An integer (**OPTIONAL**), not exceeding 40, specifying the
#'    number of multiple choice exam test items (questions) on the answers sheet
#'    to be considered. If a `key` object has been created with the
#'    \code{\link{loadKey}()} function prior to using the
#'    \code{\link{loadResponses}()} function the `noItems` function argument
#'    requires no specification and will filled automatically.
#' @param noOptions An integer (**OPTIONAL**), which specifies the number of
#'    options for each multiple choice exam test item (question). The default is
#'    4 options (answers "A", "B", "C", and "D").
#'
#' @details
#' The \code{gradeExam()} function calculates the exam grades achieved based on
#' the correct responses to the multiple choice exam test items (questions). The
#' exam grade is corrected for guessing.
#'
#' @return The `responses` data.frame object extended with a column giving the
#'    calculated exam grade (`Exam grade`) for each student.
#'
#' @author Maikel Verouden
#'
#' @seealso \code{\link{loadResponses}()}, \code{\link{loadKey}()}
#'
#' @family grade functions
#'
#' @examples
#' ## Load the answer key
#' key <- loadKey(keyFile = paste0(.libPaths()[1], "/wurmc/examples/keyFile.xlsx"))
#'
#' ## Load the student responses from a Microsoft Excel
#' responses <- loadResponses(respFile = paste0(.libPaths()[1], "/wurmc/examples/respFile.xlsx"))
#'
#' ## Determine the number of correct student responses
#' responses <- numberCorrect(responses, key)
#'
#' ## Link Student Names
#' responses <- linkNames(responses, regFile = paste0(.libPaths()[1], "/wurmc/examples/regFile.xlsx"))
#'
#' ## Grade Multiple Choice Exams
#' responses <- gradeExam(responses)
#'
#' @export
gradeExam <- function(responses, noItems = length(grep("^I", colnames(responses))), noOptions = 4) {
  ## Check whether gradeExam() was already used
  if ("Exam Grade" %in% colnames(responses)) {
    stop("Function terminated, because the gradeExam function was used before and, therefore, a column 'Exam Grade' already exists in the responses object.")
  }
  ## Check whether numberCorrect() was already used
  if (!("No. Correct" %in% colnames(responses))) {
    stop("Function terminated, because the column 'No. Correct' does not exists in the responses object. Please use the numberCorrect() function before applying the linkNames() function.")
  }
  ## Check whether linkNames() was already used
  if (!all(c("Corrected Studentnumber", "Full Name", "Programme") %in% colnames(responses))) {
    stop("Function terminated, because the columns 'Corrected Studentnumber', 'Full Name', and 'Programme' do not exist in the responses object. Please use the linkNames() function before applying the gradeExam() function.")
  }
  ## Save number of correct responses for convenience
  noCorrect <- responses[["No. Correct"]]
  ## Check whether the numberCorrect() function has been applied prior to
  ## calculating the exam grade i.e. that noCorrect is not NULL
  if (is.null(noCorrect)) {
    stop("The number of correct responses needs to be determined prior to grading the exam. Please apply the numberCorrect() function first.")
  }
  ## Calculate the exam grade with correction for guessing
  examGrade <- ifelse(test = (noCorrect > round(noItems / noOptions)),
                      yes = (1 + 9 * (noCorrect - round(noItems / noOptions)) / (noItems - round(noItems / noOptions))),
                      no = 1) # 1 when noCorrect <= round(noItems / noOptions)
  ## Add the exam grade to the responses data.frame as a column named "Grade exam"
  responses <- dplyr::mutate(responses, "Exam Grade" = examGrade)
  return(responses)
} # Checked and Done!
