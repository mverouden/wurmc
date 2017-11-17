#' Determine the Number of Correct Student Responses
#'
#' \code{numberCorrect()} checks the student's responses to the multiple choice
#' exam test items (questions) with the appropriate version key and creates a
#' column \code{No.correct} indicating the number of test items answered
#' correctly by a student. The column \code{No.correct} is added after the last
#' column of the data.frame \code{responses}.
#'
#' @param responses A data.frame object, as created with the
#'    \code{\link{loadResponses}} function, with in the first column the
#'    student registration numbers of students taking the exam. The second column
#'    specifies the exam version the student filled. Subsequent columns give the
#'    filled in answers to the multiple choice exam test items (questions) as
#'    filled by the students, with a maximum of 40 test items (questions).
#' @param key A data.frame object, as create with the \code{\link{loadKey}}
#'    function, with in the first column the exam version and the following
#'    columns giving the correct answers to the multiple choice exam test items
#'    (questions).
#' @return The \code{responses} data.frame object extended with one column
#'    indicating the number of multiple choice exam test items (questions)
#'    answered correctly by a student.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' responses <- numberCorrect(responses, key)}
numberCorrect <- function(responses, key) {
  ## define constants and variables
  noItems <- ncol(key) - 1
  noCorrect <- c(0)
  ## Count the number of correct responses to the multieple choice test items
  ## for each student
  for (i in seq_len(length.out = nrow(responses))) {
    noCorrect[i] <- sum(responses[i, seq(from = 3, to = 2 + noItems)] == key[responses$Version[i], seq(from = 2, to = ncol(key))])
  }
  ## Add a new variable No.correct to the data.frame responses
  responses <- dplyr::mutate(responses, No.correct = noCorrect)
  return(responses)
}

# link.names function compares the student registration numbers scanned from
# the multiple-choice bubble form (filled students during the exam) with the
# student registration numbers from the list of students that registered for
# the exam via the Student Service Centre (SSC). In case the registration num-
# ber was incorrectly filled on the exam bubble form this function provides the
# opportunity to manually correct this. From the exam registration list the
# student's name and study are matched with the student registration number.
#' Link Names of Students and their Study to Registration Numbers
#'
#' @export
#'
#' @examples
#' \dontrun{
#' responses <- linkNames(regFile, responses)}
linkNames <- function(regFile, responses) {
  validExt <- c("csv", "xls", "xlsx")
  ## Check whether a *.csv, *.xls or *.xlsx has been provided
  if (!(tolower(strsplit(x = regFile,
                        split = "\\.")[[1]][2]) %in% validExt)) {
    stop("List of students registered for the exam not read, due to an invalid exam registration file extension (should be either *.csv or *.(x)xlsx)")
  }
  ## Read the list of students that have registered for the exam
  if (tolower(strsplit(x = regFile, split = "\\.")[[1]][2]) == validExt[1]) { # a csv-file
    ## Read the *.csv file of registered students into data.frame registered
    L <- readLines(con = regFile, n = 1L)
    if (grepl(";", L)) {
      registered <- read.csv2(file = regFile)
    } else {
      registered <- read.csv(file = regFile)
    }
  } else {
    if (tolower(strsplit(x = regFile,
                         split = "\\.")[[1]][2]) %in% validExt[2:3]) { # a xls(x)-file
      ## Read the *.xls(x) workbook file into a structure named wb
      wb <- XLConnect::loadWorkbook(filename = regFile,
                                    create = FALSE)
      ## Obtain the sheet names in the *.xls(x) workbook as represented by wb
      sheets <- XLConnect::getSheets(object = wb)
      ## Read the first sheet of wb into data.frame registered
      registered <- XLConnect::readWorksheet(object = wb,
                                             sheet = sheets[1],
                                             startRow = 0,
                                             endRow = 0,
                                             startCol = 0,
                                             endCol = 0)
    }
  }
  ## Select only the "Registratienr.", "Naam" and "Opleiding"
  registered <- dplyr::select(registered, Registratienr., Naam, Opleiding)
  ## Remove all dashes from the student registrationnumbers in column "Registratienr."
  registered[["Registratienr."]] <- gsub("-", "", registered[["Registratienr."]])
  ## Change all columns of the registered data.frame to character
  for (j in seq_len(ncol(registered))) {
    registered[, j] <- as.character(registered[, j])
  }
  ## Rename the selected column names
  registered <- dplyr::rename(registered,
                              Reg.number = Registratienr.,
                              Name = Naam,
                              Study = Opleiding)
  # ## Order registered by registration number fields 7-8-9
  # ## (represents student's last name) and last name, initials
  # registered <- registered[order(substr(registered$Reg.number, start = 7L, stop = 9L), registered$Name),]
  ## Create three new character empty variables
  corrected.reg.number <- c("")
  name <- c("")
  study <- c("")
  ## Check registration numbers of all students
  for (i in seq_len(nrow(responses))) {
    ## Find a match between the scanned student registration number and the
    ## student registration number in the list of students registered for
    ## the exam. Store the position in registered students list as selected.
    selected <- which(registered$Reg.number == responses$Reg.number[i])
    if (length(selected) == 0) { # no match found
      ## Display the student registration number for which no match in the
      ## registered students list was found as well as the name of the
      ## previous student that has been matched. Correct the mistakes in
      ## the displayed student registration number manually.
      switch(tolower(Sys.info()[["sysname"]]),
             windows = {corrected.reg.number[i] <- svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                       name[i - 1],
                                                                                       ")"),
                                                                       default = responses$Reg.number[i])$res},
             linux   = {corrected.reg.number[i] <- svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                       name[i - 1],
                                                                                       ")"),
                                                                       default = responses$Reg.number[i])$res},
             darwin  = {corrected.reg.number[i] <- strsplit(x = svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                                    name[i - 1],
                                                                                                    ")"),
                                                                                    default = responses$Reg.number[i])$res,
                                                            split = "button returned:OK, text returned:")[[1]][2]})
      ## Match the corrected student registration number with a student
      ## registration number in the list of students registered for the
      ## exam.
      selected <- which(registered$Reg.number == corrected.reg.number[i])
      if (length(selected) == 0) {
        ## No match found, either there was a mistake in the corrected student
        ## registration number or the student was not registered for the exam.
        name[i] <- "unknown"
        study[i] <- "unknown"
      } else {
        if (length(selected) == 1) {
          ## After correcting the student registration number a match with
          ## a student registered for the exam was found.
          name[i] <- registered$Name[selected]
          study[i] <- registered$Study[selected]
        }
      }
    } else {
      if (length(selected) == 1) {
      ## A match of the scanned student registration number with a student
      ## registered for the exam was found.
      corrected.reg.number[i] <- registered$Reg.number[selected] # == responses$Reg.number[i]
      name[i] <- registered$Name[selected]
      study[i] <- registered$Study[selected]
      }
    }
  }
  ## Add the corrected registration number, name and study to the responses
  ## data.frame as columns named Corrected.reg.number, Name and Study
  responses <- dplyr::mutate(responses, Corrected.reg.number = corrected.reg.number, Name = name, Study = study)
  return(responses)
}
