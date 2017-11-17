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
#' @return The input \code{responses} data.frame object extended with one column
#'    indicating the number of multiple choice exam test items (questions)
#'    answered correctly by a student.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' responses <- numberCorrect(responses, key)}
numberCorrect <- function(responses, key) {
  ## Check whether numberCorrect was already used
  if ("No. correct" %in% colnames(responses)) {
    stop("Function terminated, because the numberCorrect function was used before and, therefore, a column 'No. correct' already exists in the responses object.")
  }
  ## define constants and variables
  noItems <- ncol(key) - 1
  noCorrect <- c(0)
  ## Count the number of correct responses to the multieple choice test items
  ## for each student
  for (i in seq_len(length.out = nrow(responses))) {
    noCorrect[i] <- sum(responses[i, seq(from = 3, to = 2 + noItems)] == key[responses$Version[i], seq(from = 2, to = ncol(key))])
  }
  ## Add a new variable No.correct to the data.frame responses
  responses <- dplyr::mutate(responses, "No. correct" = noCorrect)
  return(responses)
}

#' Link Names of Students and their Study to Registration Numbers.
#'
#' \code{linkNames()} compares the student registration numbers scanned from
#' the multiple choice answer sheets (filled by students during the exam) with
#' the student registration numbers from the list of students that registered
#' for the exam via the Student Service Centre (SSC). In case the registration
#' number was incorrectly filled on the answer sheet this function provides the
#' opportunity to manually correct this. From the exam registration list the
#' student's name and study are matched with the student registration number and
#' added to the responses object.
#'
#' @param regFile  A character string specifying the location and name of
#'    the csv or (x)lsx file containing the student's registration number, name
#'    and study for all students, which registered for the exam via the Student
#'    Service Centre (SSC).
#' @param responses A data.frame object, as created with the
#'    \code{\link{loadResponses}} function and modified by the
#'    \code{\link{numberCorrect}} function.
#' @return The input \code{responses} data.frame object extended with three columns
#'    indicating the corrected registration number (Corrected reg. number),
#'    Name, and Study of all students.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' responses <- linkNames(regFile, responses)}
linkNames <- function(regFile, responses) {
  ## Check whether numberCorrect was already used
  if (!("No. correct" %in% colnames(responses))) {
    stop("Function terminated, because the column 'No. correct' does not exists in the responses object. Please use the numberCorrect function before applying the linkNames function.")
  }
  ## Check whether linkNames was already used
  if (all(c("Corrected reg. number", "Name", "Study") %in% colnames(responses))) {
    stop("Function terminated, because the linkNames function was used before and, therefore, the columns 'Corrected reg. number', 'Name', and 'Study' already exist in the responses object.")
  }
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
  registered <- dplyr::select(registered, "Registratienr.", Naam, Opleiding)
  ## Remove all dashes from the student registrationnumbers in column "Registratienr."
  registered[["Registratienr."]] <- gsub("-", "", registered[["Registratienr."]])
  ## Change all columns of the registered data.frame to character
  for (j in seq_len(ncol(registered))) {
    registered[, j] <- as.character(registered[, j])
  }
  ## Rename the selected column names
  registered <- dplyr::rename(registered,
                              "Reg. number" = "Registratienr.",
                              Name = Naam,
                              Study = Opleiding)
  # ## Order registered by registration number fields 7-8-9
  # ## (represents student's last name) and last name, initials
  # registered <- registered[order(substr(registered$"Reg. number", start = 7L, stop = 9L), registered$Name),]
  ## Create three new character empty variables
  correctedRegNumber <- c("")
  name <- c("")
  study <- c("")
  ## Check registration numbers of all students
  for (i in seq_len(nrow(responses))) {
    ## Find a match between the scanned student registration number and the
    ## student registration number in the list of students registered for
    ## the exam. Store the position in registered students list as selected.
    selected <- which(registered$"Reg. number" == responses$"Reg. number"[i])
    if (length(selected) == 0) { # no match found
      ## Display the student registration number for which no match in the
      ## registered students list was found as well as the name of the
      ## previous student that has been matched. Correct the mistakes in
      ## the displayed student registration number manually.
      switch(tolower(Sys.info()[["sysname"]]),
             windows = {correctedRegNumber[i] <- svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                      name[i - 1],
                                                                                      ")"),
                                                                     default = responses$"Reg. number"[i])$res},
             linux   = {correctedRegNumber[i] <- svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                     name[i - 1],
                                                                                     ")"),
                                                                     default = responses$"Reg. number"[i])$res},
             darwin  = {correctedRegNumber[i] <- strsplit(x = svDialogs::dlgInput(message = paste("Registration number (previous student: ",
                                                                                                  name[i - 1],
                                                                                                  ")"),
                                                                                  default = responses$"Reg. number"[i])$res,
                                                          split = "button returned:OK, text returned:")[[1]][2]})
      ## Match the corrected student registration number with a student
      ## registration number in the list of students registered for the
      ## exam.
      selected <- which(registered$"Reg. number" == correctedRegNumber[i])
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
      correctedRegNumber[i] <- registered$"Reg. number"[selected] # == responses$"Reg. number"[i]
      name[i] <- registered$Name[selected]
      study[i] <- registered$Study[selected]
      }
    }
  }
  ## Add the corrected registration number, name and study to the responses
  ## data.frame as columns named Corrected.reg.number, Name and Study
  responses <- dplyr::mutate(responses, "Corrected reg. number" = correctedRegNumber, Name = name, Study = study)
  return(responses)
}

#' Grade the Multiple Choice Exams with Correction for Guessing
#'
#' \code{gradeExam()} calculates the exam grades achieved based on the responses
#' to the multiple choice exam test items (questions). The exam grade is
#' corrected for guessing.
#'
#' @param responses A data.frame object, as created with the
#'    \code{\link{loadResponses}} function modified by the
#'    \code{\link{numberCorrect}} and \code{\link{linkNames}} functions.
#' @param noItems An integer, not exceeding 40, specifying the number of multiple
#'    choice exam test items (questions) on the answers sheet to be considered.
#'    If a \code{key} object has been created with the \code{\link{loadKey}}
#'    function prior to using the \code{loadResponses} function the \code{noItems}
#'    function argument repuires no specification and will created automatically.
#' @param noOptions An integer, which specifies the number of options for each
#'    multiple choice exam test item (question). The default is 4 options ("A",
#'    "B", "C", and "D").
#' @return The \code{responses} data.frame object extended with a column giving
#'    the calculated exam grade (Exam grade) for each student.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' responses <- gradeExam(responses, noItems, noOptions)}
gradeExam <- function(responses, noItems = length(grep("^I", colnames(responses))), noOptions = 4) {
  ## Check whether gradeExam was already used
  if ("Exam grade" %in% colnames(responses)) {
    stop("Function terminated, because the gradeExam function was used before and, therefore, a column 'Exam grade' already exists in the responses object.")
  }
  ## Check whether numberCorrect was already used
  if (!("No. correct" %in% colnames(responses))) {
    stop("Function terminated, because the column 'No. correct' does not exists in the responses object. Please use the numberCorrect function before applying the linkNames function.")
  }
  ## Check whether linkNames was already used
  if (!all(c("Corrected reg. number", "Name", "Study") %in% colnames(responses))) {
    stop("Function terminated, because the columns 'Corrected reg. number', 'Name', and 'Study' do not exist in the responses object. Please use the linkNames function before applying the gradeExam function.")
  }
  ## Save number of correct responses for convenience
  noCorrect <- responses$"No. correct"
  ## Check whether the number.correct function has been applied prior to
  ## calculating the exam grade i.e. that noCorrect is not NULL
  if(is.null(noCorrect)) {
    stop("The number of correct responses needs to be determined prior to grading the exam. Please apply the numberCorrect function first.")
  }
  ## Calculate the exam grade with correction for guessing
  examGrade <- ifelse(test = (noCorrect > round(noItems / noOptions)),
                      yes = (1 + 9 * (noCorrect - round(noItems / noOptions)) / (noItems - round(noItems / noOptions))),
                      no = 1) # 1 when noCorrect <= round(noItems / noOptions)
  ## Add the exam grade to the responses data.frame as a column named "Grade exam"
  responses <- dplyr::mutate(responses, "Exam grade" = examGrade)
  return(responses)
}
