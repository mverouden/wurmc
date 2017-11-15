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
