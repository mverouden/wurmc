% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/writeFunctions.R
\name{formScanner2Excel}
\alias{formScanner2Excel}
\title{Convert a FormScanner file to Microsoft Excel file.}
\usage{
formScanner2Excel(fileName, noItems, courseCode, saveName = "ICT_results")
}
\arguments{
\item{fileName}{A character string with the name of the FormScanner csv-file
to be converted into a Microsoft Excel xlsx-file.}

\item{noItems}{An integer, not exceeding 40, specifying the number of items
(questions) on the answers sheet to be considered.}

\item{courseCode}{A character string with the course code, to be used as the
name of the worksheet in the Microscoft Excel workbook, e.g. MAT15403.
Preferably use the course code as specified in the university study guide.}

\item{saveName}{A character string containing the name of how the Microsoft
Excel workbook should be saved. The extension is not needed, automatically
xlsx will be added as file extension.}
}
\value{
A Microsoft Excel workbook xlsx-file named with the name specified
by the \code{saveName} function argument.
}
\description{
Convert responses of a multiple choice answer sheet as obtained by Formscanner
,as a csv-file, into a Microsoft Excel workbook xlsx-file.
}
\details{
Multiple choice answers sheets from students at Wageningen University and
Research are generally scanned and processed into a Microsoft Excel worbook,
containing a worksheet named with the course code, by EDUsupport (part of
ITsupport). The used multiple choice answer sheets, however, also allow for
scanning and processing with the free, open source OMR (optical mark
recognition) software FormScanner (\url{http://www.formscanner.org/}).

This function converts the resulting FormScanner csv-file (in reality a
semicolon separated file) into a Microsoft Excel workbook with the same
content layout as generated by EDUsupport.
}
\examples{
\dontrun{
  formScanner2Excel(fileName = "examples/2015-2016_P1_MAT15403_xm151020_scan_results.csv",
                    noItems = 25,
                    courseCode = "MAT15403",
                    saveName = "2015-2016_P1_MAT15403_xm151020_ICT_results")
}
}
\seealso{
\code{\link{processMC}}
}
