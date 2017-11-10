
<!-- README.md is generated from README.Rmd. Please edit that file -->
[![Build status](https://ci.appveyor.com/api/projects/status/pksd2fuf6w28cpsu?svg=true)](https://ci.appveyor.com/project/mverouden/wurmc)

The official repository for the wurmc R package
===============================================

This package is still under development and code will be subject to changes
---------------------------------------------------------------------------

<!-- The implementation sometimes changes minor details.-->
R package for processing and grading multiple choice exams at Wageningen University and Research.

Multiple choice answers sheets from students at Wageningen University and Research are generally scanned and processed into a Microsoft Excel worbook, containing a worksheet named with the course code, by EDUsupport (part of ITsupport). The used multiple choice answer sheets, however, also allow for scanning and processing with the free, open source OMR (optical mark recognition) software FormScanner (<http://www.formscanner.org/>). This package can convert the resulting FormScanner csv-file (in reality a semicolon separated file) into a Microsoft Excel workbook with the same content layout as generated by EDUsupport.

Next the Microsoft Excel workbook is used as one of the required inputs for grading the student multiple choice answer sheets. Other required files are: \* a file containing the answer key for the various versions (at most 4: A, B, C, D) used, \* a file containing the students that registered for the exam, \* a file containing the practicum information about the students, whether they passed the practicum and whether or not they have earned a bonus to be added to the exam grade.

Installation
============

It can be installed using:

``` r
install.packages("devtools")
devtools::install_github("mverouden/wurmc")
```
