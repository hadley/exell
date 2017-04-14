# Setup

## Platform

|setting  |value                        |
|:--------|:----------------------------|
|version  |R version 3.3.2 (2016-10-31) |
|system   |x86_64, darwin13.4.0         |
|ui       |X11                          |
|language |(EN)                         |
|collate  |en_CA.UTF-8                  |
|tz       |America/Vancouver            |
|date     |2017-04-13                   |

## Packages

|package    |*  |version    |date       |source                           |
|:----------|:--|:----------|:----------|:--------------------------------|
|cellranger |   |1.1.0      |2016-07-27 |cran (@1.1.0)                    |
|covr       |   |2.2.2      |2017-01-05 |cran (@2.2.2)                    |
|Rcpp       |   |0.12.10    |2017-04-13 |Github (RcppCore/Rcpp@21d8388)   |
|readxl     |   |0.1.1.9000 |2017-04-14 |local (jennybc/readxl@NA)        |
|rprojroot  |   |1.2        |2017-01-16 |cran (@1.2)                      |
|testthat   |   |1.0.2.9000 |2017-04-13 |Github (hadley/testthat@b72a228) |
|tibble     |   |1.3.0      |2017-04-01 |cran (@1.3.0)                    |

# Check results

2 packages with problems

|package  |version | errors| warnings| notes|
|:--------|:-------|------:|--------:|-----:|
|elementR |1.3.0   |      1|        0|     0|
|heemod   |0.9.0   |      1|        0|     0|

## elementR (1.3.0)
Maintainer: Charlotte Sirot <charlott.sirot@gmail.com>  
Bug reports: https://github.com/charlottesirot/elementR/issues

1 error  | 0 warnings | 0 notes

```
checking examples ... ERROR
Running examples in ‘elementR-Ex.R’ failed
The error most likely occurred in:

> base::assign(".ptime", proc.time(), pos = "CheckExEnv")
> ### Name: elementR_project
> ### Title: Object elementR_project
> ### Aliases: elementR_project
> 
> ### ** Examples
> 
> ## create a new elementR_repStandard object based on the "filePath" 
> ## from a folder containing sample replicate
> 
> filePath <- system.file("Example_Session", package="elementR")
> 
> exampleProject <- elementR_project$new(filePath)
Error in structure(.External(.C_dotTclObjv, objv), class = "tclObj") : 
  [tcl] invalid command name "toplevel".
Calls: <Anonymous> ... tktoplevel -> tkwidget -> tcl -> .Tcl.objv -> structure
Execution halted
```

## heemod (0.9.0)
Maintainer: Antoine Filipovic-Pierucci <pierucci@gmail.com>  
Bug reports: https://github.com/pierucci/heemod/issues

1 error  | 0 warnings | 0 notes

```
checking tests ... ERROR
Running the tests in ‘tests/testthat.R’ failed.
Last 13 lines of output:
  Regional mortality rates cannot be estimated from local data.
  1: get_who_mr(age = 0:99, region = "GLOBAL", year = 2015) at testthat/test_gho.R:35
  2: withVisible(`_f`(age = age, sex = sex, region = region, country = country, year = year, 
         local = local))
  3: `_f`(age = age, sex = sex, region = region, country = country, year = year, local = local)
  4: stop("Regional mortality rates cannot be estimated from local data.")
  
  testthat results ================================================================
  OK: 378 SKIPPED: 0 FAILED: 1
  1. Error: GHO API (@test_gho.R#35) 
  
  Error: testthat unit tests failed
  Execution halted
```

