#' List all sheets in an excel spreadsheet
#'
#' The list of sheet names is especially useful when you want to iterate over
#' all of the sheets in a workbook. The [readxl Workflows article](https://readxl.tidyverse.org/articles/readxl-workflows.html)
#' provides several worked examples of this, showing how to combine readxl with
#' other packages in the tidyverse, such as purrr, or with base R functions like
#' [lapply()].
#'
#' @inheritParams read_excel
#' @export
#' @examples
#' excel_sheets(readxl_example("datasets.xlsx"))
#' excel_sheets(readxl_example("datasets.xls"))
#'
#' # To load all sheets in a workbook, use lapply()
#' path <- readxl_example("datasets.xls")
#' lapply(excel_sheets(path), read_excel, path = path)
excel_sheets <- function(path) {
  path <- check_file(path)
  format <- check_format(path)
  path <- normalizePath(path)

  switch(format,
    xls = xls_sheets(path),
    xlsx = xlsx_sheets(path)
  )
}
