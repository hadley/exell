#' List defined named names in a workbook
#'
#' Lists all defined names in a workbook, with any associated local range references.
#' The list of defined names is especially useful when you want to avoid hard-coding
#' the range of cells to be read from a workbook. Instead, if a range of cells to be
#' read is named in the workbook, this function allows you to obtain the actual cell
#' range, which may then be used in a subsequent call to \link{read_excel} to read the
#' range.
#'
#' The function returns a \code{tibble} with three columns:
#' * name: the name defined in the workbook
#' * context: for names defined on a worksheet, the worksheet name; NA for workbook names
#' * ref: for names defined as a range in the same workbook, the associated range
#'
#' @inheritParams excel_sheets
#' @export
#' @examples
#' names <- excel_names(readxl_example("names.xlsx"))
#' ref <- names[names$name == "data"]$ref
#' read_excel(readxl_example("names.xlsx"), range = ref)
excel_names <- function(path) {
  path <- check_file(path)
  format <- check_format(path)
  path <- normalizePath(path)

  if(format == 'xls') {
    result <- xls_names(path)
  } else if(format == 'xlsx') {
    result <- xlsx_names(path)

    # replace any non-ranges or external ranges with NA
    j <- !grepl("^(('[^[']+'|[A-Za-z0-9_]+))!(\\$?[A-Z]{1,2})?(\\$?[0-9]+)?(:(\\$?[A-Z]{1,2})?(\\$?[0-9]+)?)?$", result$ref)
    if(any(j))  {
      result$ref[j] <- NA
    }
  }

  tibble::as_tibble(result)
}
