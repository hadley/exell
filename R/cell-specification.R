#' Specify cells for reading
#'
#' The `range` argument of [read_excel()] provides many ways to limit the read
#' to a specific rectangle of cells. The simplest usage is to provide an
#' Excel-like cell range, such as `range = "D12:F15"` or
#' `range ="R1C12:R6C15"`. The cell rectangle can be specified in various other ways,
#' using helper functions. You can find more examples at the
#'  [sheet geometry](https://readxl.tidyverse.org/articles/sheet-geometry.html#range)
#'  vignette. In all cases, cell range processing is handled by the
#' [cellranger][cellranger::cellranger] package, where you can find full documentation for
#' the functions used in the examples below.
#'
#' @examples
#' path <- readxl_example("geometry.xls")
#' ## Rows 1 and 2 are empty (as are rows 7 and higher)
#' ## Column 1 aka "A" is empty (as are columns 5 of "E" and higher)
#'
#' # By default, the populated data cells are "shrink-wrapped" into a
#' # minimal data frame
#' read_excel(path)
#'
#' # Specific rectangle that is subset of populated cells, possibly improper
#' read_excel(path, range = "B3:D6")
#' read_excel(path, range = "C3:D5")
#'
#' # Specific rectangle that forces inclusion of unpopulated cells
#' read_excel(path, range = "A3:D5")
#' read_excel(path, range = "A4:E5")
#' read_excel(path, range = "C5:E7")
#'
#' # Anchor a rectangle of specified size at a particular cell
#' read_excel(path, range = anchored("C4", dim = c(3, 2)), col_names = FALSE)
#'
#' # Specify only the rows
#' read_excel(path, range = cell_rows(3:5))
#' ## is equivalent to
#' read_excel(path, range = cell_rows(c(3, 5)))
#'
#' # Specify only the columns by column number or letter
#' read_excel(path, range = cell_cols("C:D"))
#' read_excel(path, range = cell_cols(2))
#'
#' # Specify exactly one row or column bound
#' read_excel(path, range = cell_rows(c(5, NA)))
#' read_excel(path, range = cell_rows(c(NA, 4)))
#' read_excel(path, range = cell_cols(c("C", NA)))
#' read_excel(path, range = cell_cols(c(NA, 2)))
#'
#' # General open rectangles
#' # upper left = C4, everything else unspecified
#' read_excel(path, range = cell_limits(c(4, 3), c(NA, NA)))
#' # upper right = D4, everything else unspecified
#' read_excel(path, range = cell_limits(c(4, NA), c(NA, 4)))
#'
#' @seealso The [cellranger][cellranger::cellranger] package has full documentation on cell
#'   specification and offers additional functions for manipulating "A1:D10"
#'   style spreadsheet ranges. Here are the most relevant:
#'   \itemize{
#'     \item [cellranger::cell_limits()]
#'     \item [cellranger::cell_rows()]
#'     \item [cellranger::cell_cols()]
#'     \item [cellranger::anchored()]
#'   }
#' @name cell-specification
NULL

#' @importFrom cellranger cell_limits
#' @name cell_limits
#' @export
#' @rdname cell-specification
NULL

#' @importFrom cellranger cell_rows
#' @name cell_rows
#' @export
#' @rdname cell-specification
NULL

#' @importFrom cellranger cell_cols
#' @name cell_cols
#' @export
#' @rdname cell-specification
NULL

#' @importFrom cellranger anchored
#' @name anchored
#' @export
#' @rdname cell-specification
NULL
