#' @useDynLib readxl, .registration = TRUE
NULL

#' Read xls and xlsx files
#'
#' @param path Path to the xls/xlsx file.
#' @param sheet Sheet to read. Either a string (the name of a sheet), or an
#'   integer (the position of the sheet). Ignored if the sheet is specified via
#'   `range`. If neither argument specifies the sheet, defaults to the first
#'   sheet.
#' @param range A cell range to read from, as described in [cell-specification].
#'   Includes typical Excel ranges like "B3:D87", possibly including the sheet
#'   name like "Budget!B2:G14", and more. Interpreted strictly, even if the
#'   range forces the inclusion of leading or trailing empty rows or columns.
#'   Takes precedence over `skip`, `n_max` and `sheet`.
#' @param col_names `TRUE` to use the first row as column names, `FALSE` to get
#'   default names, or a character vector giving a name for each column. If user
#'   provides `col_types` as a vector, `col_names` can have one entry per
#'   column, i.e. have the same length as `col_types`, or one entry per
#'   unskipped column.
#' @param col_types Either `NULL` to guess all from the spreadsheet or a
#'   character vector containing one entry per column from these options:
#'   "skip", "guess", "logical", "numeric", "date", "text" or "list". If exactly
#'   one `col_type` is specified, it will be recycled. The content of a cell in
#'   a skipped column is never read and that column will not appear in the data
#'   frame output. A list cell loads a column as a list of length 1 vectors,
#'   which are typed using the type guessing logic from `col_types = NULL`, but
#'   on a cell-by-cell basis.
#' @param na Character vector of strings to interpret as missing values. By
#'   default, readxl treats blank cells as missing data.
#' @param trim_ws Should leading and trailing whitespace be trimmed?
#' @param skip Minimum number of rows to skip before reading anything, be it
#'   column names or data. Leading empty rows are automatically skipped, so this
#'   is a lower bound. Ignored if `range` is given.
#' @param n_max Maximum number of data rows to read. Trailing empty rows are
#'   automatically skipped, so this is an upper bound on the number of rows in
#'   the returned tibble. Ignored if `range` is given.
#' @param guess_max Maximum number of data rows to use for guessing column
#'   types.
#' @param progress Display a progress spinner? By default, the spinner appears
#'   only in an interactive session, outside the context of knitting a document,
#'   and when the call is likely to run for several seconds or more. See
#'   [readxl_progress()] for more details.
#' @param .name_repair Handling of column names. Passed along to
#'   [tibble::as_tibble()]. readxl's default is `.name_repair = "unique", which
#'   ensures column names are not empty and are unique.
#' @return A [tibble][tibble::tibble-package]
#' @seealso [cell-specification] for more details on targetting cells with the
#'   `range` argument
#' @export
#' @examples
#' datasets <- readxl_example("datasets.xlsx")
#' read_excel(datasets)
#'
#' # Specify sheet either by position or by name
#' read_excel(datasets, 2)
#' read_excel(datasets, "mtcars")
#'
#' # Skip rows and use default column names
#' read_excel(datasets, skip = 10, col_names = FALSE)
#'
#' # Recycle a single column type
#' read_excel(datasets, col_types = "text")
#'
#' # Specify some col_types and guess others
#' read_excel(
#'   readxl_example("deaths.xlsx"),
#'   skip = 4, n_max = 10, col_names = TRUE,
#'   col_types = c("text", "text", "guess", "guess", "guess", "guess")
#' )
#'
#' # Accomodate a column with disparate types via col_type = "list"
#' df <- read_excel(readxl_example("clippy.xlsx"), col_types = c("text", "list"))
#' df
#' df$value
#' sapply(df$value, class)
#'
#' # Limit the number of data rows read
#' read_excel(datasets, n_max = 3)
#'
#' # Read from an Excel range using A1 or R1C1 notation
#' read_excel(datasets, range = "C1:E7")
#' read_excel(datasets, range = "R1C2:R2C5")
#'
#' # Specify the sheet as part of the range
#' read_excel(datasets, range = "mtcars!B1:D5")
#'
#' # Read only specific rows or columns
#' read_excel(datasets, range = cell_rows(102:151), col_names = FALSE)
#' read_excel(datasets, range = cell_cols("B:D"))
#'
#' # Get a preview of column names
#' names(read_excel(readxl_example("datasets.xlsx"), n_max = 0))
#'
#' # exploit full .name_repair flexibility from tibble
#'
#' # "universal" names are unique and syntactic
#' read_excel(
#'   readxl_example("deaths.xlsx"),
#'   range = "arts!A5:F15",
#'   .name_repair = "universal"
#' )
#'
#' # specify name repair as a built-in function
#' read_excel(readxl_example("clippy.xlsx"), .name_repair = toupper)
#'
#' # specify name repair as a custom function
#' my_custom_name_repair <- function(nms) tolower(gsub("[.]", "_", nms))
#' read_excel(
#'   readxl_example("datasets.xlsx"),
#'   .name_repair = my_custom_name_repair
#' )
#'
#' # specify name repair as an anonymous function
#' read_excel(
#'   readxl_example("datasets.xlsx"),
#'   sheet = "chickwts",
#'   .name_repair = ~ substr(.x, start = 1, stop = 3)
#' )
read_excel <- function(path, sheet = NULL, range = NULL,
                       col_names = TRUE, col_types = NULL,
                       na = "", trim_ws = TRUE, skip = 0, n_max = Inf,
                       guess_max = min(1000, n_max),
                       progress = readxl_progress(),
                       .name_repair = "unique") {
  path <- check_file(path)
  format <- check_format(path)
  read_excel_(
    path = path, sheet = sheet, range = range,
    col_names = col_names, col_types = col_types,
    na = na, trim_ws = trim_ws, skip = skip,
    n_max = n_max, guess_max = guess_max,
    progress = progress,
    .name_repair = .name_repair,
    format = format
  )
}

#' `read_excel()` calls [excel_format()] to determine if `path` is xls or xlsx,
#' based on the file extension and the file itself, in that order. Use
#' `read_xls()` and `read_xlsx()` directly if you know better and want to
#' prevent such guessing.
#' @rdname read_excel
#' @export
read_xls <- function(path, sheet = NULL, range = NULL,
                     col_names = TRUE, col_types = NULL,
                     na = "", trim_ws = TRUE, skip = 0, n_max = Inf,
                     guess_max = min(1000, n_max),
                     progress = readxl_progress(),
                     .name_repair = "unique") {
  path <- check_file(path)
  read_excel_(
    path = path, sheet = sheet, range = range,
    col_names = col_names, col_types = col_types,
    na = na, trim_ws = trim_ws, skip = skip,
    n_max = n_max, guess_max = guess_max,
    progress = progress,
    .name_repair = .name_repair,
    format = "xls"
  )
}

#' @rdname read_excel
#' @export
read_xlsx <- function(path, sheet = NULL, range = NULL,
                      col_names = TRUE, col_types = NULL,
                      na = "", trim_ws = TRUE, skip = 0, n_max = Inf,
                      guess_max = min(1000, n_max),
                      progress = readxl_progress(),
                      .name_repair = "unique") {
  path <- check_file(path)
  read_excel_(
    path = path, sheet = sheet, range = range,
    col_names = col_names, col_types = col_types,
    na = na, trim_ws = trim_ws, skip = skip,
    n_max = n_max, guess_max = guess_max,
    progress = progress,
    .name_repair = .name_repair,
    format = "xlsx"
  )
}

read_excel_ <- function(path, sheet = NULL, range = NULL,
                        col_names = TRUE, col_types = NULL,
                        na = "", trim_ws = TRUE, skip = 0, n_max = Inf,
                        guess_max = min(1000, n_max),
                        progress = readxl_progress(),
                        .name_repair = NULL,
                        format) {
  if (format == "xls") {
    sheets_fun <- xls_sheets
    read_fun <- read_xls_
  } else {
    sheets_fun <- xlsx_sheets
    read_fun <- read_xlsx_
  }
  path <- normalizePath(path)
  sheet <- standardise_sheet(sheet, range, sheets_fun(path))
  shim <- !is.null(range)
  limits <- standardise_limits(
    range, skip, n_max, has_col_names = isTRUE(col_names)
  )
  col_types <- check_col_types(col_types)
  guess_max <- check_guess_max(guess_max)
  trim_ws <- check_bool(trim_ws, "trim_ws")
  progress <- check_bool(progress, "progress")
  set_readxl_names(
    read_fun(
      path = path, sheet_i = sheet,
      limits = limits, shim = shim,
      col_names = col_names, col_types = col_types,
      na = na, trim_ws = trim_ws, guess_max = guess_max,
      progress = progress
    ),
    .name_repair = .name_repair
  )
}

# Helper functions -------------------------------------------------------------

## return a zero-indexed sheet number
standardise_sheet <- function(sheet, range, sheet_names) {
  range_sheet <- cellranger::as.cell_limits(range)[["sheet"]]
  if (!is.null(range_sheet) && !is.na(range_sheet)) {
    if (!is.null(sheet)) {
      message(
        "Two values given for `sheet`. ",
        "Using the `sheet` found in `range`:\n", range_sheet
      )
    }
    sheet <- range_sheet
  }

  if (is.null(sheet)) {
    sheet <- 1L
  }

  if (length(sheet) != 1) {
    stop("`sheet` must have length 1", call. = FALSE)
  }

  if (is.numeric(sheet)) {
    if (sheet < 1) {
      if(is.infinite(sheet)) sheet <- -length(sheet_names)
      length(sheet_names) + sheet
    } else {
      if(is.infinite(sheet)) sheet <- length(sheet_names)
      floor(sheet) - 1L
    }
  } else if (is.character(sheet)) {
    if (!(sheet %in% sheet_names)) {
      stop("Sheet '", sheet, "' not found", call. = FALSE)
    }
    match(sheet, sheet_names) - 1L
  } else {
    stop("`sheet` must be either an integer or a string.", call. = FALSE)
  }
}

## return a zero-indexed vector describing the corners of a cell rectangle:
## min_row, max_row, min_col, max_col
## NA becomes -1 and means "unspecified", by convention
## if both min and max are -1, for rows or cols, means "read them all"
## min_row = -2 is a special flag meaning "read no rows"
standardise_limits <- function(range, skip, n_max, has_col_names) {
  if (is.null(range)) {
    skip <- check_non_negative_integer(skip, "skip")
    n_max <- check_non_negative_integer(n_max, "n_max")
    n_read <- if (has_col_names) n_max + 1 else n_max
    limits <- c(
      min_row = if (n_read > 0) skip else -2,
      max_row = if (n_read == Inf || n_read == 0) NA else skip + n_read - 1,
      min_col = NA,
      max_col = NA
    )
  } else {
    limits <- cellranger::as.cell_limits(range)
    limits <- c(
      min_row = limits[["ul"]][1] - 1,
      max_row = limits[["lr"]][1] - 1,
      min_col = limits[["ul"]][2] - 1,
      max_col = limits[["lr"]][2] - 1
    )
  }
  limits[is.na(limits)] <- -1
  names <- names(limits)
  limits <- as.integer(limits)
  names(limits) <- names
  limits
}

check_col_types <- function(col_types) {
  if (is.null(col_types)) {
    return("guess")
  }
  stopifnot(is.character(col_types), length(col_types) > 0, !anyNA(col_types))

  blank <- col_types == "blank"
  if (any(blank)) {
    message("`col_type = \"blank\"` deprecated. Use \"skip\" instead.")
    col_types[blank] <- "skip"
  }

  accepted_types <-
    c("skip", "guess", "logical", "numeric", "date", "text", "list")
  ok <- col_types %in% accepted_types
  if (any(!ok)) {
    info <- paste(
      paste0("'", col_types[!ok], "' [", seq_along(col_types)[!ok], "]"),
      collapse = ", "
    )
    stop(paste("Illegal column type:", info), call. = FALSE)
  }
  col_types
}

check_bool <- function(bool, arg_name) {
  if (!isTRUE(bool) && !identical(bool, FALSE)) {
    stop("`", arg_name, "` must be either TRUE or FALSE", call. = FALSE)
  }
  bool
}

check_non_negative_integer <- function(i, arg_name) {
  if (length(i) != 1 || !is.numeric(i) || !is_integerish(i) ||
    is.na(i) || i < 0) {
    stop("`", arg_name, "` must be a positive integer", call. = FALSE)
  }
  i
}

## from readr
check_guess_max <- function(guess_max, max_limit = .Machine$integer.max %/% 100) {
  guess_max <- check_non_negative_integer(guess_max, "guess_max")
  if (guess_max > max_limit) {
    warning(
      "`guess_max` is a very large value, setting to `", max_limit,
      "` to avoid exhausting memory", call. = FALSE
    )
    guess_max <- max_limit
  }
  guess_max
}

set_readxl_names <- function(l, .name_repair = "unique") {
  if (is.null(.name_repair)) {
    tibble::as_tibble(l)
  } else {
    tibble::as_tibble(l, .name_repair = .name_repair)
  }
}
