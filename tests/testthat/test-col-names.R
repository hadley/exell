test_that("col_names = FALSE mimics missing column names [xlsx]", {
  df1 <- read_excel(
    test_sheet("no-yes-col-names.xlsx"),
    sheet = "no-col-names",
    col_names = FALSE
  )
  df2 <- read_excel(
    test_sheet("no-yes-col-names.xlsx"),
    sheet = "yes-col-names",
    skip = 1,
    col_names = c("", "")
  )
  df3 <- read_excel(
    test_sheet("no-yes-col-names.xlsx"),
    ## A1 and B1 aren't truly empty -- contain invisible single quotes, so we
    ## don't skip over them
    sheet = "empty-col-names"
  )
  expect_identical(df1, df2)
  expect_identical(df2, df3)
})

test_that("col_names = FALSE mimics missing column names [xls]", {
  df1 <- read_excel(
    test_sheet("no-yes-col-names.xls"),
    sheet = "no-col-names",
    col_names = FALSE
  )
  df2 <- read_excel(
    test_sheet("no-yes-col-names.xls"),
    sheet = "yes-col-names",
    skip = 1,
    col_names = c("", "")
  )
  df3 <- read_excel(
    test_sheet("no-yes-col-names.xls"),
    ## A1 and B1 aren't truly empty -- contain invisible single quotes, so we
    ## don't skip over them
    sheet = "empty-col-names"
  )
  expect_identical(df1, df2)
  expect_identical(df2, df3)
})

test_that("missing column names are populated", {
  nms <- c("...1", "...3")
  df <- read_excel(test_sheet("unnamed-duplicated-columns.xlsx"))
  expect_identical(names(df)[c(1, 3)], nms)
  df <- read_excel(test_sheet("unnamed-duplicated-columns.xls"))
  expect_identical(names(df)[c(1, 3)], nms)
})

test_that("column names are de-duplicated", {
  nm <- "var2...4"

  df <- read_excel(test_sheet("unnamed-duplicated-columns.xlsx"))
  expect_identical(names(df)[4], nm)
  df <- read_excel(test_sheet("unnamed-duplicated-columns.xls"))
  expect_identical(names(df)[4], nm)
})

test_that("wrong length column names are rejected", {
  expect_snapshot(
    error = TRUE,
    read_excel(test_sheet("iris-excel-xlsx.xlsx"), col_names = LETTERS[1:3])
  )
  expect_snapshot(
    error = TRUE,
    read_excel(test_sheet("iris-excel-xls.xls"), col_names = LETTERS[1:3])
  )
})

test_that("column_names can anticipate skipping", {

  ## xlsx
  df <- read_excel(
    test_sheet("iris-excel-xlsx.xlsx"),
    col_names = c("one", "two", "three"), skip = 1,
    col_types = c("numeric", "numeric", "skip", "skip", "text")
  )
  expect_identical(dim(df), c(150L, 3L))
  expect_identical(names(df), c("one", "two", "three"))

  ## xls
  df <- read_excel(
    test_sheet("iris-excel-xls.xls"),
    col_names = c("one", "two", "three"), skip = 1,
    col_types = c("numeric", "numeric", "skip", "skip", "text")
  )
  expect_identical(dim(df), c(150L, 3L))
  expect_identical(names(df), c("one", "two", "three"))
})

test_that(".name_repair is deprecated", {
  expect_colnames <- function(df, expected) {
    expect_identical(colnames(df), expected)
  }

  ## "universal" names are available
  nms <- c("a.b...1", "a.b...2", "...3", "c...")
  lifecycle::expect_deprecated(
    xlsx <- read_excel(test_sheet("names-need-repair-xlsx.xlsx"), .name_repair = "universal")
  )
  expect_colnames(xlsx, nms)
})

test_that("name_repair is passed through to tibble", {
  expect_colnames <- function(df, expected) {
    expect_identical(colnames(df), expected)
  }

  ## default is "unique"
  nms <- c("a b...1", "a b...2", "...3", "c%&$")
  xlsx <- read_excel(test_sheet("names-need-repair-xlsx.xlsx"))
  expect_colnames(xlsx, nms)
  xls <- read_excel(test_sheet("names-need-repair-xls.xls"))
  expect_colnames(xls, nms)

  ## "universal" names are available
  nms <- c("a.b...1", "a.b...2", "...3", "c...")
  xlsx <- read_excel(test_sheet("names-need-repair-xlsx.xlsx"), name_repair = "universal")
  expect_colnames(xlsx, nms)
  xls <- read_excel(test_sheet("names-need-repair-xls.xls"), name_repair = "universal")
  expect_colnames(xls, nms)

  ## "minimal" names are available
  nms <- c("", "var2", "", "var2")
  xlsx <- read_excel(test_sheet("unnamed-duplicated-columns.xlsx"), name_repair = "minimal")
  expect_colnames(xlsx, nms)
  xls <- read_excel(test_sheet("unnamed-duplicated-columns.xls"), name_repair = "minimal")
  expect_colnames(xls, nms)

  ## specify name repair as a built-in function
  nms <- c("SEPAL.LENGTH", "SEPAL.WIDTH", "PETAL.LENGTH", "PETAL.WIDTH", "SPECIES")
  xlsx <- read_excel(test_sheet("iris-excel-xlsx.xlsx"), name_repair = toupper)
  expect_colnames(xlsx, nms)
  xls <- read_excel(test_sheet("iris-excel-xls.xls"), name_repair = toupper)
  expect_colnames(xls, nms)

  ## specify name repair as a custom function
  nms <- c("sepal_length", "sepal_width", "petal_length", "petal_width", "species")
  custom_name_repair <- function(nms) tolower(gsub("[.]", "_", nms))
  xlsx <- read_excel(test_sheet("iris-excel-xlsx.xlsx"), name_repair = custom_name_repair)
  expect_colnames(xlsx, nms)
  xls <- read_excel(test_sheet("iris-excel-xls.xls"), name_repair = custom_name_repair)
  expect_colnames(xls, nms)

  ## specify name repair as an anonymous function
  nms <- c("wei", "fee")
  xlsx <- read_excel(
    readxl_example("datasets.xlsx"),
    sheet = "chickwts",
    name_repair = ~ substr(.x, start = 1, stop = 3)
  )
  expect_colnames(xlsx, nms)
  xls <- read_excel(
    readxl_example("datasets.xls"),
    sheet = "chickwts",
    name_repair = ~ substr(.x, start = 1, stop = 3)
  )
  expect_colnames(xls, nms)
})
