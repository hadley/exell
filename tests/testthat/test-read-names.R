test_that("excel_names can read names from xlsx workbook", {
  x <- excel_names(test_sheet("named-ranges.xlsx"))
  expect_s3_class(x, 'data.frame')
  expect_equal(ncol(x), 3)
  expect_equal(nrow(x), 16)
  expect_setequal(unique(x$context), c(NA, 'Sheet1', 'Sheet 2'))
  expect_length(na.omit(x$context), 5)
  expect_length(na.omit(x$ref), 14)
})

test_that("excel_names can read names from xls workbook", {
  x <- excel_names(test_sheet("named-ranges.xls"))
  expect_s3_class(x, 'data.frame')
  expect_equal(ncol(x), 3)
  expect_equal(nrow(x), 16)
  expect_setequal(unique(x$context), c(NA, 'Sheet1', 'Sheet 2'))
  expect_length(na.omit(x$context), 5)
  expect_length(na.omit(x$ref), 14)
})

test_that("excel_names succeeds if no names in xlsx workbook", {
  x <- excel_names(test_sheet("types.xlsx"))
  expect_s3_class(x, 'data.frame')
  expect_equal(ncol(x), 3)
  expect_equal(nrow(x), 0)
})

test_that("excel_names succeeds if no names in xls workbook", {
  x <- excel_names(test_sheet("mtcars.xls"))
  expect_s3_class(x, 'data.frame')
  expect_equal(ncol(x), 3)
  expect_equal(nrow(x), 0)
})
