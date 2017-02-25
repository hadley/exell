# readxl 0.1.1.9000

* xls files written by some third party software report both row and column dimensions as 0-indexed, which prevents libxls from reading the last column. Small change to libxls restores access to those cells. (#273, #180, #152, #99 @jennybc)

* `col_types = "list"` loads data as a list of length-1 vectors, that are typed using the logic from `col_types = NULL`, but on a cell-by-cell basis (#262 @gergness).

* A user-specified `col_types` of length one will be replicated to have length equal to the number of columns. (#127, #114, #261 @jennybc)

* Column type `"blank"` has been deprecated in favor of the more descriptive `"skip"`, which also supports the goal to become more consistent with readr. (#260, #193, #261 @jennybc)

* User-supplied `col_names` are processed relative to user-supplied `col_types`, if given. Specifically, `col_names` is considered valid if it has the same length as `col_types`, before *or after* removing skipped columns. (#81, #261 @jennybc)

* Leading or embedded empty columns are no longer dropped, regardless of whether there is a column name. (#157, #261 @jennybc)

* New argument `guess_max` lets user adjust the number of rows used to guess column types, similar to functions in readr. (#223, #257 @tklebel, @jennybc)

* Improved handling of empty cells. (xlsx #248, xls #271 @jennybc)

    - Cells with no content are not loaded. Sheet extent is always computed from loaded cells, instead of the nominal dimensions reported in the worksheet, which count cells with data *or having an explicit format*. Empty, formatted cells are detectable in Excel as seemingly empty cells with a format other than "General".
    - Eliminates a source of trailing rows (#203) and columns (#236, #162, #146) consisting entirely of `NA`.
    - Eliminates a subtle source of disagreement between user-provided column names and guessed column types (#169, #81). 

* `tibble::repair_names()` is used to prevent empty, `NA`, or duplicated names. (#216, #208, #199 #182, #53, #247 @jennybc)

* Fix compilation warning/failure (FreeBSD 10.3 #221, gcc 4.9.3 #124) and/or problems reading xls (CentOS 6.6 #189). (#244, #245, #246 @jeroenooms)

* Improved parsing of sheet geometry (xlsx #240, xls #271 @jennybc)

    - Location is inferred for cells that do not declare their location (xlsx only! e.g. xlsx written by JMP). (#163, #102)
    - Worksheets that are completely empty or that contain only column names no longer error, but return a tibble with zero rows. (#222, #144, #65)
    - Better handling of leading and embedded blank rows and explicit row skipping. (#224, #194, #178, #156, #101)

* `read_xls()` and `read_xlsx()` are now exposed, such that files without an `.xls` or `.xlsx` extension can be read. (#85, @jirkalewandowski)

* Logic for sheet lookup in xlsx is more robust. Improves compatibility with xlsx written by tools other than Excel and/or xlsx containing chartsheets. (#233, #104, #200, #168, #116, @jimhester, @jennybc)

* Support multiple NA values, e.g., `read_excel("missing-values.xls", na = c("NA", "1"))`. (#13, #56, @jmarshallnz)

* Parse dates from .xlsx files saved with LibreOffice. (#134, @zeehio)

* Default column names on .xlsx files now start with X1 instead of X0. (#98, @zeehio, @krlmlr)

* Unwanted printed output (e.g., `DEFINEDNAME: 21 00 00 ...`) is suppressed when reading .xls that contains a defined range. (#82, #188, @PedramNavid)

* Don't access value of `numFmtId` attribute when it does not exist. Can occur in xlsx written by <http://epplus.codeplex.com/>. (#191, #229)

* Import the `tibble` package (#175, @krlmlr).

# readxl 0.1.1

* Add support for correctly reading strings in .xlsx files containing escaped 
  unicode characters (e.g. `_x005F_`). (#51, thanks to @jmarshallnz)
