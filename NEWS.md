# readxl 0.1.1.9000

* Improved parsing of sheet geometry for xlsx. (#224, #222, #194, #178, #163, #157, #156, #144, #102, #101, #65, @jennybc).

    - Better handling of leading and embedded blank rows and explicit row skipping.
    - Worksheets that are completely empty or that contain only column names no longer error, but return a tibble with zero rows.
    - Location is inferred for cells that do not declare their location (e.g. xlsx written by JMP).

* Logic for sheet lookup in xlsx is more robust. Improves compatibility with xlsx written by tools other than Excel and/or xlsx containing chartsheets. (#233, #104, #200, #168, #116, @jimhester, @jennybc)

* Support multiple NA values, e.g., `read_excel("missing-values.xls", na = c("NA", "1"))` (#13, #56, @jmarshallnz).

* Parse dates from .xlsx files saved with LibreOffice (#134, @zeehio).

* Default column names on .xlsx files now start with X1 instead of X0 (#98, @zeehio, @krlmlr).

* Unwanted printed output (e.g., `DEFINEDNAME: 21 00 00 ...`) is suppressed when reading .xls that contains a defined range, (#82, #188, @PedramNavid).

* Don't access value of `numFmtId` attribute when it does not exist. Can occur in xlsx written by <http://epplus.codeplex.com/> (#191, #229).

* Import the `tibble` package (#175, @krlmlr).

# readxl 0.1.1

* Add support for correctly reading strings in .xlsx files containing escaped 
  unicode characters (e.g. `_x005F_`) (#51, thanks to @jmarshallnz).
