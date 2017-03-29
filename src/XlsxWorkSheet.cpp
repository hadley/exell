#include <Rcpp.h>
#include "XlsxWorkSheet.h"
#include "ColSpec.h"
#include "utils.h"
using namespace Rcpp;

// [[Rcpp::export]]
IntegerVector parse_ref(std::string ref) {
  std::pair<int,int> parsed = parseRef(ref.c_str());

  return IntegerVector::create(parsed.first, parsed.second);
}

// [[Rcpp::export]]
CharacterVector xlsx_col_types(std::string path, int sheet_i = 0,
                               CharacterVector na = CharacterVector(),
                               int skip = 0, int n_max = -1,
                               int guess_max = 1000,
                               bool has_col_names = false) {

  XlsxWorkSheet ws(path, sheet_i, skip, n_max);
  std::vector<ColType> types(ws.ncol());
  std::fill(types.begin(), types.end(), COL_UNKNOWN);
  types = ws.colTypes(types, na, guess_max, has_col_names);
  return colTypeDescs(types);
}

// [[Rcpp::export]]
CharacterVector xlsx_col_names(std::string path,
                               CharacterVector na = CharacterVector(),
                               int sheet_i = 0, int skip = 0, int n_max = 1) {
  return XlsxWorkSheet(path, sheet_i, skip, n_max).colNames(na);
}

// [[Rcpp::export]]
List read_xlsx_(std::string path, int sheet_i, RObject col_names,
                RObject col_types, std::vector<std::string> na,
                int skip = 0, int n_max = -1, int guess_max = 1000) {

  // has_col_names = TRUE iff we will read col names from sheet -------
  CharacterVector colNames;
  bool has_col_names = false;
  switch(TYPEOF(col_names)) {
  case STRSXP:
    colNames = as<CharacterVector>(col_names);
    break;
  case LGLSXP:
    has_col_names = as<bool>(col_names);
    break;
  default:
    Rcpp::stop("`col_names` must be a logical or character vector");
  }

  // Adjust n_max -----------------------------------------------------
  if (n_max >= 0 && has_col_names) {
    n_max++;
  }

  // Construct worksheet ----------------------------------------------
  XlsxWorkSheet ws(path, sheet_i, skip, n_max);

  // catches empty sheets and sheets where we skip past all data
  if (ws.nrow() == 0 && ws.ncol() == 0) {
    return Rcpp::List(0);
  }

  // Populate column names, if necessary ------------------------------
  if (TYPEOF(col_names) == LGLSXP) {
    colNames = has_col_names ? ws.colNames(na) : CharacterVector(ws.ncol(), "");
  }

  // Get column types --------------------------------------------------
  if (TYPEOF(col_types) != STRSXP) {
    Rcpp::stop("`col_types` must be a character vector");
  }
  std::vector<ColType> colTypes = colTypeStrings(as<CharacterVector>(col_types));
  colTypes = recycleTypes(colTypes, ws.ncol());
  if ((int) colTypes.size() != ws.ncol()) {
    Rcpp::stop("Sheet %d has %d columns, but `col_types` has length %d.",
               sheet_i + 1, ws.ncol(), colTypes.size());
  }
  if (requiresGuess(colTypes)) {
    colTypes = ws.colTypes(colTypes, na, guess_max, has_col_names);
  }
  colTypes = finalizeTypes(colTypes);

  // Reconcile column names and types ----------------------------------
  colNames = reconcileNames(colNames, colTypes, sheet_i);

  // Get data ----------------------------------------------------------
  return ws.readCols(colNames, colTypes, na, has_col_names);
}
