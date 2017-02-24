#ifndef READXL_COLSPEC_
#define READXL_COLSPEC_

#include <Rcpp.h>
#include <libxls/xls.h>
#include "StringSet.h"

enum CellType {
  CELL_BLANK,
  CELL_DATE,
  CELL_NUMERIC,
  CELL_TEXT
};

// COL_BLANK is a column full of CELL_BLANKs,
// while COL_SKIP is a user-specified column type
enum ColType {
  COL_BLANK,
  COL_DATE,
  COL_NUMERIC,
  COL_TEXT,
  COL_LIST,
  COL_SKIP
};

// ColType enum is ordered such that a simple cast
// is sufficient to convert from CellType (user-specified
// columns come after cell types)
ColType inline as_ColType(CellType cell) {
  return (ColType) cell;
}

inline std::vector<ColType> colTypeStrings(Rcpp::CharacterVector x) {
  std::vector<ColType> types;
  types.reserve(x.size());

  for (size_t i = 0; i < x.size(); ++i) {
    std::string type(x[i]);
    if (type == "blank") {
      types.push_back(COL_BLANK);
    } else if (type == "date") {
      types.push_back(COL_DATE);
    } else if (type == "numeric") {
      types.push_back(COL_NUMERIC);
    } else if (type == "text") {
      types.push_back(COL_TEXT);
    } else if (type == "list") {
      types.push_back(COL_LIST);
    } else if (type == "skip") {
      types.push_back(COL_SKIP);
    } else {
      Rcpp::stop("Unknown type '%s' at position %i", type, i + 1);
    }
  }

  return types;
}

inline std::string colTypeDesc(ColType type) {
  switch(type) {
  case COL_BLANK:   return "blank";
  case COL_DATE:    return "date";
  case COL_NUMERIC: return "numeric";
  case COL_TEXT:    return "text";
  case COL_LIST:    return "list";
  case COL_SKIP:    return "skip";
  }
  return "???";
}

bool inline isDateTime(int id, const std::set<int> custom) {
  // Date formats:
  // ECMA-376 (http://www.ecma-international.org/publications/standards/Ecma-376.htm)
  // 18.8.30 numFmt (Number Format)  (p1777)
  // Date times: 14-22, 27-36, 45-47, 50-58, 71-81 (inclusive)
  if ((id >= 14 && id <= 22) ||
      (id >= 27 && id <= 36) ||
      (id >= 45 && id <= 47) ||
      (id >= 50 && id <= 58) ||
      (id >= 71 && id <= 81))
    return true;

  // Built-in format that's not a date
  if (id < 164)
    return false;

  return custom.count(id) > 0;
}

inline Rcpp::CharacterVector colTypeDescs(std::vector<ColType> types) {
  Rcpp::CharacterVector out(types.size());
  for (size_t i = 0; i < types.size(); ++i) {
    out[i] = colTypeDesc(types[i]);
  }
  return out;
}

inline bool isDateFormat(std::string x) {
  for (size_t i = 0; i < x.size(); ++i) {
    switch (x[i]) {
    case 'd':
    case 'D':
    case 'm': // 'mm' for minutes
    case 'M':
    case 'y':
    case 'Y':
    case 'h': // 'hh'
    case 'H':
    case 's': // 'ss'
    case 'S':
      return true;
    default:
      break;
    }
  }

  return false;
}

inline std::vector<ColType> recycleTypes(std::vector<ColType> types,
                                         int ncol) {
  if (types.size() == 1) {
    types.resize(ncol);
    std::fill(types.begin(), types.end(), types[0]);
  }
  return types;
}

inline std::vector<ColType> finalizeTypes(std::vector<ColType> types) {
  // convert blank columns to a default type (numeric today, but logical soon)
  // can only happen when
  //   * col_types = NULL and we've learned them from data
  //   * all cells in column are empty or match one of the na strings
  for (size_t i = 0; i < types.size(); i++) {
    if (types[i] == COL_BLANK) {
      types[i] = COL_NUMERIC;
    }
  }
  return types;
}

inline Rcpp::CharacterVector reconcileNames(Rcpp::CharacterVector names,
                                            const std::vector<ColType>& types,
                                            int sheet_i) {
  size_t ncol_names = names.size();
  size_t ncol_types = types.size();

  if (ncol_names == ncol_types) {
    return names;
  }

  size_t ncol_noskip = 0;
  for (size_t i = 0; i < types.size(); i++) {
    if (types[i] != COL_SKIP) {
      ncol_noskip++;
    }
  }
  if (ncol_names != ncol_noskip) {
    Rcpp::stop("Sheet %d has %d columns (%d unskipped), but `col_names` has length %d.",
               sheet_i + 1, ncol_types, ncol_noskip, ncol_names);
  }

  Rcpp::CharacterVector newNames(ncol_types, "");
  size_t j_short = 0;
  for (size_t j_long = 0; j_long < ncol_types; ++j_long) {
    if (types[j_long] == COL_SKIP) {
      continue;
    }
    newNames[j_long] = names[j_short];
    j_short++;
  }
  return newNames;
}

inline Rcpp::RObject makeCol(ColType type, int n) {
  switch(type) {
  case COL_BLANK:
  case COL_SKIP:
    return R_NilValue;
    break;
  case COL_DATE: {
    Rcpp::RObject col = Rcpp::NumericVector(n, NA_REAL);
    col.attr("class") = Rcpp::CharacterVector::create("POSIXct", "POSIXt");
    col.attr("tzone") = "UTC";
    return col;
  }
    break;
  case COL_NUMERIC:
    return Rcpp::NumericVector(n, NA_REAL);
    break;
  case COL_TEXT:
    return Rcpp::CharacterVector(n, NA_STRING);
    break;
  case COL_LIST:
    return Rcpp::List(n, Rcpp::LogicalVector(1, NA_LOGICAL));
    break;
  }

  return R_NilValue;
}

inline Rcpp::List removeSkippedColumns(Rcpp::List cols,
                                       Rcpp::CharacterVector names,
                                       std::vector<ColType> types) {
  int p = cols.size();

  int p_out = 0;
  for (int j = 0; j < p; ++j) {
    if (types[j] != COL_SKIP)
      p_out++;
  }

  Rcpp::List out(p_out);
  Rcpp::CharacterVector names_out(p_out);
  int j_out = 0;
  for (int j = 0; j < p; ++j) {
    if (types[j] == COL_SKIP) {
      continue;
    }

    out[j_out] = cols[j];
    names_out[j_out] = names[j];
    j_out++;
  }

  // Turn list into a data frame
  out.attr("names") = names_out;

  return out;
}

#endif
