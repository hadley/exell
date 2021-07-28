// Generated by cpp11: do not edit by hand
// clang-format off

#include <cpp11/R.hpp>
#include <Rcpp.h>
using namespace Rcpp;
#include "cpp11/declarations.hpp"

// XlsWorkBook.cpp
CharacterVector xls_sheets(std::string path);
extern "C" SEXP _readxl_xls_sheets(SEXP path) {
  BEGIN_CPP11
    return cpp11::as_sexp(xls_sheets(cpp11::as_cpp<cpp11::decay_t<std::string>>(path)));
  END_CPP11
}
// XlsWorkBook.cpp
std::set<int> xls_date_formats(std::string path);
extern "C" SEXP _readxl_xls_date_formats(SEXP path) {
  BEGIN_CPP11
    return cpp11::as_sexp(xls_date_formats(cpp11::as_cpp<cpp11::decay_t<std::string>>(path)));
  END_CPP11
}
// XlsWorkSheet.cpp
List read_xls_(std::string path, int sheet_i, IntegerVector limits, bool shim, RObject col_names, RObject col_types, std::vector<std::string> na, bool trim_ws, int guess_max, bool progress);
extern "C" SEXP _readxl_read_xls_(SEXP path, SEXP sheet_i, SEXP limits, SEXP shim, SEXP col_names, SEXP col_types, SEXP na, SEXP trim_ws, SEXP guess_max, SEXP progress) {
  BEGIN_CPP11
    return cpp11::as_sexp(read_xls_(cpp11::as_cpp<cpp11::decay_t<std::string>>(path), cpp11::as_cpp<cpp11::decay_t<int>>(sheet_i), cpp11::as_cpp<cpp11::decay_t<IntegerVector>>(limits), cpp11::as_cpp<cpp11::decay_t<bool>>(shim), cpp11::as_cpp<cpp11::decay_t<RObject>>(col_names), cpp11::as_cpp<cpp11::decay_t<RObject>>(col_types), cpp11::as_cpp<cpp11::decay_t<std::vector<std::string>>>(na), cpp11::as_cpp<cpp11::decay_t<bool>>(trim_ws), cpp11::as_cpp<cpp11::decay_t<int>>(guess_max), cpp11::as_cpp<cpp11::decay_t<bool>>(progress)));
  END_CPP11
}
// XlsxWorkBook.cpp
CharacterVector xlsx_sheets(std::string path);
extern "C" SEXP _readxl_xlsx_sheets(SEXP path) {
  BEGIN_CPP11
    return cpp11::as_sexp(xlsx_sheets(cpp11::as_cpp<cpp11::decay_t<std::string>>(path)));
  END_CPP11
}
// XlsxWorkBook.cpp
std::vector<std::string> xlsx_strings(std::string path);
extern "C" SEXP _readxl_xlsx_strings(SEXP path) {
  BEGIN_CPP11
    return cpp11::as_sexp(xlsx_strings(cpp11::as_cpp<cpp11::decay_t<std::string>>(path)));
  END_CPP11
}
// XlsxWorkBook.cpp
std::set<int> xlsx_date_formats(std::string path);
extern "C" SEXP _readxl_xlsx_date_formats(SEXP path) {
  BEGIN_CPP11
    return cpp11::as_sexp(xlsx_date_formats(cpp11::as_cpp<cpp11::decay_t<std::string>>(path)));
  END_CPP11
}
// XlsxWorkSheet.cpp
IntegerVector parse_ref(std::string ref);
extern "C" SEXP _readxl_parse_ref(SEXP ref) {
  BEGIN_CPP11
    return cpp11::as_sexp(parse_ref(cpp11::as_cpp<cpp11::decay_t<std::string>>(ref)));
  END_CPP11
}
// XlsxWorkSheet.cpp
List read_xlsx_(std::string path, int sheet_i, IntegerVector limits, bool shim, RObject col_names, RObject col_types, std::vector<std::string> na, bool trim_ws, int guess_max, bool progress);
extern "C" SEXP _readxl_read_xlsx_(SEXP path, SEXP sheet_i, SEXP limits, SEXP shim, SEXP col_names, SEXP col_types, SEXP na, SEXP trim_ws, SEXP guess_max, SEXP progress) {
  BEGIN_CPP11
    return cpp11::as_sexp(read_xlsx_(cpp11::as_cpp<cpp11::decay_t<std::string>>(path), cpp11::as_cpp<cpp11::decay_t<int>>(sheet_i), cpp11::as_cpp<cpp11::decay_t<IntegerVector>>(limits), cpp11::as_cpp<cpp11::decay_t<bool>>(shim), cpp11::as_cpp<cpp11::decay_t<RObject>>(col_names), cpp11::as_cpp<cpp11::decay_t<RObject>>(col_types), cpp11::as_cpp<cpp11::decay_t<std::vector<std::string>>>(na), cpp11::as_cpp<cpp11::decay_t<bool>>(trim_ws), cpp11::as_cpp<cpp11::decay_t<int>>(guess_max), cpp11::as_cpp<cpp11::decay_t<bool>>(progress)));
  END_CPP11
}
// zip.cpp
void zip_xml(const std::string& zip_path, const std::string& file_path);
extern "C" SEXP _readxl_zip_xml(SEXP zip_path, SEXP file_path) {
  BEGIN_CPP11
    zip_xml(cpp11::as_cpp<cpp11::decay_t<const std::string&>>(zip_path), cpp11::as_cpp<cpp11::decay_t<const std::string&>>(file_path));
    return R_NilValue;
  END_CPP11
}

extern "C" {
static const R_CallMethodDef CallEntries[] = {
    {"_readxl_parse_ref",         (DL_FUNC) &_readxl_parse_ref,          1},
    {"_readxl_read_xls_",         (DL_FUNC) &_readxl_read_xls_,         10},
    {"_readxl_read_xlsx_",        (DL_FUNC) &_readxl_read_xlsx_,        10},
    {"_readxl_xls_date_formats",  (DL_FUNC) &_readxl_xls_date_formats,   1},
    {"_readxl_xls_sheets",        (DL_FUNC) &_readxl_xls_sheets,         1},
    {"_readxl_xlsx_date_formats", (DL_FUNC) &_readxl_xlsx_date_formats,  1},
    {"_readxl_xlsx_sheets",       (DL_FUNC) &_readxl_xlsx_sheets,        1},
    {"_readxl_xlsx_strings",      (DL_FUNC) &_readxl_xlsx_strings,       1},
    {"_readxl_zip_xml",           (DL_FUNC) &_readxl_zip_xml,            2},
    {NULL, NULL, 0}
};
}

extern "C" void R_init_readxl(DllInfo* dll){
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
  R_forceSymbols(dll, TRUE);
}
