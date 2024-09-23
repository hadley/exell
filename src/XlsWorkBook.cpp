#include "XlsWorkBook.h"

[[cpp11::register]]
cpp11::list xls_names(std::string path) {
  return XlsWorkBook(path, true).definedNames();
}

[[cpp11::register]]
cpp11::strings xls_sheets(std::string path) {
  XlsWorkBook wb(path);
  return wb.sheets();
}

[[cpp11::register]]
std::set<int> xls_date_formats(std::string path) {
  return XlsWorkBook(path).dateFormats();
}
