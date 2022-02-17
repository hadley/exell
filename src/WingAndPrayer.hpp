#include "XlsWorkBook.h"
#include "XlsxWorkBook.h"
#include "XlsSheetData.h"
#include "XlsxSheetData.h"

#include "cpp11/R.hpp"

#include <string>
#include <vector>

//class XlsCell {};
//class XlsxCell {};

class Xls {
public:
 typedef XlsWorkBook  Book;
 typedef XlsSheetData SheetData;
 //typedef XlsCell      Cell;
};

class Xlsx {
public:
  typedef XlsxWorkBook  Book;
  typedef XlsxSheetData SheetData;
  //typedef XlsxCell      Cell;
};

template <typename T>
class Sheet {

public:
  Sheet(const std::string& path,
        int sheet_i, cpp11::integers limits, bool shim, bool progress)
    : wb_(path),
      sd_(wb_, sheet_i, limits, shim, progress)
  {
    Rprintf("Sheet() constructor\n");
    Rprintf("Reading from: %s\n", wb_.path().c_str());
    Rprintf(
      "Reading %d rows x %d cols from worksheet '%s'\n",
      sd_.nrow(), sd_.ncol(), sd_.sheetName().c_str());
  }

 void commonMethod() {
  Rprintf("I'm the common Sheet method!\n");
  Rprintf("Path to this Sheet's Book: %s\n", wb_.path().c_str());
 }

private:
 typename T::Book wb_;
 typename T::SheetData sd_;
 //std::vector<typename T::Cell> cells_;
};

template <>
void Sheet<Xlsx>::commonMethod() {
  Rprintf("I'm doing something special for xlsx Sheets!\n");
  Rprintf("Path to xlsx: %s\n", wb_.path().c_str());
}