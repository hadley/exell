#ifndef READXL_XLSXWORKSHEET_
#define READXL_XLSXWORKSHEET_

#include <Rcpp.h>
#include "rapidxml.h"
#include "XlsxWorkBook.h"
#include "XlsxCell.h"

// Key reference for understanding the structure of the XML is
// ECMA-376 (http://www.ecma-international.org/publications/standards/Ecma-376.htm)
// Section and page numbers below refer to the 4th edition
// 18.3.1.73  row         (Row)        [p1677]
// 18.3.1.4   c           (Cell)       [p1598]
// 18.3.1.96  v           (Cell Value) [p1709]
// 18.18.11   ST_CellType (Cell Type)  [p2443]

class XlsxWorkSheet {
  XlsxWorkBook wb_;
  std::string sheet_;
  rapidxml::xml_document<> sheetXml_;
  rapidxml::xml_node<>* rootNode_;
  rapidxml::xml_node<>* sheetData_;
  int ncol_, nrow_;

public:

  XlsxWorkSheet(XlsxWorkBook wb, int sheet_i): wb_(wb) {
    if (sheet_i > wb.n_sheets()) {
      Rcpp::stop("Can't retrieve sheet in position %d, only %d sheets found.",
                 sheet_i,  wb.n_sheets());
    }
    std::string sheetPath = wb.sheetPath(sheet_i);
    sheet_ = zip_buffer(wb.path(), sheetPath);
    sheetXml_.parse<0>(&sheet_[0]);

    rootNode_ = sheetXml_.first_node("worksheet");
    if (rootNode_ == NULL)
      Rcpp::stop("Invalid sheet xml (no <worksheet>)");

    sheetData_ = rootNode_->first_node("sheetData");
    if (sheetData_ == NULL)
      Rcpp::stop("Invalid sheet xml (no <sheetData>)");

    cacheDimension();
    duplicateMergedCells();
  }
  
  void duplicateMergedCells() {
    rapidxml::xml_node<>* mergeCells_ = rootNode_->first_node("mergeCells");
    if (mergeCells_ == NULL)
      return;
    for (rapidxml::xml_node<>* merged_cell = mergeCells_->first_node("mergeCell");
         merged_cell; merged_cell = merged_cell->next_sibling("mergeCell")) {
    
      rapidxml::xml_attribute<>* ref = merged_cell->first_attribute("ref");
      if (ref == NULL)
        Rcpp::stop("Invalid sheet xml (no ref for <mergeCell>)");
      
      const char* refv = ref->value();
      std::string ref_string = std::string(refv);
      
      if (ref_string.find(":") == std::string::npos)
        Rcpp::stop("No ':' in mergeCell ref '%s'", ref_string);
        
      const char * first_str = ref_string.substr(0, ref_string.find(':')).c_str();
      const char * second_str = ref_string.substr(ref_string.find(':')+1, ref_string.length()).c_str();
      
      std::pair<int, int> first_coord = parseRef(first_str);
      std::pair<int, int> second_coord = parseRef(second_str);
        
      rapidxml::xml_node<>* row = getRow(first_coord.first);
      rapidxml::xml_node<>* to_be_duplicated = getColumn(row, first_coord.second); // node to be duplicated in all the merged cells
     
      for (int r_i = first_coord.first; r_i <= second_coord.first && row; r_i++) {
        
        rapidxml::xml_node<>* current_node = getColumn(row, first_coord.second);  
        for (int c_i = first_coord.second; c_i <= second_coord.second; c_i++) {
            
          rapidxml::xml_attribute<>* current_node_r = current_node->first_attribute("r"); // gets the old node's name
           
          rapidxml::xml_node<>* copied_node = sheetXml_.clone_node( to_be_duplicated ); // clones node 
          copied_node->remove_attribute(copied_node->first_attribute("r")); // removes new cloned node's name
          copied_node->prepend_attribute(sheetXml_.allocate_attribute(current_node_r->name(), 
                                                                      current_node_r->value(), 
                                                                      current_node_r->name_size(), 
                                                                      current_node_r->value_size())); // adds old node's name to new node by creating new attribute
          row->insert_node(current_node, copied_node ); // inserts new node
          row->remove_node(current_node); // removes old node
          current_node = current_node->next_sibling("c");
        }          
        row = row->next_sibling("row");     
      }  
    }
    return;
  }

  int ncol() {
    return ncol_;
  }

  int nrow() {
    return nrow_;
  }

  void printCells() {
    for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
         row; row = row->next_sibling("row")) {

      for (rapidxml::xml_node<>* cell = row->first_node("c");
           cell; cell = cell->next_sibling("c")) {

        XlsxCell xcell(cell);
        Rcpp::Rcout << xcell.row() << "," << xcell.row() << ": " <<
          cellTypeDesc(xcell.type("", wb_.stringTable(), wb_.dateStyles())) << "\n";
      }
    }
  }


  std::vector<CellType> colTypes(const StringSet& na, int nskip = 0, int n_max = 100, bool has_col_names = false) {
    rapidxml::xml_node<>* row = getRow(nskip + has_col_names);
    std::vector<CellType> types;
    types.resize(ncol_);

    int i = 0;
    while(i < n_max && row != NULL) {
      for (rapidxml::xml_node<>* cell = row->first_node("c");
           cell; cell = cell->next_sibling("c")) {

        XlsxCell xcell(cell);
        if (xcell.col() >= ncol_)
          continue;

        CellType type = xcell.type(na, wb_.stringTable(), wb_.dateStyles());
        if (type >= types[xcell.col()]) {
          types[xcell.col()] = type;
        }
      }

      row = row->next_sibling("row");
      i++;
    }

    if (has_col_names) {
      // blank columns with a name aren't blank
      Rcpp::CharacterVector names = colNames(nskip);
      for (size_t i = 0; i < types.size(); i++) {
        if (types[i] == CELL_BLANK && names[i] != NA_STRING && names[i] != "")
          types[i] = CELL_NUMERIC;
      }
    }

    return types;
  }

  Rcpp::CharacterVector colNames(int nskip = 0) {
    rapidxml::xml_node<>* row = getRow(nskip);

    Rcpp::CharacterVector out(ncol_);
    for (rapidxml::xml_node<>* cell = row->first_node("c");
         cell; cell = cell->next_sibling("c")) {
      XlsxCell xcell(cell);

      if (xcell.col() >= ncol_)
        continue;
      out[xcell.col()] = xcell.asCharSxp("", wb_.stringTable());
    }

    return out;
  }

  Rcpp::List readCols(Rcpp::CharacterVector names,
                      const std::vector<CellType>& types,
                      const StringSet& na, int nskip = 0) {
    if ((int) names.size() != ncol_ || (int) types.size() != ncol_)
      Rcpp::stop("Need one name and type for each column");

    // Initialise columns
    int n = nrow_ - nskip;
    Rcpp::List cols(ncol_);
    for (int j = 0; j < ncol_; ++j) {
      cols[j] = makeCol(types[j], n);
    }
    
    int i = 0;
    for (rapidxml::xml_node<>* row = getRow(nskip);
         row; row = row->next_sibling("row")) {
      if ((i + 1) % 1000 == 0)
        Rcpp::checkUserInterrupt();

      for (rapidxml::xml_node<>* cell = row->first_node("c");
           cell; cell = cell->next_sibling("c")) {

        XlsxCell xcell(cell);
        CellType type = xcell.type(na, wb_.stringTable(), wb_.dateStyles());
        if (xcell.col() >= ncol_)
          continue;

        Rcpp::RObject col = cols[xcell.col()];
        // Needs to compare to actual cell type to give warnings
        switch(types[xcell.col()]) {
        case CELL_BLANK:
          break;
        case CELL_NUMERIC:
          switch(type) {
          case CELL_NUMERIC:
          case CELL_DATE:
            REAL(col)[i] = xcell.asDouble(na);
            break;
          case CELL_BLANK:
            REAL(col)[i] = NA_REAL;
            break;
          case CELL_TEXT:
            Rcpp::warning("[%i, %i]: expecting numeric: got '%s'",
              xcell.row() + 1, xcell.col() + 1, xcell.asStdString(wb_.stringTable()));
            REAL(col)[i] = NA_REAL;
          }
          break;
        case CELL_DATE:
          switch(type) {
          case CELL_DATE:
            REAL(col)[i] = xcell.asDate(na, wb_.offset());
            break;
          case CELL_BLANK:
            REAL(col)[i] = NA_REAL;
            break;
          case CELL_NUMERIC:
          case CELL_TEXT:
            Rcpp::warning("[%i, %i]: expecting date: got '%s'",
              xcell.row() + 1, xcell.col() + 1, xcell.asStdString(wb_.stringTable()));
            REAL(col)[i] = NA_REAL;
            break;
          }
          break;
        case CELL_TEXT:
          if (type == CELL_BLANK) {
            SET_STRING_ELT(col, i, NA_STRING);
          } else {
            SET_STRING_ELT(col, i, xcell.asCharSxp(na, wb_.stringTable()));
          }
          break;
        }
      }

      ++i;
    }

    return removeBlankColumns(cols, names, types);
  }


private:

  rapidxml::xml_node<>* getRow(int i) {
    rapidxml::xml_node<>* row = sheetData_->first_node("row");
    while(i > 0 && row != NULL) {
      row = row->next_sibling("row");
      i--;
    }
    if (row == NULL)
      Rcpp::stop("Skipped over all data");

    return row;
  }

  rapidxml::xml_node<>* getColumn(rapidxml::xml_node<>* row_node, int i) {
    rapidxml::xml_node<>* cell = row_node->first_node("c");
    if (cell == NULL)
        Rcpp::stop("Row does not have columns");
      
    rapidxml::xml_attribute<>* col_ref;
    int col;
    
    while(cell != NULL) {
      col_ref = cell->first_attribute("r");
      if (col_ref == NULL)
        Rcpp::stop("Cell doesn't have name");
      col = parseRef(col_ref->value()).second;
      if (col == i)
         break;
      cell = cell->next_sibling("c");    
    }
    if (cell == NULL)
      Rcpp::stop("No cell exists at column %d in current row. Presumably because improperly formatted .xlsx file", i);
    return cell;
  }

  void cacheDimension() {
    // 18.3.1.35 dimension (Worksheet Dimensions) [p 1627]
    rapidxml::xml_node<>* dimension = rootNode_->first_node("dimension");
    if (dimension == NULL)
      return computeDimensions();

    rapidxml::xml_attribute<>* ref = dimension->first_attribute("ref");
    if (ref == NULL)
      return computeDimensions();

    const char* refv = ref->value();
    while (*refv != ':' && *refv != '\0')
      ++refv;
    if (*refv == '\0')
      return computeDimensions();

    ++refv; // advanced past :
    std::pair<int, int> dim = parseRef(refv);
    if (dim.first == -1 || dim.second == -1)
      return computeDimensions();

    nrow_ = dim.first + 1; // size is one greater than max position
    ncol_ = dim.second + 1;
  }

  void computeDimensions() {
    // If <dimension> not present, iterate over all rows and cells to count
    nrow_ = 0;
    ncol_ = 0;

    for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {

      for (rapidxml::xml_node<>* cell = row->first_node("c");
        cell; cell = cell->next_sibling("c")) {

        XlsxCell xcell(cell);
        if (nrow_ < xcell.row())
          nrow_ = xcell.row();

        if (ncol_ < xcell.col())
          ncol_ = xcell.col();

      }
    }
    nrow_++;
    ncol_++;
  }

};

#endif
