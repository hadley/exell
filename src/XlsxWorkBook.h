#pragma once

#include "ColSpec.h"
#include "utils.h"
#include "XlsxString.h"
#include "zip.h"

#include "rapidxml/rapidxml.h"

#include "cpp11/protect.hpp"
#include "cpp11/strings.hpp"

#include <map>

class XlsxWorkBook {

  // holds objects related to package parts, such as file paths, but also
  // worksheet position, name, and id
  class PackageRelations {
    // non-worksheet package parts
    std::map<std::string, std::string> part_;

    // worksheets
    int n_;
    cpp11::writable::strings names_;
    cpp11::writable::strings id_;
    std::map<std::string, std::string> target_;

    // defined names
    int defn_n_;
    cpp11::writable::strings defn_names_;
    cpp11::writable::strings defn_context_;
    cpp11::writable::strings defn_ref_;

    // populate workbook element of part_
    void parse_package_rels(const std::string& path) {
      std::string rels_xml_file = zip_buffer(path, "_rels/.rels");
      rapidxml::xml_document<> rels_xml;
      rels_xml.parse<rapidxml::parse_strip_xml_namespaces>(&rels_xml_file[0]);

      rapidxml::xml_node<>* relationships = rels_xml.first_node("Relationships");
      if (relationships == NULL) {
        cpp11::stop("Spreadsheet has no package-level relationships");
      }

      std::map<std::string, std::string> scratch;
      for (rapidxml::xml_node<>* relationship = relationships->first_node();
           relationship; relationship = relationship->next_sibling()) {

        rapidxml::xml_attribute<>* type = relationship->first_attribute("Type");
        rapidxml::xml_attribute<>* target = relationship->first_attribute("Target");

        if (type != NULL && target != NULL) {
          // keys are derived by abbreviating Type
          // in XML: http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
          // key is just the last part: officeDocument
          scratch[baseName(type->value())] = target->value();
        }
      }

      // identify and store the workbook part = the main part
      // ECMA-376 Part 1 section 8.5 SpreadsheetML
      std::map<std::string, std::string>::iterator m = scratch.find("officeDocument");
      if (m == scratch.end()) {
        cpp11::stop("No workbook part found");
      }
      // store 'xl/workbook.xml', not '/xl/workbook.xml' (rare, but have seen)
      part_["officeDocument"] = removeLeadingSlashes(m->second);
    }

    // populates n_, names_, id_
    void parse_workbook(const std::string& path) {
      std::string workbookXml = zip_buffer(path, part_["officeDocument"]);

      rapidxml::xml_document<> workbook;
      workbook.parse<rapidxml::parse_strip_xml_namespaces>(&workbookXml[0]);

      rapidxml::xml_node<>* root = workbook.first_node("workbook");
      if (root == NULL) {
        return;
      }

      rapidxml::xml_node<>* sheets = root->first_node("sheets");
      if (sheets == NULL) {
        return;
      }

      int i = 0;
      for (rapidxml::xml_node<>* sheet = sheets->first_node();
           sheet; sheet = sheet->next_sibling()) {
        if (i >= n_) {
          n_ *= 2;
          names_ = Rf_lengthgets(names_, n_);
          id_ = Rf_lengthgets(id_, n_);
        }
        rapidxml::xml_attribute<>* name = sheet->first_attribute("name");
        names_[i] = (name != NULL) ? Rf_mkCharCE(name->value(), CE_UTF8) : NA_STRING;

        rapidxml::xml_attribute<>* id = sheet->first_attribute("id");
        id_[i] = (id != NULL) ? Rf_mkCharCE(id->value(), CE_UTF8) : NA_STRING;

        i++;
      }

      if (i != n_) {
        names_ = Rf_lengthgets(names_, i);
        id_ = Rf_lengthgets(id_, i);
        n_ = i;
      }

      // parse defined names
      rapidxml::xml_node<>* defns = root->first_node("definedNames");
      i = 0;
      if (defns != NULL) {
        for (rapidxml::xml_node<>* defn = defns->first_node();
             defn; defn = defn->next_sibling()) {
          if (i >= defn_n_) {
            defn_n_ *= 2;
            defn_names_ = Rf_lengthgets(names_, defn_n_);
            defn_context_ = Rf_lengthgets(names_, defn_n_);
            defn_ref_ = Rf_lengthgets(names_, defn_n_);
          }

          rapidxml::xml_attribute<>* defn_name = defn->first_attribute("name");
          defn_names_[i] = (defn_name != NULL) ? Rf_mkCharCE(defn_name->value(), CE_UTF8) : NA_STRING;

          rapidxml::xml_attribute<>* defn_sheetId = defn->first_attribute("localSheetId");
          if(defn_sheetId == NULL) {
            defn_context_[i] = NA_STRING;
          } else {
            int index = std::stoi(defn_sheetId->value());
            if(index >= n_) {
              defn_context_[i] = NA_STRING;
            } else {
              cpp11::r_string rs = names_.at(index);
              defn_context_[i] = (SEXP)rs;
            }
          }

          defn_ref_[i] = Rf_mkCharCE(defn->value(), CE_UTF8);

          i++;
        }
      }

      if (i != defn_n_) {
        defn_names_ = Rf_lengthgets(defn_names_, i);
        defn_context_ = Rf_lengthgets(defn_context_, i);
        defn_ref_ = Rf_lengthgets(defn_ref_, i);
        defn_n_ = i;
      }
    }

    // populates target_
    void parse_workbook_rels(const std::string& path) {
      const std::string workbook_path = part_["officeDocument"];
      const std::string workbook_dir = dirName(workbook_path);
      std::string rels_xml_path =
        workbook_dir + "/_rels/" + baseName(workbook_path) + ".rels";
      rels_xml_path = removeLeadingSlashes(rels_xml_path);
      std::string rels_xml_file = zip_buffer(path, rels_xml_path);

      rapidxml::xml_document<> rels_xml;
      rels_xml.parse<rapidxml::parse_strip_xml_namespaces>(&rels_xml_file[0]);

      rapidxml::xml_node<>* relationships = rels_xml.first_node("Relationships");
      if (relationships == NULL) {
        return;
      }

      for (rapidxml::xml_node<>* relationship = relationships->first_node();
           relationship; relationship = relationship->next_sibling()) {

        rapidxml::xml_attribute<>* id = relationship->first_attribute("Id");
        rapidxml::xml_attribute<>* type = relationship->first_attribute("Type");
        rapidxml::xml_attribute<>* target = relationship->first_attribute("Target");

        if (id != NULL && type != NULL && target != NULL) {
          // store 'xl/blah' instead of '/xl/blah' (rare, but we've seen)
          std::string target_value = removeLeadingSlashes(target->value());
          // abbreviate Type
          // in XML: http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings
          // whereas we only want: sharedStrings
          std::string type_value = baseName(type->value());

          // prepend workbook_dir iff needed; yes, we've seen both forms
          // xl/blah --> xl/blah
          //    blah --> xl/blah
          if (target_value.substr(0, workbook_dir.size()) != workbook_dir) {
            target_value = workbook_dir + "/" + target_value;
          }

          if (type_value.compare("worksheet") == 0) { // worksheet
            target_[id->value()] = target_value;
          } else { // not a worksheet, e.g. styles or sharedStrings
            part_[type_value] = target_value;
          }
        }
      }
    }

  public:
    PackageRelations(const std::string& path) :
    n_(100),
    names_(n_),
    id_(n_),
    defn_n_(20),
    defn_names_(20),
    defn_context_(20),
    defn_ref_(20)
    {
      parse_package_rels(path);
      parse_workbook(path);
      parse_workbook_rels(path);
    }

    cpp11::strings names() const {
      return names_;
    }

    cpp11::list definedNames() const {
      cpp11::writable::list result(3);
      cpp11::strings names({"name", "context", "ref"});
      result.attr("names") = names;
      result[0] = defn_names_;
      result[1] = defn_context_;
      result[2] = defn_ref_;
      return result;
    }

    int n_sheets() const {
      return n_;
    }

    std::string target(int sheet_i) const {
      std::string id = cpp11::r_string(id_[sheet_i]);
      std::map<std::string, std::string>::const_iterator it = target_.find(id);
      if (it == target_.end()) {
        cpp11::stop("`%s` not found", id.c_str());
      }
      return it->second;
    }

    std::string part(std::string type) const {
      std::map<std::string, std::string>::const_iterator it = part_.find(type);
      if (it == part_.end()) {
        // e.g., sharedStrings may not be present
        // downstream functions should handle non-existent path
        return "";
      }
      return it->second;
    }
  }; // end of class PackageRelations

  // common to Xls[x]WorkBook
  std::string path_;
  bool is1904_;
  std::set<int> dateFormats_;

  // specific to XlsxWorkBook
  PackageRelations rel_;
  std::vector<std::string> stringTable_;

public:

  XlsxWorkBook(const std::string& path):
  path_(path),
  rel_(path)
  {
    is1904_ = uses1904();
    cacheStringTable();
    cacheDateFormats();
  }

  const std::string& path() const{
    return path_;
  }

  int n_sheets() const {
    return rel_.n_sheets();
  }

  cpp11::list definedNames() const {
    return rel_.definedNames();
  }

  cpp11::strings sheets() const {
    return rel_.names();
  }

  bool is1904() const {
    return is1904_;
  }

  const std::set<int>& dateFormats() const {
    return dateFormats_;
  }

  std::string sheetPath(int sheet_i) const {
    return rel_.target(sheet_i);
  }

  const std::vector<std::string>& stringTable() const {
    return stringTable_;
  }

private:

  void cacheStringTable() {
    if (!zip_has_file(path_, rel_.part("sharedStrings"))) {
      return;
    }

    std::string sharedStringsXml = zip_buffer(path_, rel_.part("sharedStrings"));
    rapidxml::xml_document<> sharedStrings;
    sharedStrings.parse<rapidxml::parse_strip_xml_namespaces>(&sharedStringsXml[0]);

    rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");
    if (sst == NULL) {
      return;
    }

    rapidxml::xml_attribute<>* count = sst->first_attribute("count");
    if (count != NULL) {
      int n = atoi(count->value());
      stringTable_.reserve(n);
    }

    // 18.4.8 si (String Item) [p1725]
    for (rapidxml::xml_node<>* string = sst->first_node();
         string; string = string->next_sibling()) {
      std::string out;
      if (parseString(string, &out)) {
        stringTable_.push_back(out);
      }
    }
  }

  void cacheDateFormats() {
    if (!zip_has_file(path_, rel_.part("styles"))) {
      return;
    }

    std::string stylesXml = zip_buffer(path_, rel_.part("styles"));
    rapidxml::xml_document<> styles;
    styles.parse<rapidxml::parse_strip_xml_namespaces>(&stylesXml[0]);

    rapidxml::xml_node<>* styleSheet = styles.first_node("styleSheet");
    if (styleSheet == NULL) {
      return;
    }

    // Figure out which custom formats are dates
    // 18.8.31 numFmts (Number Formats) p1793 5th ed
    std::set<int> customDateFormats;
    rapidxml::xml_node<>* numFmts = styleSheet->first_node("numFmts");
    if (numFmts != NULL) {
      // Example with just one custom format (non date, in this case)
      // <numFmts count="1">
      //   <numFmt numFmtId="166" formatCode="&quot;$&quot;#,##0.00"/>
      // </numFmts>
      for (rapidxml::xml_node<>* numFmt = numFmts->first_node();
           numFmt; numFmt = numFmt->next_sibling()) {
        // formatCode: The number format code for this number format.
        std::string code(numFmt->first_attribute("formatCode")->value());
        //   numFmtId: Id used by the master style records (xf's) to reference
        //             this number format.
        int id = atoi(numFmt->first_attribute("numFmtId")->value());

        if (isDateFormat(code)) {
          customDateFormats.insert(id);
        }
      }
    }

    // Cache 0-based indices of the master cell style records that refer to a
    // number format that is a date format
    rapidxml::xml_node<>* cellXfs = styleSheet->first_node("cellXfs");
    if (cellXfs == NULL) {
      return;
    }

    // inspecting i-th child of cellXfs = i-th cell style record = an <xf>
    int i = 0;
    for (rapidxml::xml_node<>* cellXf = cellXfs->first_node();
         cellXf; cellXf = cellXf->next_sibling()) {
      if (cellXf->first_attribute("numFmtId") == NULL) {
        ++i;
        continue;
      }
      int formatId = atoi(cellXf->first_attribute("numFmtId")->value());
      if (isDateTime(formatId, customDateFormats))
        dateFormats_.insert(i);
      ++i;
    }
  }

  bool uses1904() {
    std::string workbookXml = zip_buffer(path_, rel_.part("officeDocument"));
    rapidxml::xml_document<> workbook;
    workbook.parse<rapidxml::parse_strip_xml_namespaces>(&workbookXml[0]);

    rapidxml::xml_node<>* root = workbook.first_node("workbook");
    if (root == NULL) {
      return false;
    }

    rapidxml::xml_node<>* workbookPr = root->first_node("workbookPr");
    if (workbookPr == NULL) {
      return false;
    }

    rapidxml::xml_attribute<>* date1904 = workbookPr->first_attribute("date1904");
    if (date1904 == NULL) {
      return false;
    }

    return atoi(date1904->value()) == 1;
  }

};
