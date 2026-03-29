#include "docxcpp/document.hpp"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <fstream>
#include <unordered_map>
#include <sstream>
#include <stdexcept>
#include <string_view>
#include <utility>

#include "pugixml.hpp"

#include "internal/document_model_support.hpp"
#include "internal/header_footer_support.hpp"
#include "internal/hyperlink_comment_support.hpp"
#include "internal/layout_support.hpp"
#include "internal/media_support.hpp"
#include "internal/paragraph_support.hpp"

namespace docxcpp {

namespace {

constexpr const char* kDocumentEntry = "word/document.xml";
constexpr const char* kDocumentRelationshipsEntry = "word/_rels/document.xml.rels";
constexpr const char* kContentTypesEntry = "[Content_Types].xml";
constexpr const char* kTableStyleValue = "TableGrid";
constexpr const char* kImageRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kPictureUri =
    "http://schemas.openxmlformats.org/drawingml/2006/picture";

std::runtime_error missing_node(const char* node_name) {
  return std::runtime_error(std::string("required DOCX node is missing: ") + node_name);
}

bool needs_preserve_space(const std::string& text) {
  if (text.empty()) {
    return false;
  }
  if (text.front() == ' ' || text.back() == ' ') {
    return true;
  }
  return text.find("  ") != std::string::npos;
}

pugi::xml_node child_named(const pugi::xml_node& parent, const char* name);
std::size_t count_named_children(pugi::xml_node parent, const char* name);
Table table_from_xml(const pugi::xml_node& table_node);
Table create_table_in_parent(pugi::xml_node parent, std::size_t rows, std::size_t cols,
                             pugi::xml_node insert_before = {});

const char* alignment_value(ParagraphAlignment alignment) {
  switch (alignment) {
    case ParagraphAlignment::Left:
      return "left";
    case ParagraphAlignment::Center:
      return "center";
    case ParagraphAlignment::Right:
      return "right";
    case ParagraphAlignment::Justify:
      return "both";
    case ParagraphAlignment::Inherit:
    default:
      return nullptr;
  }
}

ParagraphAlignment alignment_from_value(std::string_view value) {
  if (value == "left") {
    return ParagraphAlignment::Left;
  }
  if (value == "center") {
    return ParagraphAlignment::Center;
  }
  if (value == "right") {
    return ParagraphAlignment::Right;
  }
  if (value == "both" || value == "justify") {
    return ParagraphAlignment::Justify;
  }
  return ParagraphAlignment::Inherit;
}

bool is_hex_color(std::string_view value) {
  if (value.size() != 6) {
    return false;
  }
  for (const char ch : value) {
    const bool is_digit = ch >= '0' && ch <= '9';
    const bool is_lower = ch >= 'a' && ch <= 'f';
    const bool is_upper = ch >= 'A' && ch <= 'F';
    if (!is_digit && !is_lower && !is_upper) {
      return false;
    }
  }
  return true;
}

std::string normalized_hex_color(std::string value) {
  if (!value.empty() && value.front() == '#') {
    value.erase(value.begin());
  }
  if (!is_hex_color(value)) {
    throw std::invalid_argument("color_hex must be a 6-digit RGB hex string");
  }
  std::transform(value.begin(), value.end(), value.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::toupper(ch)); });
  return value;
}

bool is_word_token(std::string_view value) {
  if (value.empty()) {
    return false;
  }
  for (const char ch : value) {
    const bool is_letter = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
    const bool is_digit = ch >= '0' && ch <= '9';
    const bool is_dash = ch == '-';
    if (!is_letter && !is_digit && !is_dash) {
      return false;
    }
  }
  return true;
}

std::string normalized_highlight(std::string value) {
  std::transform(value.begin(), value.end(), value.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  if (!is_word_token(value)) {
    throw std::invalid_argument("highlight must be a Word highlight token like yellow or green");
  }
  return value;
}

void collect_text(const pugi::xml_node& node, std::string& text) {
  const std::string_view name = node.name();
  if (name == "w:pPr" || name == "w:rPr") {
    return;
  }
  if (name == "w:t") {
    text += node.text().get();
  } else if (name == "w:tab") {
    text.push_back('\t');
  } else if (name == "w:br" || name == "w:cr") {
    text.push_back('\n');
  }

  for (const pugi::xml_node& child : node.children()) {
    collect_text(child, text);
  }
}

bool paragraph_has_page_break(const pugi::xml_node& paragraph) {
  if (std::string_view(paragraph.name()) == "w:br" &&
      std::string_view(paragraph.attribute("w:type").value()) == "page") {
    return true;
  }
  for (const pugi::xml_node& child : paragraph.children()) {
    if (paragraph_has_page_break(child)) {
      return true;
    }
  }
  return false;
}

ParagraphAlignment paragraph_alignment_from_xml(const pugi::xml_node& paragraph) {
  const pugi::xml_node p_pr = child_named(paragraph, "w:pPr");
  if (!p_pr) {
    return ParagraphAlignment::Inherit;
  }
  const pugi::xml_node jc = child_named(p_pr, "w:jc");
  if (!jc) {
    return ParagraphAlignment::Inherit;
  }
  return alignment_from_value(jc.attribute("w:val").value());
}

std::vector<std::uint8_t> xml_to_bytes(const pugi::xml_document& xml) {
  std::ostringstream stream;
  xml.save(stream, "  ", pugi::format_default, pugi::encoding_utf8);
  const std::string content = stream.str();
  return std::vector<std::uint8_t>(content.begin(), content.end());
}

pugi::xml_node child_named(const pugi::xml_node& parent, const char* name) {
  for (const pugi::xml_node& child : parent.children()) {
    if (std::string_view(child.name()) == name) {
      return child;
    }
  }
  return pugi::xml_node();
}

pugi::xml_node document_body(const pugi::xml_document& xml);
std::size_t next_docpr_id(const pugi::xml_node& node, std::size_t current_max);
pugi::xml_document parse_package_xml(const OpcPackage& package, const char* entry_name);

pugi::xml_node nth_child_named(pugi::xml_node parent, const char* name, std::size_t index) {
  std::size_t current = 0;
  for (pugi::xml_node child : parent.children(name)) {
    if (current == index) {
      return child;
    }
    ++current;
  }
  return pugi::xml_node();
}

pugi::xml_node require_table_node(pugi::xml_document& xml, std::size_t table_index) {
  pugi::xml_node table = nth_child_named(document_body(xml), "w:tbl", table_index);
  if (!table) {
    throw std::out_of_range("table_index out of range");
  }
  return table;
}

pugi::xml_node require_row_node(pugi::xml_node table, std::size_t row_index) {
  pugi::xml_node row = nth_child_named(table, "w:tr", row_index);
  if (!row) {
    throw std::out_of_range("row_index out of range");
  }
  return row;
}

pugi::xml_node cell_properties_node(pugi::xml_node cell) { return child_named(cell, "w:tcPr"); }

std::size_t cell_grid_span(pugi::xml_node cell) {
  if (const pugi::xml_node tc_pr = cell_properties_node(cell)) {
    if (const pugi::xml_node grid_span_node = child_named(tc_pr, "w:gridSpan")) {
      return static_cast<std::size_t>(grid_span_node.attribute("w:val").as_uint(1));
    }
  }
  return 1;
}

pugi::xml_node require_physical_cell_node(pugi::xml_node row, std::size_t physical_index) {
  pugi::xml_node cell = nth_child_named(row, "w:tc", physical_index);
  if (!cell) {
    throw std::out_of_range("col_index out of range");
  }
  return cell;
}

pugi::xml_node require_nested_table_node(pugi::xml_node cell, std::size_t nested_table_index) {
  pugi::xml_node table = nth_child_named(cell, "w:tbl", nested_table_index);
  if (!table) {
    throw std::out_of_range("nested_table_index out of range");
  }
  return table;
}

pugi::xml_node require_cell_node(pugi::xml_node row, std::size_t col_index) {
  std::size_t current_col = 0;
  for (pugi::xml_node cell : row.children("w:tc")) {
    const std::size_t span = cell_grid_span(cell);
    if (col_index < current_col + span) {
      return cell;
    }
    current_col += span;
  }
  throw std::out_of_range("col_index out of range");
}

std::string table_grid_column_width(pugi::xml_node table, std::size_t column_index) {
  if (const pugi::xml_node grid = child_named(table, "w:tblGrid")) {
    if (const pugi::xml_node grid_col = nth_child_named(grid, "w:gridCol", column_index)) {
      const std::string width = grid_col.attribute("w:w").value();
      if (!width.empty()) {
        return width;
      }
    }
  }
  return "2390";
}

std::size_t table_column_count(pugi::xml_node table) {
  if (const pugi::xml_node grid = child_named(table, "w:tblGrid")) {
    const std::size_t grid_count = count_named_children(grid, "w:gridCol");
    if (grid_count > 0) {
      return grid_count;
    }
  }

  std::size_t max_columns = 0;
  for (const pugi::xml_node& row : table.children("w:tr")) {
    std::size_t row_columns = 0;
    for (const pugi::xml_node& cell : row.children("w:tc")) {
      std::size_t grid_span = 1;
      if (const pugi::xml_node tc_pr = child_named(cell, "w:tcPr")) {
        if (const pugi::xml_node grid_span_node = child_named(tc_pr, "w:gridSpan")) {
          grid_span = static_cast<std::size_t>(grid_span_node.attribute("w:val").as_uint(1));
        }
      }
      row_columns += grid_span;
    }
    max_columns = std::max(max_columns, row_columns);
  }
  return max_columns;
}

pugi::xml_node append_empty_table_cell(pugi::xml_node row, const std::string& width) {
  pugi::xml_node cell = row.append_child("w:tc");
  pugi::xml_node tc_pr = cell.append_child("w:tcPr");
  tc_pr.append_child("w:tcW").append_attribute("w:w").set_value(width.c_str());
  tc_pr.child("w:tcW").append_attribute("w:type").set_value("dxa");
  pugi::xml_node paragraph = cell.append_child("w:p");
  set_paragraph_text_in_xml(paragraph, "");
  return cell;
}

void append_table_grid_column(pugi::xml_node table, const std::string& width) {
  pugi::xml_node grid = child_named(table, "w:tblGrid");
  if (!grid) {
    grid = table.prepend_child("w:tblGrid");
  }
  grid.append_child("w:gridCol").append_attribute("w:w").set_value(width.c_str());
}

Table create_table_in_parent(pugi::xml_node parent, std::size_t rows, std::size_t cols,
                             pugi::xml_node insert_before) {
  if (rows == 0 || cols == 0) {
    throw std::invalid_argument("table rows and cols must be greater than zero");
  }

  pugi::xml_node table = insert_before ? parent.insert_child_before("w:tbl", insert_before)
                                       : parent.append_child("w:tbl");
  pugi::xml_node properties = table.append_child("w:tblPr");
  pugi::xml_node style = properties.append_child("w:tblStyle");
  style.append_attribute("w:val").set_value(kTableStyleValue);
  properties.append_child("w:tblW").append_attribute("w:w").set_value("0");
  properties.child("w:tblW").append_attribute("w:type").set_value("auto");

  pugi::xml_node grid = table.append_child("w:tblGrid");
  for (std::size_t col = 0; col < cols; ++col) {
    grid.append_child("w:gridCol").append_attribute("w:w").set_value("2390");
  }

  std::vector<TableRow> table_rows;
  table_rows.reserve(rows);
  for (std::size_t row_index = 0; row_index < rows; ++row_index) {
    pugi::xml_node row = table.append_child("w:tr");
    std::vector<TableCell> table_cells;
    table_cells.reserve(cols);
    for (std::size_t col_index = 0; col_index < cols; ++col_index) {
      append_empty_table_cell(row, table_grid_column_width(table, col_index));
      table_cells.emplace_back("");
    }
    table_rows.emplace_back(std::move(table_cells));
  }

  return Table(std::move(table_rows));
}

pugi::xml_node get_or_add_child(pugi::xml_node parent, const char* name) {
  pugi::xml_node child = child_named(parent, name);
  if (!child) {
    child = parent.append_child(name);
  }
  return child;
}

pugi::xml_node get_or_add_paragraph_properties(pugi::xml_node paragraph) {
  pugi::xml_node p_pr = child_named(paragraph, "w:pPr");
  if (!p_pr) {
    p_pr = paragraph.prepend_child("w:pPr");
  }
  return p_pr;
}

pugi::xml_node document_body(const pugi::xml_document& xml) {
  pugi::xml_node document = child_named(xml, "w:document");
  if (!document) {
    throw missing_node("w:document");
  }
  pugi::xml_node body = child_named(document, "w:body");
  if (!body) {
    throw missing_node("w:body");
  }
  return body;
}

void ensure_picture_namespaces(pugi::xml_document& xml) {
  pugi::xml_node document = child_named(xml, "w:document");
  if (!document) {
    throw missing_node("w:document");
  }
  auto ensure_attr = [&](const char* name, const char* value) {
    if (!document.attribute(name)) {
      document.append_attribute(name).set_value(value);
    }
  };
  ensure_attr("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
  ensure_attr("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
}

pugi::xml_node append_block_before_section(pugi::xml_node body, const char* name) {
  pugi::xml_node sect_pr = child_named(body, "w:sectPr");
  return sect_pr ? body.insert_child_before(name, sect_pr) : body.append_child(name);
}

void clear_children(pugi::xml_node node) {
  while (node.first_child()) {
    node.remove_child(node.first_child());
  }
}

void apply_alignment(pugi::xml_node paragraph, ParagraphAlignment alignment) {
  pugi::xml_node properties = child_named(paragraph, "w:pPr");
  if (alignment == ParagraphAlignment::Inherit) {
    if (!properties) {
      return;
    }
    pugi::xml_node jc = child_named(properties, "w:jc");
    if (jc) {
      properties.remove_child(jc);
    }
    if (!properties.first_child()) {
      paragraph.remove_child(properties);
    }
    return;
  }

  if (!properties) {
    properties = paragraph.prepend_child("w:pPr");
  }
  pugi::xml_node jc = child_named(properties, "w:jc");
  if (!jc) {
    jc = properties.append_child("w:jc");
  }
  if (jc.attribute("w:val")) {
    jc.attribute("w:val").set_value(alignment_value(alignment));
  } else {
    jc.append_attribute("w:val").set_value(alignment_value(alignment));
  }
}

void apply_run_style(pugi::xml_node run, const RunStyle& style) {
  if (!style.bold && !style.italic && !style.underline && style.font_size_pt <= 0 &&
      style.color_hex.empty() && style.font_name.empty() && style.highlight.empty()) {
    return;
  }

  pugi::xml_node properties = child_named(run, "w:rPr");
  if (!properties) {
    properties = run.prepend_child("w:rPr");
  }
  if (!style.font_name.empty()) {
    pugi::xml_node fonts = properties.append_child("w:rFonts");
    fonts.append_attribute("w:ascii").set_value(style.font_name.c_str());
    fonts.append_attribute("w:hAnsi").set_value(style.font_name.c_str());
  }
  if (style.bold) {
    properties.append_child("w:b");
  }
  if (style.italic) {
    properties.append_child("w:i");
  }
  if (style.underline) {
    pugi::xml_node underline = properties.append_child("w:u");
    underline.append_attribute("w:val").set_value("single");
  }
  if (style.font_size_pt > 0) {
    pugi::xml_node size = properties.append_child("w:sz");
    size.append_attribute("w:val").set_value(std::to_string(style.font_size_pt * 2).c_str());
  }
  if (!style.color_hex.empty()) {
    pugi::xml_node color = properties.append_child("w:color");
    color.append_attribute("w:val").set_value(normalized_hex_color(style.color_hex).c_str());
  }
  if (!style.highlight.empty()) {
    pugi::xml_node highlight = properties.append_child("w:highlight");
    highlight.append_attribute("w:val").set_value(normalized_highlight(style.highlight).c_str());
  }
}

void apply_hyperlink_character_style(pugi::xml_node run) {
  pugi::xml_node properties = child_named(run, "w:rPr");
  if (!properties) {
    properties = run.prepend_child("w:rPr");
  }
  if (!child_named(properties, "w:rStyle")) {
    pugi::xml_node style = properties.append_child("w:rStyle");
    style.append_attribute("w:val").set_value("Hyperlink");
  }
}

void append_run_text(pugi::xml_node run, const std::string& text) {
  std::size_t start = 0;
  bool wrote_text = false;
  while (start <= text.size()) {
    const std::size_t newline = text.find('\n', start);
    const std::string segment =
        newline == std::string::npos ? text.substr(start) : text.substr(start, newline - start);

    if (!segment.empty() || !wrote_text) {
      pugi::xml_node text_node = run.append_child("w:t");
      text_node.text().set(segment.c_str());
      if (needs_preserve_space(segment)) {
        text_node.append_attribute("xml:space").set_value("preserve");
      }
      wrote_text = true;
    }

    if (newline == std::string::npos) {
      break;
    }
    run.append_child("w:br");
    start = newline + 1;
  }
}

Table table_from_xml(const pugi::xml_node& table_node) {
  std::vector<TableRow> rows;
  for (const pugi::xml_node& row_node : table_node.children("w:tr")) {
    std::vector<TableCell> cells;
    for (const pugi::xml_node& cell_node : row_node.children("w:tc")) {
      std::string cell_text;
      std::vector<Table> nested_tables;
      for (const pugi::xml_node& child : cell_node.children()) {
        if (std::string_view(child.name()) == "w:tbl") {
          nested_tables.push_back(table_from_xml(child));
          continue;
        }
        if (std::string_view(child.name()) != "w:p") {
          continue;
        }
        std::string paragraph_text;
        collect_text(child, paragraph_text);
        if (!cell_text.empty()) {
          cell_text.push_back('\n');
        }
        cell_text += paragraph_text;
      }
      std::size_t grid_span = cell_grid_span(cell_node);
      std::string vertical_merge;
      if (const pugi::xml_node tc_pr = cell_properties_node(cell_node)) {
        if (const pugi::xml_node v_merge_node = child_named(tc_pr, "w:vMerge")) {
          vertical_merge = v_merge_node.attribute("w:val").value();
          if (vertical_merge.empty()) {
            vertical_merge = "continue";
          }
        }
      }
      cells.emplace_back(cell_text, grid_span, vertical_merge, std::move(nested_tables));
    }
    rows.emplace_back(std::move(cells));
  }
  return Table(std::move(rows));
}

enum class ParagraphBindingKind {
  Body,
  TableCell,
};

std::shared_ptr<ParagraphBinding> bind_body_paragraph(pugi::xml_document& xml, bool& dirty,
                                                      std::size_t paragraph_index) {
  auto binding = std::make_shared<ParagraphBinding>();
  binding->xml = &xml;
  binding->dirty = &dirty;
  binding->kind = static_cast<int>(ParagraphBindingKind::Body);
  binding->paragraph_index = paragraph_index;
  return binding;
}

std::shared_ptr<ParagraphBinding> bind_table_cell_paragraph(pugi::xml_document& xml, bool& dirty,
                                                            std::size_t table_index,
                                                            std::size_t row_index,
                                                            std::size_t col_index,
                                                            std::size_t cell_paragraph_index) {
  auto binding = std::make_shared<ParagraphBinding>();
  binding->xml = &xml;
  binding->dirty = &dirty;
  binding->kind = static_cast<int>(ParagraphBindingKind::TableCell);
  binding->table_index = table_index;
  binding->row_index = row_index;
  binding->col_index = col_index;
  binding->cell_paragraph_index = cell_paragraph_index;
  return binding;
}

std::size_t count_named_children(pugi::xml_node parent, const char* name) {
  std::size_t count = 0;
  for (const pugi::xml_node& child : parent.children(name)) {
    (void)child;
    ++count;
  }
  return count;
}

std::size_t body_paragraph_count(const pugi::xml_document& xml) {
  return count_named_children(document_body(xml), "w:p");
}

std::size_t paragraph_index_in_cell(pugi::xml_node cell) {
  return count_named_children(cell, "w:p");
}

pugi::xml_node paragraph_node_from_binding(const ParagraphBinding& binding) {
  if (binding.xml == nullptr) {
    return {};
  }
  if (binding.kind == static_cast<int>(ParagraphBindingKind::Body)) {
    return nth_child_named(document_body(*binding.xml), "w:p", binding.paragraph_index);
  }
  pugi::xml_node table = require_table_node(*binding.xml, binding.table_index);
  pugi::xml_node row = require_row_node(table, binding.row_index);
  pugi::xml_node cell = require_cell_node(row, binding.col_index);
  return nth_child_named(cell, "w:p", binding.cell_paragraph_index);
}

std::size_t parse_numeric_suffix(std::string_view value, std::string_view prefix) {
  if (value.size() < prefix.size() || value.substr(0, prefix.size()) != prefix) {
    return 0;
  }
  std::size_t number = 0;
  for (const char ch : value.substr(prefix.size())) {
    if (ch < '0' || ch > '9') {
      return 0;
    }
    number = number * 10 + static_cast<std::size_t>(ch - '0');
  }
  return number;
}

std::size_t next_relationship_id(const pugi::xml_document& rels_xml) {
  std::size_t max_id = 0;
  pugi::xml_node relationships = child_named(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    max_id = std::max(max_id, parse_numeric_suffix(rel.attribute("Id").value(), "rId"));
  }
  return max_id + 1;
}

std::size_t next_image_index(const OpcPackage& package) {
  std::size_t max_index = 0;
  for (std::size_t i = 1; i < 10000; ++i) {
    for (const auto& ext : {"png", "jpg", "jpeg"}) {
      const std::string name = "word/media/image" + std::to_string(i) + "." + ext;
      if (package.has_entry(name)) {
        max_index = i;
      }
    }
  }
  return max_index + 1;
}

void ensure_content_type_default(pugi::xml_document& content_types_xml, const std::string& extension,
                                 const std::string& content_type) {
  pugi::xml_node types = child_named(content_types_xml, "Types");
  for (const pugi::xml_node& node : types.children("Default")) {
    if (extension == node.attribute("Extension").value()) {
      return;
    }
  }
  pugi::xml_node entry = types.append_child("Default");
  entry.append_attribute("Extension").set_value(extension.c_str());
  entry.append_attribute("ContentType").set_value(content_type.c_str());
}

void ensure_content_type_override(pugi::xml_document& content_types_xml, const std::string& part_name,
                                  const std::string& content_type) {
  pugi::xml_node types = child_named(content_types_xml, "Types");
  for (pugi::xml_node node : types.children("Override")) {
    if (part_name == node.attribute("PartName").value()) {
      if (node.attribute("ContentType")) {
        node.attribute("ContentType").set_value(content_type.c_str());
      } else {
        node.append_attribute("ContentType").set_value(content_type.c_str());
      }
      return;
    }
  }
  pugi::xml_node entry = types.append_child("Override");
  entry.append_attribute("PartName").set_value(part_name.c_str());
  entry.append_attribute("ContentType").set_value(content_type.c_str());
}

std::string register_picture_part(OpcPackage& package, const ImageInfo& image) {
  const std::size_t image_index = next_image_index(package);
  const std::string media_name =
      "word/media/image" + std::to_string(image_index) + "." + image.extension;
  package.set_entry(media_name, image.bytes);

  pugi::xml_document rels_xml = parse_package_xml(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named(rels_xml, "Relationships");
  const std::string rel_id = "rId" + std::to_string(next_relationship_id(rels_xml));
  pugi::xml_node rel = relationships.append_child("Relationship");
  rel.append_attribute("Id").set_value(rel_id.c_str());
  rel.append_attribute("Type").set_value(kImageRelationshipType);
  rel.append_attribute("Target")
      .set_value(("media/image" + std::to_string(image_index) + "." + image.extension).c_str());
  package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes(rels_xml));

  pugi::xml_document content_types_xml = parse_package_xml(package, kContentTypesEntry);
  ensure_content_type_default(content_types_xml, image.extension, image.content_type);
  package.set_entry(kContentTypesEntry, xml_to_bytes(content_types_xml));
  return rel_id;
}

std::size_t next_docpr_id(const pugi::xml_node& node, std::size_t current_max = 0) {
  if (std::string_view(node.name()) == "wp:docPr") {
    current_max = std::max(current_max, static_cast<std::size_t>(node.attribute("id").as_uint()));
  }
  for (const pugi::xml_node& child : node.children()) {
    current_max = next_docpr_id(child, current_max);
  }
  return current_max;
}

std::size_t count_named_descendants(const pugi::xml_node& node, std::string_view name) {
  std::size_t count = std::string_view(node.name()) == name ? 1 : 0;
  for (const pugi::xml_node& child : node.children()) {
    count += count_named_descendants(child, name);
  }
  return count;
}

pugi::xml_document parse_package_xml(const OpcPackage& package, const char* entry_name) {
  pugi::xml_document xml;
  const auto& bytes = package.entry(entry_name);
  const auto result =
      xml.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!result) {
    throw std::runtime_error(std::string("failed to parse package xml: ") + entry_name);
  }
  return xml;
}

}  // namespace

pugi::xml_node paragraph_node_for_model_binding(const ParagraphBinding& binding) {
  return paragraph_node_from_binding(binding);
}

void apply_run_style_for_model(pugi::xml_node run, const RunStyle& style) {
  apply_run_style(run, style);
}

void append_run_text_for_model(pugi::xml_node run, const std::string& text) {
  append_run_text(run, text);
}

Document::Document() : package_(OpcPackage::from_directory(DOCXCPP_DEFAULT_TEMPLATE_DIR)) {
  load_document_xml();
}

Document::Document(const std::filesystem::path& path) : package_(OpcPackage::open(path)) {
  load_document_xml();
}

Document Document::open(const std::filesystem::path& path) { return Document(path); }

Paragraph Document::add_paragraph(const std::string& text) {
  return add_styled_paragraph(text, {});
}

Paragraph Document::add_paragraph(const std::vector<Run>& runs, ParagraphAlignment alignment) {
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  apply_alignment(paragraph, alignment);
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
  Paragraph result = paragraph_from_runs(runs, alignment);
  result = Paragraph(result.text(), result.alignment(), result.first_run_style(),
                     result.has_page_break(), result.runs(), result.style_id(),
                     result.heading_level(), result.format(), {}, bind_body_paragraph(*xml_, dirty_, paragraph_index));
  return result;
}

Paragraph Document::add_page_break() {
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  pugi::xml_node run = paragraph.append_child("w:r");
  pugi::xml_node br = run.append_child("w:br");
  br.append_attribute("w:type").set_value("page");
  dirty_ = true;
  return Paragraph("\n", ParagraphAlignment::Inherit, {}, true, {Run("\n", {})}, {}, -1, {}, {},
                   bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Paragraph Document::add_styled_paragraph(const std::string& text, const RunStyle& style,
                                         ParagraphAlignment alignment) {
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  apply_alignment(paragraph, alignment);
  set_paragraph_text_in_xml(paragraph, text, style);
  dirty_ = true;
  return Paragraph(text, alignment, style, false, {Run(text, style)}, {}, -1, {}, {},
                   bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Paragraph Document::add_heading(const std::string& text, int level) {
  return add_styled_heading(text, level, {});
}

Paragraph Document::add_heading(const std::vector<Run>& runs, int level,
                                ParagraphAlignment alignment) {
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  pugi::xml_node properties = paragraph.append_child("w:pPr");
  const std::string style_id = paragraph_style_id_for_level(level);
  pugi::xml_node style_node = properties.append_child("w:pStyle");
  style_node.append_attribute("w:val").set_value(style_id.c_str());
  apply_alignment(paragraph, alignment);
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
  Paragraph result = paragraph_from_runs(runs, alignment, style_id, level);
  return Paragraph(result.text(), result.alignment(), result.first_run_style(), result.has_page_break(),
                   result.runs(), style_id, level, result.format(), {},
                   bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Paragraph Document::add_styled_heading(const std::string& text, int level, const RunStyle& run_style,
                                       ParagraphAlignment alignment) {
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  pugi::xml_node properties = paragraph.append_child("w:pPr");
  const std::string style_id = paragraph_style_id_for_level(level);
  pugi::xml_node style_node = properties.append_child("w:pStyle");
  style_node.append_attribute("w:val").set_value(style_id.c_str());
  apply_alignment(paragraph, alignment);
  set_paragraph_text_in_xml(paragraph, text, run_style);
  dirty_ = true;
  return Paragraph(text, alignment, run_style, false, {Run(text, run_style)}, style_id, level, {},
                   {}, bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Table Document::add_table(std::size_t rows, std::size_t cols) {
  Table result = create_table_in_parent(document_body(*xml_), rows, cols,
                                        child_named(document_body(*xml_), "w:sectPr"));
  dirty_ = true;
  return result;
}

Table Document::add_nested_table(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 std::size_t rows, std::size_t cols) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);

  pugi::xml_node trailing_paragraph;
  if (const pugi::xml_node last_child = cell.last_child();
      last_child && std::string_view(last_child.name()) == "w:p") {
    trailing_paragraph = last_child;
    const bool only_child = cell.first_child() == last_child;
    const bool empty_text = trailing_paragraph.text().as_string()[0] == '\0' &&
                            !child_named(trailing_paragraph, "w:r") &&
                            !child_named(trailing_paragraph, "w:hyperlink");
    if (only_child && empty_text) {
      cell.remove_child(trailing_paragraph);
      trailing_paragraph = {};
    }
  }

  Table result = create_table_in_parent(cell, rows, cols, trailing_paragraph);
  if (!trailing_paragraph) {
    pugi::xml_node paragraph = cell.append_child("w:p");
    set_paragraph_text_in_xml(paragraph, "");
  }
  dirty_ = true;
  return result;
}

void Document::add_table_row(std::size_t table_index) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  const std::size_t cols = table_column_count(table);
  if (cols == 0) {
    throw std::out_of_range("table has no columns");
  }

  pugi::xml_node row = table.append_child("w:tr");
  for (std::size_t col_index = 0; col_index < cols; ++col_index) {
    append_empty_table_cell(row, table_grid_column_width(table, col_index));
  }
  dirty_ = true;
}

void Document::add_table_column(std::size_t table_index) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  const std::size_t previous_cols = table_column_count(table);
  if (previous_cols == 0) {
    throw std::out_of_range("table has no columns");
  }

  const std::string width = table_grid_column_width(table, previous_cols - 1);
  append_table_grid_column(table, width);
  for (pugi::xml_node row : table.children("w:tr")) {
    append_empty_table_cell(row, width);
  }
  dirty_ = true;
}

void Document::set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                              const std::string& text) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);

  pugi::xml_node paragraph = child_named(cell, "w:p");
  if (!paragraph) {
    paragraph = cell.append_child("w:p");
  }
  set_paragraph_text_in_xml(paragraph, text);
  dirty_ = true;
}

void Document::set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                              const std::vector<Run>& runs) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);

  pugi::xml_node paragraph = child_named(cell, "w:p");
  if (!paragraph) {
    paragraph = cell.append_child("w:p");
  }
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
}

void Document::set_nested_table_cell(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, std::size_t nested_table_index,
                                     std::size_t nested_row_index, std::size_t nested_col_index,
                                     const std::string& text) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node nested_table = require_nested_table_node(cell, nested_table_index);
  pugi::xml_node nested_row = require_row_node(nested_table, nested_row_index);
  pugi::xml_node nested_cell = require_cell_node(nested_row, nested_col_index);

  pugi::xml_node paragraph = child_named(nested_cell, "w:p");
  if (!paragraph) {
    paragraph = nested_cell.append_child("w:p");
  }
  set_paragraph_text_in_xml(paragraph, text);
  dirty_ = true;
}

void Document::set_nested_table_cell(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, std::size_t nested_table_index,
                                     std::size_t nested_row_index, std::size_t nested_col_index,
                                     const std::vector<Run>& runs) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node nested_table = require_nested_table_node(cell, nested_table_index);
  pugi::xml_node nested_row = require_row_node(nested_table, nested_row_index);
  pugi::xml_node nested_cell = require_cell_node(nested_row, nested_col_index);

  pugi::xml_node paragraph = child_named(nested_cell, "w:p");
  if (!paragraph) {
    paragraph = nested_cell.append_child("w:p");
  }
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
}

Paragraph Document::add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                             std::size_t col_index, const std::string& text) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  const std::size_t cell_paragraph_index = paragraph_index_in_cell(cell);
  pugi::xml_node paragraph = cell.append_child("w:p");
  set_paragraph_text_in_xml(paragraph, text);
  dirty_ = true;
  return Paragraph(text, ParagraphAlignment::Inherit, {}, false, {Run(text, {})}, {}, -1, {}, {},
                   bind_table_cell_paragraph(*xml_, dirty_, table_index, row_index, col_index,
                                             cell_paragraph_index));
}

Paragraph Document::add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                             std::size_t col_index,
                                             const std::vector<Run>& runs) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  const std::size_t cell_paragraph_index = paragraph_index_in_cell(cell);
  pugi::xml_node paragraph = cell.append_child("w:p");
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
  Paragraph result = paragraph_from_runs(runs, ParagraphAlignment::Inherit);
  return Paragraph(result.text(), result.alignment(), result.first_run_style(), result.has_page_break(),
                   result.runs(), result.style_id(), result.heading_level(), result.format(), {},
                   bind_table_cell_paragraph(*xml_, dirty_, table_index, row_index, col_index,
                                             cell_paragraph_index));
}

Paragraph Document::add_nested_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                                    std::size_t col_index,
                                                    std::size_t nested_table_index,
                                                    std::size_t nested_row_index,
                                                    std::size_t nested_col_index,
                                                    const std::string& text) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node nested_table = require_nested_table_node(cell, nested_table_index);
  pugi::xml_node nested_row = require_row_node(nested_table, nested_row_index);
  pugi::xml_node nested_cell = require_cell_node(nested_row, nested_col_index);
  pugi::xml_node paragraph = nested_cell.append_child("w:p");
  set_paragraph_text_in_xml(paragraph, text);
  dirty_ = true;
  return Paragraph(text, ParagraphAlignment::Inherit, {}, false, {Run(text, {})}, {}, -1, {}, {},
                   nullptr);
}

Paragraph Document::add_nested_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                                    std::size_t col_index,
                                                    std::size_t nested_table_index,
                                                    std::size_t nested_row_index,
                                                    std::size_t nested_col_index,
                                                    const std::vector<Run>& runs) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node nested_table = require_nested_table_node(cell, nested_table_index);
  pugi::xml_node nested_row = require_row_node(nested_table, nested_row_index);
  pugi::xml_node nested_cell = require_cell_node(nested_row, nested_col_index);
  pugi::xml_node paragraph = nested_cell.append_child("w:p");
  set_paragraph_runs_in_xml(paragraph, runs);
  dirty_ = true;
  Paragraph result = paragraph_from_runs(runs, ParagraphAlignment::Inherit);
  return Paragraph(result.text(), result.alignment(), result.first_run_style(), result.has_page_break(),
                   result.runs(), result.style_id(), result.heading_level(), result.format(), {},
                   nullptr);
}

void Document::merge_table_cells(std::size_t table_index, std::size_t start_row, std::size_t start_col,
                                 std::size_t end_row, std::size_t end_col) {
  if (end_row < start_row || end_col < start_col) {
    throw std::invalid_argument("merge range must be top-left to bottom-right");
  }

  pugi::xml_node table = require_table_node(*xml_, table_index);
  const std::size_t row_span = end_row - start_row + 1;
  const std::size_t col_span = end_col - start_col + 1;

  for (std::size_t row_index = start_row; row_index <= end_row; ++row_index) {
    pugi::xml_node row = require_row_node(table, row_index);
    pugi::xml_node lead_cell = require_cell_node(row, start_col);
    pugi::xml_node tc_pr = child_named(lead_cell, "w:tcPr");
    if (!tc_pr) {
      tc_pr = lead_cell.prepend_child("w:tcPr");
    }

    if (col_span > 1) {
      pugi::xml_node grid_span = child_named(tc_pr, "w:gridSpan");
      if (!grid_span) {
        grid_span = tc_pr.append_child("w:gridSpan");
      }
      if (grid_span.attribute("w:val")) {
        grid_span.attribute("w:val").set_value(std::to_string(col_span).c_str());
      } else {
        grid_span.append_attribute("w:val").set_value(std::to_string(col_span).c_str());
      }

      for (std::size_t col_index = end_col; col_index > start_col; --col_index) {
        pugi::xml_node extra_cell = require_physical_cell_node(row, col_index);
        row.remove_child(extra_cell);
      }
    }

    if (row_span > 1) {
      pugi::xml_node v_merge = child_named(tc_pr, "w:vMerge");
      if (!v_merge) {
        v_merge = tc_pr.append_child("w:vMerge");
      }
      const char* merge_val = row_index == start_row ? "restart" : "continue";
      if (v_merge.attribute("w:val")) {
        v_merge.attribute("w:val").set_value(merge_val);
      } else {
        v_merge.append_attribute("w:val").set_value(merge_val);
      }
    }
  }

  dirty_ = true;
}

void Document::set_page_size_pt(int width_pt, int height_pt) {
  if (width_pt <= 0 || height_pt <= 0) {
    throw std::invalid_argument("page size must be greater than zero");
  }
  set_page_size_pt_in_xml(*xml_, width_pt, height_pt);
  dirty_ = true;
}

void Document::set_page_orientation(PageOrientation orientation) {
  set_page_orientation_in_xml(*xml_, orientation);
  dirty_ = true;
}

void Document::set_page_margins_pt(const PageMargins& margins) {
  if (margins.top_pt < 0 || margins.right_pt < 0 || margins.bottom_pt < 0 || margins.left_pt < 0 ||
      margins.header_pt < 0 || margins.footer_pt < 0 || margins.gutter_pt < 0) {
    throw std::invalid_argument("page margins cannot be negative");
  }
  set_page_margins_pt_in_xml(*xml_, margins);
  dirty_ = true;
}

void Document::add_section_break() {
  add_section_break(Section(page_size_pt(), page_orientation(), page_margins_pt()));
}

void Document::add_section_break(const Section& next_section) {
  const PageSize& size = next_section.page_size_pt();
  const PageMargins& margins = next_section.page_margins_pt();
  if (size.width_pt <= 0 || size.height_pt <= 0) {
    throw std::invalid_argument("section page size must be greater than zero");
  }
  if (margins.top_pt < 0 || margins.right_pt < 0 || margins.bottom_pt < 0 || margins.left_pt < 0 ||
      margins.header_pt < 0 || margins.footer_pt < 0 || margins.gutter_pt < 0) {
    throw std::invalid_argument("section page margins cannot be negative");
  }

  pugi::xml_node body = document_body(*xml_);
  pugi::xml_node body_sect_pr = child_named(body, "w:sectPr");
  if (!body_sect_pr) {
    body_sect_pr = body.append_child("w:sectPr");
    set_section_properties_in_xml(body_sect_pr,
                                  Section(page_size_pt(), page_orientation(), page_margins_pt()));
  }

  pugi::xml_node paragraph = body.insert_child_before("w:p", body_sect_pr);
  pugi::xml_node p_pr = paragraph.append_child("w:pPr");
  p_pr.append_copy(body_sect_pr);
  set_section_properties_in_xml(body_sect_pr, next_section);
  dirty_ = true;
}

void Document::set_even_and_odd_headers(bool enabled) {
  set_even_and_odd_headers_in_settings(package_, enabled);
  dirty_ = true;
}

bool Document::even_and_odd_headers() const { return read_even_and_odd_headers_from_settings(package_); }

void Document::set_section_different_first_page(std::size_t section_index, bool enabled) {
  set_section_different_first_page_in_xml(*xml_, section_index, enabled);
  dirty_ = true;
}

bool Document::section_different_first_page(std::size_t section_index) const {
  return read_section_different_first_page_from_xml(*xml_, section_index);
}

void Document::clear_header(std::size_t section_index) {
  clear_section_header(package_, *xml_, section_index, HeaderFooterType::Default);
  dirty_ = true;
}

void Document::clear_footer(std::size_t section_index) {
  clear_section_footer(package_, *xml_, section_index, HeaderFooterType::Default);
  dirty_ = true;
}

void Document::set_header_text(std::size_t section_index, const std::string& text, HeaderFooterType type) {
  set_section_header_text(package_, *xml_, section_index, text, type);
  dirty_ = true;
}

void Document::set_footer_text(std::size_t section_index, const std::string& text, HeaderFooterType type) {
  set_section_footer_text(package_, *xml_, section_index, text, type);
  dirty_ = true;
}

Paragraph Document::add_header_paragraph(std::size_t section_index, const std::string& text,
                                         ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph =
      add_section_header_paragraph(package_, *xml_, section_index, text, alignment, type);
  dirty_ = true;
  return paragraph;
}

Paragraph Document::add_header_paragraph(std::size_t section_index, const std::vector<Run>& runs,
                                         ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph =
      add_section_header_paragraph(package_, *xml_, section_index, runs, alignment, type);
  dirty_ = true;
  return paragraph;
}

Paragraph Document::add_styled_header_paragraph(std::size_t section_index, const std::string& text,
                                                const RunStyle& style,
                                                ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph = add_section_styled_header_paragraph(package_, *xml_, section_index, text,
                                                            style, alignment, type);
  dirty_ = true;
  return paragraph;
}

Paragraph Document::add_footer_paragraph(std::size_t section_index, const std::string& text,
                                         ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph =
      add_section_footer_paragraph(package_, *xml_, section_index, text, alignment, type);
  dirty_ = true;
  return paragraph;
}

Paragraph Document::add_footer_paragraph(std::size_t section_index, const std::vector<Run>& runs,
                                         ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph =
      add_section_footer_paragraph(package_, *xml_, section_index, runs, alignment, type);
  dirty_ = true;
  return paragraph;
}

Paragraph Document::add_styled_footer_paragraph(std::size_t section_index, const std::string& text,
                                                const RunStyle& style,
                                                ParagraphAlignment alignment, HeaderFooterType type) {
  Paragraph paragraph = add_section_styled_footer_paragraph(package_, *xml_, section_index, text,
                                                            style, alignment, type);
  dirty_ = true;
  return paragraph;
}

void Document::add_picture(const std::filesystem::path& image_path) {
  add_picture(image_path, {});
}

void Document::add_picture(const std::filesystem::path& image_path, const PictureSize& size) {
  const ImageInfo image = load_image_info(image_path);
  const std::int64_t cx_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, true);
  const std::int64_t cy_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, false);
  const std::string rel_id = register_picture_part(package_, image);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  append_picture_run(*xml_, paragraph, rel_id, image, cx_emu, cy_emu);
  dirty_ = true;
}

void Document::add_picture_data(const std::vector<std::uint8_t>& image_bytes,
                                const std::string& extension) {
  add_picture_data(image_bytes, extension, {}, "image");
}

void Document::add_picture_data(const std::vector<std::uint8_t>& image_bytes,
                                const std::string& extension, const PictureSize& size,
                                const std::string& filename_hint) {
  const std::string filename = filename_hint + "." + extension;
  const ImageInfo image = load_image_info(image_bytes, extension, filename);
  const std::int64_t cx_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, true);
  const std::int64_t cy_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, false);
  const std::string rel_id = register_picture_part(package_, image);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  append_picture_run(*xml_, paragraph, rel_id, image, cx_emu, cy_emu);
  dirty_ = true;
}

void Document::add_picture_to_table_cell(std::size_t table_index, std::size_t row_index,
                                         std::size_t col_index,
                                         const std::filesystem::path& image_path) {
  add_picture_to_table_cell(table_index, row_index, col_index, image_path, {});
}

void Document::add_picture_to_table_cell(std::size_t table_index, std::size_t row_index,
                                         std::size_t col_index,
                                         const std::filesystem::path& image_path,
                                         const PictureSize& size) {
  const ImageInfo image = load_image_info(image_path);
  const std::int64_t cx_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, true);
  const std::int64_t cy_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, false);
  const std::string rel_id = register_picture_part(package_, image);

  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node paragraph = cell.append_child("w:p");
  append_picture_run(*xml_, paragraph, rel_id, image, cx_emu, cy_emu);
  dirty_ = true;
}

void Document::add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                              std::size_t col_index,
                                              const std::vector<std::uint8_t>& image_bytes,
                                              const std::string& extension,
                                              const std::string& filename_hint) {
  add_picture_data_to_table_cell(table_index, row_index, col_index, image_bytes, extension, {},
                                 filename_hint);
}

void Document::add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                              std::size_t col_index,
                                              const std::vector<std::uint8_t>& image_bytes,
                                              const std::string& extension, const PictureSize& size,
                                              const std::string& filename_hint) {
  const std::string filename = filename_hint + "." + extension;
  const ImageInfo image = load_image_info(image_bytes, extension, filename);
  const std::int64_t cx_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, true);
  const std::int64_t cy_emu =
      scaled_emu(image.cx_emu, image.cy_emu, size.width_pt, size.height_pt, false);
  const std::string rel_id = register_picture_part(package_, image);

  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  pugi::xml_node paragraph = cell.append_child("w:p");
  append_picture_run(*xml_, paragraph, rel_id, image, cx_emu, cy_emu);
  dirty_ = true;
}

Paragraph Document::add_hyperlink(const std::string& text, const std::string& url,
                                  ParagraphAlignment alignment, const RunStyle& style) {
  const std::string rel_id =
      register_external_hyperlink_relationship(package_, url);
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  apply_alignment(paragraph, alignment);
  append_external_hyperlink(paragraph, rel_id, text, style);
  dirty_ = true;
  return Paragraph(text, alignment, style, false, {Run(text, style)}, {}, -1, {},
                   {HyperlinkInfo{text, url}}, bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Paragraph Document::add_hyperlink_to_table_cell(std::size_t table_index, std::size_t row_index,
                                                std::size_t col_index, const std::string& text,
                                                const std::string& url, const RunStyle& style) {
  const std::string rel_id =
      register_external_hyperlink_relationship(package_, url);
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  const std::size_t cell_paragraph_index = paragraph_index_in_cell(cell);
  pugi::xml_node paragraph = cell.append_child("w:p");
  append_external_hyperlink(paragraph, rel_id, text, style);
  dirty_ = true;
  return Paragraph(text, ParagraphAlignment::Inherit, style, false, {Run(text, style)}, {}, -1,
                   {}, {HyperlinkInfo{text, url}},
                   bind_table_cell_paragraph(*xml_, dirty_, table_index, row_index, col_index,
                                             cell_paragraph_index));
}

std::size_t Document::add_comment(std::size_t paragraph_index, const std::string& text,
                                  const std::string& author, const std::string& initials) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }

  const std::size_t comment_id =
      add_comment_to_paragraph(package_, paragraph, text, author, initials);
  dirty_ = true;
  return comment_id;
}

std::vector<Paragraph> Document::paragraphs() const {
  std::vector<Paragraph> result;
  const RelTargetMap hyperlink_targets = hyperlink_relationship_targets(package_);
  std::size_t paragraph_index = 0;
  for (const pugi::xml_node& child : document_body(*xml_).children()) {
    if (std::string_view(child.name()) != "w:p") {
      continue;
    }
    std::string text;
    const std::string style_id = read_paragraph_style_id_from_xml(child);
    collect_text(child, text);
    result.emplace_back(text, paragraph_alignment_from_xml(child), read_first_run_style_from_xml(child),
                        paragraph_has_page_break(child), read_runs_from_xml(child), style_id,
                        read_heading_level_from_style_id(style_id), read_paragraph_format_from_xml(child),
                        hyperlinks_from_xml(child, hyperlink_targets),
                        bind_body_paragraph(*xml_, dirty_, paragraph_index));
    ++paragraph_index;
  }
  return result;
}

std::vector<Table> Document::tables() const {
  std::vector<Table> result;
  for (const pugi::xml_node& child : document_body(*xml_).children()) {
    if (std::string_view(child.name()) != "w:tbl") {
      continue;
    }
    result.push_back(table_from_xml(child));
  }
  return result;
}

std::vector<Section> Document::sections() const {
  return read_sections_from_xml(*xml_);
}

PageSize Document::page_size_pt() const {
  return read_page_size_pt_from_xml(*xml_);
}

PageOrientation Document::page_orientation() const {
  return read_page_orientation_from_xml(*xml_);
}

PageMargins Document::page_margins_pt() const {
  return read_page_margins_pt_from_xml(*xml_);
}

std::vector<CommentInfo> Document::comments() const {
  return read_comments(package_);
}

std::size_t Document::comment_count() const { return comments().size(); }

std::vector<std::string> Document::headers(HeaderFooterType type) const {
  return read_section_headers(package_, *xml_, type);
}

std::vector<std::string> Document::footers(HeaderFooterType type) const {
  return read_section_footers(package_, *xml_, type);
}

std::vector<Paragraph> Document::header_paragraphs(std::size_t section_index,
                                                   HeaderFooterType type) const {
  return read_header_paragraphs(package_, *xml_, section_index, type);
}

std::vector<Paragraph> Document::footer_paragraphs(std::size_t section_index,
                                                   HeaderFooterType type) const {
  return read_footer_paragraphs(package_, *xml_, section_index, type);
}

std::vector<PictureSize> Document::picture_sizes_pt() const {
  std::vector<PictureSize> result;
  std::vector<pugi::xml_node> stack{*xml_};
  while (!stack.empty()) {
    const pugi::xml_node node = stack.back();
    stack.pop_back();
    if (std::string_view(node.name()) == "wp:inline") {
      result.push_back(picture_size_from_inline(node));
    }
    for (pugi::xml_node child = node.last_child(); child; child = child.previous_sibling()) {
      stack.push_back(child);
    }
  }
  return result;
}

std::vector<PictureInfo> Document::pictures() const {
  std::vector<PictureInfo> result;
  const RelTargetMap rel_targets = image_relationship_targets(package_);

  std::size_t paragraph_index = 0;
  std::size_t table_index = 0;
  for (const pugi::xml_node& child : document_body(*xml_).children()) {
    if (std::string_view(child.name()) == "w:p") {
      collect_pictures_from_paragraph(child, paragraph_index, rel_targets, result);
      ++paragraph_index;
      continue;
    }
    if (std::string_view(child.name()) != "w:tbl") {
      continue;
    }
    std::size_t row_index = 0;
    for (const pugi::xml_node& row : child.children("w:tr")) {
      std::size_t col_index = 0;
      for (const pugi::xml_node& cell : row.children("w:tc")) {
        for (const pugi::xml_node& paragraph : cell.children("w:p")) {
          const std::size_t before = result.size();
          collect_pictures_from_paragraph(paragraph, paragraph_index, rel_targets, result);
          for (std::size_t i = before; i < result.size(); ++i) {
            result[i].in_table_cell = true;
            result[i].table_index = table_index;
            result[i].row_index = row_index;
            result[i].col_index = col_index;
          }
        }
        ++col_index;
      }
      ++row_index;
    }
    ++table_index;
  }

  return result;
}

std::size_t Document::image_count() const {
  return image_count_in_xml(*xml_);
}

void Document::save(const std::filesystem::path& path) {
  if (dirty_) {
    store_document_xml();
  }
  package_.save(path);
}

Document::Document(OpcPackage package) : package_(std::move(package)) { load_document_xml(); }

void Document::load_document_xml() {
  const std::vector<std::uint8_t>& bytes = package_.entry(kDocumentEntry);
  const pugi::xml_parse_result parse_result =
      (xml_ = std::make_unique<pugi::xml_document>())
          ->load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse_result) {
    throw std::runtime_error(std::string("failed to parse word/document.xml: ") +
                             parse_result.description());
  }
  dirty_ = false;
}

void Document::store_document_xml() {
  pugi::xml_document& xml = *xml_;
  package_.set_entry(kDocumentEntry, xml_to_bytes(xml));
  dirty_ = false;
}

}  // namespace docxcpp
