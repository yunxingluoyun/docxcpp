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

namespace docxcpp {

struct ParagraphBinding {
  pugi::xml_document* xml{nullptr};
  bool* dirty{nullptr};
  int kind{0};
  std::size_t paragraph_index{0};
  std::size_t table_index{0};
  std::size_t row_index{0};
  std::size_t col_index{0};
  std::size_t cell_paragraph_index{0};
};

namespace {

constexpr const char* kDocumentEntry = "word/document.xml";
constexpr const char* kDocumentRelationshipsEntry = "word/_rels/document.xml.rels";
constexpr const char* kContentTypesEntry = "[Content_Types].xml";
constexpr const char* kCommentsEntry = "word/comments.xml";
constexpr const char* kTableStyleValue = "TableGrid";
constexpr const char* kImageRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kHyperlinkRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
constexpr const char* kCommentsRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
constexpr const char* kCommentsContentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
constexpr const char* kPictureUri =
    "http://schemas.openxmlformats.org/drawingml/2006/picture";

struct ImageInfo {
  std::string extension;
  std::string content_type;
  std::string filename;
  std::vector<std::uint8_t> bytes;
  std::int64_t cx_emu;
  std::int64_t cy_emu;
};

constexpr std::int64_t kEmuPerPoint = 12700;
constexpr std::int64_t kTwipsPerPoint = 20;

std::runtime_error missing_node(const char* node_name) {
  return std::runtime_error(std::string("required DOCX node is missing: ") + node_name);
}

std::vector<std::uint8_t> read_binary_file(const std::filesystem::path& path) {
  std::ifstream stream(path, std::ios::binary);
  if (!stream) {
    throw std::runtime_error("failed to open file: " + path.string());
  }
  stream.seekg(0, std::ios::end);
  const auto size = static_cast<std::size_t>(stream.tellg());
  stream.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(size);
  if (size > 0) {
    stream.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(size));
  }
  return bytes;
}

std::uint32_t read_be32(const std::vector<std::uint8_t>& bytes, std::size_t offset) {
  return (static_cast<std::uint32_t>(bytes[offset]) << 24U) |
         (static_cast<std::uint32_t>(bytes[offset + 1]) << 16U) |
         (static_cast<std::uint32_t>(bytes[offset + 2]) << 8U) |
         static_cast<std::uint32_t>(bytes[offset + 3]);
}

std::uint16_t read_be16(const std::vector<std::uint8_t>& bytes, std::size_t offset) {
  return static_cast<std::uint16_t>((bytes[offset] << 8U) | bytes[offset + 1]);
}

ImageInfo load_image_info(const std::vector<std::uint8_t>& bytes, std::string extension,
                          std::string filename);

ImageInfo load_image_info(const std::filesystem::path& image_path) {
  const std::vector<std::uint8_t> bytes = read_binary_file(image_path);
  std::string extension = image_path.extension().string();
  std::transform(extension.begin(), extension.end(), extension.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  if (!extension.empty() && extension.front() == '.') {
    extension.erase(extension.begin());
  }
  return load_image_info(bytes, extension, image_path.filename().string());
}

ImageInfo load_image_info(const std::vector<std::uint8_t>& bytes, std::string extension,
                          std::string filename) {
  ImageInfo info;
  info.filename = std::move(filename);
  info.bytes = bytes;
  std::transform(extension.begin(), extension.end(), extension.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  if (!extension.empty() && extension.front() == '.') {
    extension.erase(extension.begin());
  }
  info.extension = extension;

  std::uint32_t width = 0;
  std::uint32_t height = 0;
  double dpi_x = 96.0;
  double dpi_y = 96.0;

  if (extension == "png") {
    if (info.bytes.size() < 24 || read_be32(info.bytes, 12) != 0x49484452U) {
      throw std::runtime_error("unsupported or invalid PNG: " + info.filename);
    }
    info.content_type = "image/png";
    width = read_be32(info.bytes, 16);
    height = read_be32(info.bytes, 20);

    std::size_t offset = 8;
    while (offset + 12 <= info.bytes.size()) {
      const std::uint32_t chunk_length = read_be32(info.bytes, offset);
      const std::uint32_t chunk_type = read_be32(info.bytes, offset + 4);
      if (chunk_type == 0x70485973U && offset + 8 + chunk_length <= info.bytes.size()) {
        const std::uint32_t x_ppm = read_be32(info.bytes, offset + 8);
        const std::uint32_t y_ppm = read_be32(info.bytes, offset + 12);
        const std::uint8_t unit = info.bytes[offset + 16];
        if (unit == 1 && x_ppm > 0 && y_ppm > 0) {
          dpi_x = static_cast<double>(x_ppm) * 0.0254;
          dpi_y = static_cast<double>(y_ppm) * 0.0254;
        }
        break;
      }
      offset += static_cast<std::size_t>(chunk_length) + 12;
    }
  } else if (extension == "jpg" || extension == "jpeg") {
    info.content_type = "image/jpeg";
    std::size_t offset = 2;
    while (offset + 4 < info.bytes.size()) {
      if (info.bytes[offset] != 0xFF) {
        ++offset;
        continue;
      }

      const std::uint8_t marker = info.bytes[offset + 1];
      if (marker == 0xD8 || marker == 0xD9) {
        offset += 2;
        continue;
      }

      const std::uint16_t segment_length = read_be16(info.bytes, offset + 2);
      if (segment_length < 2 || offset + 2 + segment_length > info.bytes.size()) {
        break;
      }

      if (marker == 0xE0 && segment_length >= 14 &&
          std::string(reinterpret_cast<const char*>(&info.bytes[offset + 4]), 5) == "JFIF\0") {
        const std::uint8_t units = info.bytes[offset + 11];
        const std::uint16_t x_density = read_be16(info.bytes, offset + 12);
        const std::uint16_t y_density = read_be16(info.bytes, offset + 14);
        if (x_density > 0 && y_density > 0) {
          if (units == 1) {
            dpi_x = x_density;
            dpi_y = y_density;
          } else if (units == 2) {
            dpi_x = x_density * 2.54;
            dpi_y = y_density * 2.54;
          }
        }
      }

      const bool is_sof =
          (marker >= 0xC0 && marker <= 0xC3) || (marker >= 0xC5 && marker <= 0xC7) ||
          (marker >= 0xC9 && marker <= 0xCB) || (marker >= 0xCD && marker <= 0xCF);
      if (is_sof && segment_length >= 7) {
        height = read_be16(info.bytes, offset + 5);
        width = read_be16(info.bytes, offset + 7);
        break;
      }

      offset += 2 + segment_length;
    }

    if (width == 0 || height == 0) {
      throw std::runtime_error("unsupported or invalid JPEG: " + info.filename);
    }
  } else {
    throw std::runtime_error("only png/jpg/jpeg images are supported right now");
  }

  info.cx_emu = static_cast<std::int64_t>(914400.0 * static_cast<double>(width) / dpi_x);
  info.cy_emu = static_cast<std::int64_t>(914400.0 * static_cast<double>(height) / dpi_y);
  return info;
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

const char* orientation_value(PageOrientation orientation) {
  switch (orientation) {
    case PageOrientation::Landscape:
      return "landscape";
    case PageOrientation::Portrait:
    default:
      return "portrait";
  }
}

PageOrientation orientation_from_value(std::string_view value) {
  if (value == "landscape") {
    return PageOrientation::Landscape;
  }
  return PageOrientation::Portrait;
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

std::optional<int> twips_attr_to_pt(const pugi::xml_node& node, const char* attr_name) {
  const pugi::xml_attribute attr = node.attribute(attr_name);
  if (!attr) {
    return std::nullopt;
  }
  return attr.as_int() / static_cast<int>(kTwipsPerPoint);
}

std::optional<bool> on_off_property_from_xml(const pugi::xml_node& parent, const char* name) {
  const pugi::xml_node node = child_named(parent, name);
  if (!node) {
    return std::nullopt;
  }
  const pugi::xml_attribute val = node.attribute("w:val");
  if (!val) {
    return true;
  }
  const std::string_view value = val.value();
  if (value == "0" || value == "false" || value == "off") {
    return false;
  }
  return true;
}

ParagraphFormat paragraph_format_from_xml(const pugi::xml_node& paragraph) {
  ParagraphFormat format;
  const pugi::xml_node p_pr = child_named(paragraph, "w:pPr");
  if (!p_pr) {
    return format;
  }

  if (const pugi::xml_node ind = child_named(p_pr, "w:ind")) {
    format.left_indent_pt = twips_attr_to_pt(ind, "w:left");
    format.right_indent_pt = twips_attr_to_pt(ind, "w:right");
    if (const pugi::xml_attribute first_line = ind.attribute("w:firstLine")) {
      format.first_line_indent_pt = first_line.as_int() / static_cast<int>(kTwipsPerPoint);
    } else if (const pugi::xml_attribute hanging = ind.attribute("w:hanging")) {
      format.first_line_indent_pt = -(hanging.as_int() / static_cast<int>(kTwipsPerPoint));
    }
  }

  if (const pugi::xml_node spacing = child_named(p_pr, "w:spacing")) {
    format.space_before_pt = twips_attr_to_pt(spacing, "w:before");
    format.space_after_pt = twips_attr_to_pt(spacing, "w:after");
    if (const pugi::xml_attribute line = spacing.attribute("w:line")) {
      format.line_spacing_pt = line.as_int() / static_cast<int>(kTwipsPerPoint);
    }
  }

  format.keep_together = on_off_property_from_xml(p_pr, "w:keepLines");
  format.keep_with_next = on_off_property_from_xml(p_pr, "w:keepNext");
  format.page_break_before = on_off_property_from_xml(p_pr, "w:pageBreakBefore");

  return format;
}

std::string paragraph_style_id_from_xml(const pugi::xml_node& paragraph) {
  const pugi::xml_node p_pr = child_named(paragraph, "w:pPr");
  if (!p_pr) {
    return {};
  }
  const pugi::xml_node p_style = child_named(p_pr, "w:pStyle");
  if (!p_style) {
    return {};
  }
  return p_style.attribute("w:val").value();
}

int heading_level_from_style_id(std::string_view style_id) {
  if (style_id == "Title") {
    return 0;
  }
  if (style_id.size() <= 7 || style_id.substr(0, 7) != "Heading") {
    return -1;
  }
  int level = 0;
  for (const char ch : style_id.substr(7)) {
    if (ch < '0' || ch > '9') {
      return -1;
    }
    level = level * 10 + (ch - '0');
  }
  return level;
}

RunStyle run_style_from_xml(const pugi::xml_node& run);

RunStyle first_run_style_from_xml(const pugi::xml_node& paragraph) {
  for (const pugi::xml_node& child : paragraph.children()) {
    if (std::string_view(child.name()) == "w:r") {
      return run_style_from_xml(child);
    }
    if (std::string_view(child.name()) == "w:hyperlink") {
      for (const pugi::xml_node& run : child.children("w:r")) {
        return run_style_from_xml(run);
      }
    }
  }
  return {};
}

RunStyle run_style_from_xml(const pugi::xml_node& run) {
  const pugi::xml_node r_pr = child_named(run, "w:rPr");
  if (!r_pr) {
    return {};
  }

  RunStyle style;
  style.bold = static_cast<bool>(child_named(r_pr, "w:b"));
  style.italic = static_cast<bool>(child_named(r_pr, "w:i"));
  style.underline = static_cast<bool>(child_named(r_pr, "w:u"));
  if (const pugi::xml_node size = child_named(r_pr, "w:sz")) {
    style.font_size_pt = size.attribute("w:val").as_int() / 2;
  }
  if (const pugi::xml_node color = child_named(r_pr, "w:color")) {
    style.color_hex = color.attribute("w:val").value();
  }
  if (const pugi::xml_node fonts = child_named(r_pr, "w:rFonts")) {
    style.font_name = fonts.attribute("w:ascii").value();
    if (style.font_name.empty()) {
      style.font_name = fonts.attribute("w:hAnsi").value();
    }
  }
  if (const pugi::xml_node highlight = child_named(r_pr, "w:highlight")) {
    style.highlight = highlight.attribute("w:val").value();
  }
  return style;
}

std::vector<Run> runs_from_xml(const pugi::xml_node& paragraph) {
  std::vector<Run> runs;
  for (const pugi::xml_node& child : paragraph.children()) {
    if (std::string_view(child.name()) == "w:r") {
      std::string text;
      collect_text(child, text);
      runs.emplace_back(text, run_style_from_xml(child));
      continue;
    }
    if (std::string_view(child.name()) == "w:hyperlink") {
      for (const pugi::xml_node& run : child.children("w:r")) {
        std::string text;
        collect_text(run, text);
        runs.emplace_back(text, run_style_from_xml(run));
      }
    }
  }
  return runs;
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

pugi::xml_node require_cell_node(pugi::xml_node row, std::size_t col_index) {
  pugi::xml_node cell = nth_child_named(row, "w:tc", col_index);
  if (!cell) {
    throw std::out_of_range("col_index out of range");
  }
  return cell;
}

pugi::xml_node document_sect_pr(pugi::xml_document& xml) {
  pugi::xml_node body = document_body(xml);
  pugi::xml_node sect_pr = child_named(body, "w:sectPr");
  if (!sect_pr) {
    sect_pr = body.append_child("w:sectPr");
  }
  return sect_pr;
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

void apply_indent_attr(pugi::xml_node ind, const char* name, std::optional<int> pt_value) {
  if (!pt_value.has_value()) {
    if (ind.attribute(name)) {
      ind.remove_attribute(name);
    }
    return;
  }
  const auto twips = std::to_string(*pt_value * static_cast<int>(kTwipsPerPoint));
  if (ind.attribute(name)) {
    ind.attribute(name).set_value(twips.c_str());
  } else {
    ind.append_attribute(name).set_value(twips.c_str());
  }
}

void cleanup_if_empty(pugi::xml_node parent, pugi::xml_node child) {
  if (child && !child.first_attribute() && !child.first_child()) {
    parent.remove_child(child);
  }
}

void set_on_off_property(pugi::xml_node parent, const char* name, std::optional<bool> value) {
  pugi::xml_node node = child_named(parent, name);
  if (!value.has_value()) {
    if (node) {
      parent.remove_child(node);
    }
    return;
  }
  if (!node) {
    node = parent.append_child(name);
  }
  if (node.attribute("w:val")) {
    node.remove_attribute("w:val");
  }
  if (!*value) {
    node.append_attribute("w:val").set_value("0");
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

void set_paragraph_text(pugi::xml_node paragraph, const std::string& text,
                        const RunStyle& style = {}) {
  pugi::xml_node existing_properties = child_named(paragraph, "w:pPr");
  pugi::xml_document preserved_properties_doc;
  if (existing_properties) {
    preserved_properties_doc.append_copy(existing_properties);
  }
  clear_children(paragraph);
  if (preserved_properties_doc.first_child()) {
    paragraph.append_copy(preserved_properties_doc.first_child());
  }
  pugi::xml_node run = paragraph.append_child("w:r");
  apply_run_style(run, style);
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

std::string combined_text(const std::vector<Run>& runs) {
  std::string text;
  for (const Run& run : runs) {
    text += run.text();
  }
  return text;
}

Paragraph paragraph_from_runs(const std::vector<Run>& runs, ParagraphAlignment alignment,
                              std::string style_id = {}, int heading_level = -1) {
  RunStyle first_style;
  if (!runs.empty()) {
    first_style = runs.front().style();
  }
  bool has_page_break = false;
  for (const Run& run : runs) {
    if (run.text().find('\f') != std::string::npos) {
      has_page_break = true;
      break;
    }
  }
  return Paragraph(combined_text(runs), alignment, first_style, has_page_break, runs,
                   std::move(style_id), heading_level, {}, {}, {});
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

void set_paragraph_runs(pugi::xml_node paragraph, const std::vector<Run>& runs) {
  pugi::xml_node existing_properties = child_named(paragraph, "w:pPr");
  pugi::xml_document preserved_properties_doc;
  if (existing_properties) {
    preserved_properties_doc.append_copy(existing_properties);
  }
  clear_children(paragraph);
  if (preserved_properties_doc.first_child()) {
    paragraph.append_copy(preserved_properties_doc.first_child());
  }

  if (runs.empty()) {
    pugi::xml_node empty_run = paragraph.append_child("w:r");
    empty_run.append_child("w:t").text().set("");
    return;
  }

  for (const Run& run_value : runs) {
    pugi::xml_node run = paragraph.append_child("w:r");
    apply_run_style(run, run_value.style());
    append_run_text(run, run_value.text());
  }
}

std::string paragraph_style_id_for_level(int level) {
  if (level < 0 || level > 9) {
    throw std::invalid_argument("heading level must be in range 0-9");
  }
  if (level == 0) {
    return "Title";
  }
  return "Heading" + std::to_string(level);
}

Table table_from_xml(const pugi::xml_node& table_node) {
  std::vector<TableRow> rows;
  for (const pugi::xml_node& row_node : table_node.children("w:tr")) {
    std::vector<TableCell> cells;
    for (const pugi::xml_node& cell_node : row_node.children("w:tc")) {
      std::string cell_text;
      for (const pugi::xml_node& child : cell_node.children()) {
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
      std::size_t grid_span = 1;
      std::string vertical_merge;
      if (const pugi::xml_node tc_pr = child_named(cell_node, "w:tcPr")) {
        if (const pugi::xml_node grid_span_node = child_named(tc_pr, "w:gridSpan")) {
          grid_span = static_cast<std::size_t>(grid_span_node.attribute("w:val").as_uint(1));
        }
        if (const pugi::xml_node v_merge_node = child_named(tc_pr, "w:vMerge")) {
          vertical_merge = v_merge_node.attribute("w:val").value();
          if (vertical_merge.empty()) {
            vertical_merge = "continue";
          }
        }
      }
      cells.emplace_back(cell_text, grid_span, vertical_merge);
    }
    rows.emplace_back(std::move(cells));
  }
  return Table(std::move(rows));
}

PictureSize picture_size_from_inline(const pugi::xml_node& inline_node) {
  const pugi::xml_node extent = child_named(inline_node, "wp:extent");
  PictureSize size;
  size.width_pt = static_cast<int>(extent.attribute("cx").as_llong() / kEmuPerPoint);
  size.height_pt = static_cast<int>(extent.attribute("cy").as_llong() / kEmuPerPoint);
  return size;
}

using RelTargetMap = std::unordered_map<std::string, std::string>;

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

RelTargetMap relationship_targets_by_type(const OpcPackage& package, const char* relationship_type) {
  RelTargetMap targets;
  const pugi::xml_document rels_xml = parse_package_xml(package, kDocumentRelationshipsEntry);
  const pugi::xml_node relationships = child_named(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) != relationship_type) {
      continue;
    }
    targets.emplace(rel.attribute("Id").value(), rel.attribute("Target").value());
  }
  return targets;
}

RelTargetMap image_relationship_targets(const OpcPackage& package) {
  return relationship_targets_by_type(package, kImageRelationshipType);
}

std::vector<HyperlinkInfo> hyperlinks_from_xml(const pugi::xml_node& paragraph,
                                               const RelTargetMap& hyperlink_targets) {
  std::vector<HyperlinkInfo> result;
  for (const pugi::xml_node& child : paragraph.children("w:hyperlink")) {
    HyperlinkInfo info;
    const std::string rel_id = child.attribute("r:id").value();
    const auto it = hyperlink_targets.find(rel_id);
    if (it != hyperlink_targets.end()) {
      info.url = it->second;
    }
    collect_text(child, info.text);
    result.push_back(std::move(info));
  }
  return result;
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

void collect_pictures_from_paragraph(const pugi::xml_node& paragraph, std::size_t paragraph_index,
                                     const RelTargetMap& rel_targets,
                                     std::vector<PictureInfo>& pictures) {
  for (const pugi::xml_node& node : paragraph.children()) {
    for (const pugi::xml_node& child : node.children()) {
      if (std::string_view(child.name()) != "w:drawing") {
        continue;
      }
      for (const pugi::xml_node& inline_node : child.children("wp:inline")) {
        PictureInfo info;
        info.size_pt = picture_size_from_inline(inline_node);
        info.paragraph_index = paragraph_index;
        const pugi::xml_node blip =
            child_named(child_named(child_named(child_named(inline_node, "a:graphic"), "a:graphicData"),
                                    "pic:pic"),
                        "pic:blipFill");
        pugi::xml_node embed_holder = child_named(blip, "a:blip");
        info.relationship_id = embed_holder.attribute("r:embed").value();
        const auto it = rel_targets.find(info.relationship_id);
        if (it != rel_targets.end()) {
          info.target = it->second;
        }
        pictures.push_back(std::move(info));
      }
    }
  }
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

std::int64_t scaled_emu(std::int64_t original_primary, std::int64_t original_secondary,
                        int requested_primary_pt, int requested_secondary_pt, bool primary_axis) {
  const std::int64_t requested_primary =
      requested_primary_pt > 0 ? static_cast<std::int64_t>(requested_primary_pt) * kEmuPerPoint : 0;
  const std::int64_t requested_secondary =
      requested_secondary_pt > 0 ? static_cast<std::int64_t>(requested_secondary_pt) * kEmuPerPoint : 0;

  if (requested_primary > 0 && requested_secondary > 0) {
    return primary_axis ? requested_primary : requested_secondary;
  }
  if (requested_primary > 0) {
    if (primary_axis) {
      return requested_primary;
    }
    return static_cast<std::int64_t>(
        (static_cast<long double>(original_secondary) * requested_primary) / original_primary);
  }
  if (requested_secondary > 0) {
    if (!primary_axis) {
      return requested_secondary;
    }
    return static_cast<std::int64_t>(
        (static_cast<long double>(original_primary) * requested_secondary) / original_secondary);
  }
  return primary_axis ? original_primary : original_secondary;
}

void append_picture_run(pugi::xml_document& xml, pugi::xml_node paragraph, const std::string& rel_id,
                        const ImageInfo& image, std::int64_t cx_emu, std::int64_t cy_emu) {
  ensure_picture_namespaces(xml);
  pugi::xml_node run = paragraph.append_child("w:r");
  pugi::xml_node drawing = run.append_child("w:drawing");
  pugi::xml_node inline_node = drawing.append_child("wp:inline");
  inline_node.append_attribute("distT").set_value("0");
  inline_node.append_attribute("distB").set_value("0");
  inline_node.append_attribute("distL").set_value("0");
  inline_node.append_attribute("distR").set_value("0");

  pugi::xml_node extent = inline_node.append_child("wp:extent");
  extent.append_attribute("cx").set_value(std::to_string(cx_emu).c_str());
  extent.append_attribute("cy").set_value(std::to_string(cy_emu).c_str());

  const std::size_t doc_pr_id = next_docpr_id(xml, 0) + 1;
  pugi::xml_node doc_pr = inline_node.append_child("wp:docPr");
  doc_pr.append_attribute("id").set_value(std::to_string(doc_pr_id).c_str());
  doc_pr.append_attribute("name").set_value(("Picture " + std::to_string(doc_pr_id)).c_str());

  pugi::xml_node graphic_frame = inline_node.append_child("wp:cNvGraphicFramePr");
  graphic_frame.append_child("a:graphicFrameLocks")
      .append_attribute("noChangeAspect")
      .set_value("1");

  pugi::xml_node graphic = inline_node.append_child("a:graphic");
  pugi::xml_node graphic_data = graphic.append_child("a:graphicData");
  graphic_data.append_attribute("uri").set_value(kPictureUri);

  pugi::xml_node pic = graphic_data.append_child("pic:pic");
  pugi::xml_node nv_pic_pr = pic.append_child("pic:nvPicPr");
  pugi::xml_node c_nv_pr = nv_pic_pr.append_child("pic:cNvPr");
  c_nv_pr.append_attribute("id").set_value("0");
  c_nv_pr.append_attribute("name").set_value(image.filename.c_str());
  nv_pic_pr.append_child("pic:cNvPicPr");

  pugi::xml_node blip_fill = pic.append_child("pic:blipFill");
  blip_fill.append_child("a:blip").append_attribute("r:embed").set_value(rel_id.c_str());
  pugi::xml_node stretch = blip_fill.append_child("a:stretch");
  stretch.append_child("a:fillRect");

  pugi::xml_node sp_pr = pic.append_child("pic:spPr");
  pugi::xml_node xfrm = sp_pr.append_child("a:xfrm");
  pugi::xml_node off = xfrm.append_child("a:off");
  off.append_attribute("x").set_value("0");
  off.append_attribute("y").set_value("0");
  pugi::xml_node ext = xfrm.append_child("a:ext");
  ext.append_attribute("cx").set_value(std::to_string(cx_emu).c_str());
  ext.append_attribute("cy").set_value(std::to_string(cy_emu).c_str());
  sp_pr.append_child("a:prstGeom").append_attribute("prst").set_value("rect");
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

std::string register_external_relationship(OpcPackage& package, const std::string& relationship_type,
                                           const std::string& target) {
  pugi::xml_document rels_xml = parse_package_xml(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) == relationship_type &&
        std::string_view(rel.attribute("Target").value()) == target &&
        std::string_view(rel.attribute("TargetMode").value()) == "External") {
      return rel.attribute("Id").value();
    }
  }
  const std::string rel_id = "rId" + std::to_string(next_relationship_id(rels_xml));
  pugi::xml_node rel = relationships.append_child("Relationship");
  rel.append_attribute("Id").set_value(rel_id.c_str());
  rel.append_attribute("Type").set_value(relationship_type.c_str());
  rel.append_attribute("Target").set_value(target.c_str());
  rel.append_attribute("TargetMode").set_value("External");
  package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes(rels_xml));
  return rel_id;
}

const char* default_comments_xml() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"/>)";
}

void ensure_comments_part(OpcPackage& package) {
  if (!package.has_entry(kCommentsEntry)) {
    const std::string xml = default_comments_xml();
    package.set_entry(kCommentsEntry, std::vector<std::uint8_t>(xml.begin(), xml.end()));
  }

  pugi::xml_document rels_xml = parse_package_xml(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named(rels_xml, "Relationships");
  bool has_comments_rel = false;
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) == kCommentsRelationshipType) {
      has_comments_rel = true;
      break;
    }
  }
  if (!has_comments_rel) {
    const std::string rel_id = "rId" + std::to_string(next_relationship_id(rels_xml));
    pugi::xml_node rel = relationships.append_child("Relationship");
    rel.append_attribute("Id").set_value(rel_id.c_str());
    rel.append_attribute("Type").set_value(kCommentsRelationshipType);
    rel.append_attribute("Target").set_value("comments.xml");
    package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes(rels_xml));
  }

  pugi::xml_document content_types_xml = parse_package_xml(package, kContentTypesEntry);
  ensure_content_type_override(content_types_xml, "/word/comments.xml", kCommentsContentType);
  package.set_entry(kContentTypesEntry, xml_to_bytes(content_types_xml));
}

std::size_t next_comment_id(const pugi::xml_document& comments_xml) {
  std::size_t max_id = 0;
  const pugi::xml_node comments = child_named(comments_xml, "w:comments");
  for (const pugi::xml_node& comment : comments.children("w:comment")) {
    max_id = std::max(max_id, static_cast<std::size_t>(comment.attribute("w:id").as_uint()));
  }
  return comments.first_child() ? max_id + 1 : 0;
}

void append_comment_markers(pugi::xml_node paragraph, std::size_t comment_id) {
  pugi::xml_node p_pr = child_named(paragraph, "w:pPr");
  pugi::xml_node start = p_pr ? paragraph.insert_child_after("w:commentRangeStart", p_pr)
                              : paragraph.prepend_child("w:commentRangeStart");
  start.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());

  pugi::xml_node end = paragraph.append_child("w:commentRangeEnd");
  end.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());

  pugi::xml_node ref_run = paragraph.append_child("w:r");
  pugi::xml_node ref = ref_run.append_child("w:commentReference");
  ref.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());
}

void append_hyperlink(pugi::xml_node paragraph, const std::string& rel_id, const std::string& text,
                      const RunStyle& style) {
  pugi::xml_node hyperlink = paragraph.append_child("w:hyperlink");
  hyperlink.append_attribute("r:id").set_value(rel_id.c_str());
  hyperlink.append_attribute("w:history").set_value("1");
  pugi::xml_node run = hyperlink.append_child("w:r");
  apply_hyperlink_character_style(run);
  apply_run_style(run, style);
  append_run_text(run, text);
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

Run::Run(std::string text, RunStyle style)
    : text_(std::move(text)), style_(std::move(style)) {}

const std::string& Run::text() const noexcept { return text_; }

const RunStyle& Run::style() const noexcept { return style_; }

Paragraph::Paragraph(std::string text, ParagraphAlignment alignment, RunStyle first_run_style,
                     bool has_page_break, std::vector<Run> runs, std::string style_id,
                     int heading_level, ParagraphFormat format,
                     std::vector<HyperlinkInfo> hyperlinks,
                     std::shared_ptr<ParagraphBinding> binding)
    : text_(std::move(text)),
      alignment_(alignment),
      first_run_style_(std::move(first_run_style)),
      has_page_break_(has_page_break),
      runs_(std::move(runs)),
      style_id_(std::move(style_id)),
      heading_level_(heading_level),
      format_(std::move(format)),
      hyperlinks_(std::move(hyperlinks)),
      binding_(std::move(binding)) {}

const std::string& Paragraph::text() const noexcept { return text_; }

ParagraphAlignment Paragraph::alignment() const noexcept { return alignment_; }

const RunStyle& Paragraph::first_run_style() const noexcept { return first_run_style_; }

bool Paragraph::has_page_break() const noexcept { return has_page_break_; }

const std::vector<Run>& Paragraph::runs() const noexcept { return runs_; }

const std::string& Paragraph::style_id() const noexcept { return style_id_; }

int Paragraph::heading_level() const noexcept { return heading_level_; }

const ParagraphFormat& Paragraph::format() const noexcept { return format_; }

const std::vector<HyperlinkInfo>& Paragraph::hyperlinks() const noexcept { return hyperlinks_; }

Run Paragraph::add_run(const std::string& text, const RunStyle& style) {
  if (!binding_) {
    throw std::logic_error("paragraph is not bound to a writable document");
  }
  pugi::xml_node paragraph = paragraph_node_from_binding(*binding_);
  if (!paragraph) {
    throw std::out_of_range("bound paragraph no longer exists");
  }

  pugi::xml_node run = paragraph.append_child("w:r");
  apply_run_style(run, style);
  append_run_text(run, text);
  if (binding_->dirty != nullptr) {
    *binding_->dirty = true;
  }

  Run added_run(text, style);
  if (runs_.empty()) {
    first_run_style_ = style;
  }
  runs_.push_back(added_run);
  text_ += text;
  return added_run;
}

TableCell::TableCell(std::string text, std::size_t grid_span, std::string vertical_merge)
    : text_(std::move(text)),
      grid_span_(grid_span),
      vertical_merge_(std::move(vertical_merge)) {}

const std::string& TableCell::text() const noexcept { return text_; }

std::size_t TableCell::grid_span() const noexcept { return grid_span_; }

const std::string& TableCell::vertical_merge() const noexcept { return vertical_merge_; }

TableRow::TableRow(std::vector<TableCell> cells) : cells_(std::move(cells)) {}

const std::vector<TableCell>& TableRow::cells() const noexcept { return cells_; }

Table::Table(std::vector<TableRow> rows) : rows_(std::move(rows)) {}

const std::vector<TableRow>& Table::rows() const noexcept { return rows_; }

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
  set_paragraph_runs(paragraph, runs);
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
  set_paragraph_text(paragraph, text, style);
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
  set_paragraph_runs(paragraph, runs);
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
  set_paragraph_text(paragraph, text, run_style);
  dirty_ = true;
  return Paragraph(text, alignment, run_style, false, {Run(text, run_style)}, style_id, level, {},
                   {}, bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Table Document::add_table(std::size_t rows, std::size_t cols) {
  if (rows == 0 || cols == 0) {
    throw std::invalid_argument("table rows and cols must be greater than zero");
  }

  pugi::xml_node table = append_block_before_section(document_body(*xml_), "w:tbl");
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
      pugi::xml_node cell = row.append_child("w:tc");
      pugi::xml_node tc_pr = cell.append_child("w:tcPr");
      tc_pr.append_child("w:tcW").append_attribute("w:w").set_value("2390");
      tc_pr.child("w:tcW").append_attribute("w:type").set_value("dxa");
      pugi::xml_node paragraph = cell.append_child("w:p");
      set_paragraph_text(paragraph, "");
      table_cells.emplace_back("");
    }
    table_rows.emplace_back(std::move(table_cells));
  }

  dirty_ = true;
  return Table(std::move(table_rows));
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
  set_paragraph_text(paragraph, text);
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
  set_paragraph_runs(paragraph, runs);
  dirty_ = true;
}

Paragraph Document::add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                             std::size_t col_index, const std::string& text) {
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  const std::size_t cell_paragraph_index = paragraph_index_in_cell(cell);
  pugi::xml_node paragraph = cell.append_child("w:p");
  set_paragraph_text(paragraph, text);
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
  set_paragraph_runs(paragraph, runs);
  dirty_ = true;
  Paragraph result = paragraph_from_runs(runs, ParagraphAlignment::Inherit);
  return Paragraph(result.text(), result.alignment(), result.first_run_style(), result.has_page_break(),
                   result.runs(), result.style_id(), result.heading_level(), result.format(), {},
                   bind_table_cell_paragraph(*xml_, dirty_, table_index, row_index, col_index,
                                             cell_paragraph_index));
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
        pugi::xml_node extra_cell = require_cell_node(row, col_index);
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

void Document::set_paragraph_alignment(std::size_t paragraph_index, ParagraphAlignment alignment) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }
  apply_alignment(paragraph, alignment);
  dirty_ = true;
}

void Document::set_paragraph_indentation_pt(std::size_t paragraph_index,
                                            std::optional<int> left_indent_pt,
                                            std::optional<int> right_indent_pt,
                                            std::optional<int> first_line_indent_pt) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }
  if ((left_indent_pt && *left_indent_pt < 0) || (right_indent_pt && *right_indent_pt < 0)) {
    throw std::invalid_argument("left/right paragraph indent cannot be negative");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties(paragraph);
  pugi::xml_node ind = child_named(p_pr, "w:ind");
  if (!ind && (left_indent_pt || right_indent_pt || first_line_indent_pt)) {
    ind = p_pr.append_child("w:ind");
  }
  if (ind) {
    apply_indent_attr(ind, "w:left", left_indent_pt);
    apply_indent_attr(ind, "w:right", right_indent_pt);
    if (ind.attribute("w:firstLine")) {
      ind.remove_attribute("w:firstLine");
    }
    if (ind.attribute("w:hanging")) {
      ind.remove_attribute("w:hanging");
    }
    if (first_line_indent_pt.has_value()) {
      const auto twips =
          std::to_string(std::abs(*first_line_indent_pt) * static_cast<int>(kTwipsPerPoint));
      const char* attr_name = *first_line_indent_pt >= 0 ? "w:firstLine" : "w:hanging";
      ind.append_attribute(attr_name).set_value(twips.c_str());
    }
    cleanup_if_empty(p_pr, ind);
    cleanup_if_empty(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_spacing_pt(std::size_t paragraph_index,
                                        std::optional<int> space_before_pt,
                                        std::optional<int> space_after_pt) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }
  if ((space_before_pt && *space_before_pt < 0) || (space_after_pt && *space_after_pt < 0)) {
    throw std::invalid_argument("paragraph spacing cannot be negative");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties(paragraph);
  pugi::xml_node spacing = child_named(p_pr, "w:spacing");
  if (!spacing && (space_before_pt || space_after_pt)) {
    spacing = p_pr.append_child("w:spacing");
  }
  if (spacing) {
    apply_indent_attr(spacing, "w:before", space_before_pt);
    apply_indent_attr(spacing, "w:after", space_after_pt);
    cleanup_if_empty(p_pr, spacing);
    cleanup_if_empty(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_line_spacing_pt(std::size_t paragraph_index,
                                             std::optional<int> line_spacing_pt) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }
  if (line_spacing_pt && *line_spacing_pt <= 0) {
    throw std::invalid_argument("line spacing must be greater than zero");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties(paragraph);
  pugi::xml_node spacing = child_named(p_pr, "w:spacing");
  if (!spacing && line_spacing_pt) {
    spacing = p_pr.append_child("w:spacing");
  }
  if (spacing) {
    if (spacing.attribute("w:line")) {
      spacing.remove_attribute("w:line");
    }
    if (spacing.attribute("w:lineRule")) {
      spacing.remove_attribute("w:lineRule");
    }
    if (line_spacing_pt.has_value()) {
      const auto twips =
          std::to_string(*line_spacing_pt * static_cast<int>(kTwipsPerPoint));
      spacing.append_attribute("w:line").set_value(twips.c_str());
      spacing.append_attribute("w:lineRule").set_value("exact");
    }
    cleanup_if_empty(p_pr, spacing);
    cleanup_if_empty(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_pagination(std::size_t paragraph_index,
                                        std::optional<bool> keep_together,
                                        std::optional<bool> keep_with_next,
                                        std::optional<bool> page_break_before) {
  pugi::xml_node paragraph = nth_child_named(document_body(*xml_), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties(paragraph);
  set_on_off_property(p_pr, "w:keepLines", keep_together);
  set_on_off_property(p_pr, "w:keepNext", keep_with_next);
  set_on_off_property(p_pr, "w:pageBreakBefore", page_break_before);
  cleanup_if_empty(paragraph, p_pr);
  dirty_ = true;
}

void Document::set_page_size_pt(int width_pt, int height_pt) {
  if (width_pt <= 0 || height_pt <= 0) {
    throw std::invalid_argument("page size must be greater than zero");
  }
  pugi::xml_node sect_pr = document_sect_pr(*xml_);
  pugi::xml_node pg_sz = get_or_add_child(sect_pr, "w:pgSz");
  const auto width_twips = std::to_string(width_pt * kTwipsPerPoint);
  const auto height_twips = std::to_string(height_pt * kTwipsPerPoint);
  if (pg_sz.attribute("w:w")) {
    pg_sz.attribute("w:w").set_value(width_twips.c_str());
  } else {
    pg_sz.append_attribute("w:w").set_value(width_twips.c_str());
  }
  if (pg_sz.attribute("w:h")) {
    pg_sz.attribute("w:h").set_value(height_twips.c_str());
  } else {
    pg_sz.append_attribute("w:h").set_value(height_twips.c_str());
  }
  dirty_ = true;
}

void Document::set_page_orientation(PageOrientation orientation) {
  pugi::xml_node sect_pr = document_sect_pr(*xml_);
  pugi::xml_node pg_sz = get_or_add_child(sect_pr, "w:pgSz");
  const char* value = orientation_value(orientation);
  if (pg_sz.attribute("w:orient")) {
    pg_sz.attribute("w:orient").set_value(value);
  } else {
    pg_sz.append_attribute("w:orient").set_value(value);
  }
  dirty_ = true;
}

void Document::set_page_margins_pt(const PageMargins& margins) {
  if (margins.top_pt < 0 || margins.right_pt < 0 || margins.bottom_pt < 0 || margins.left_pt < 0 ||
      margins.header_pt < 0 || margins.footer_pt < 0 || margins.gutter_pt < 0) {
    throw std::invalid_argument("page margins cannot be negative");
  }
  pugi::xml_node sect_pr = document_sect_pr(*xml_);
  pugi::xml_node pg_mar = get_or_add_child(sect_pr, "w:pgMar");

  auto set_attr = [&](const char* name, int pt_value) {
    const auto twips = std::to_string(pt_value * kTwipsPerPoint);
    if (pg_mar.attribute(name)) {
      pg_mar.attribute(name).set_value(twips.c_str());
    } else {
      pg_mar.append_attribute(name).set_value(twips.c_str());
    }
  };

  set_attr("w:top", margins.top_pt);
  set_attr("w:right", margins.right_pt);
  set_attr("w:bottom", margins.bottom_pt);
  set_attr("w:left", margins.left_pt);
  set_attr("w:header", margins.header_pt);
  set_attr("w:footer", margins.footer_pt);
  set_attr("w:gutter", margins.gutter_pt);
  dirty_ = true;
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
      register_external_relationship(package_, kHyperlinkRelationshipType, url);
  const std::size_t paragraph_index = body_paragraph_count(*xml_);
  pugi::xml_node paragraph = append_block_before_section(document_body(*xml_), "w:p");
  apply_alignment(paragraph, alignment);
  append_hyperlink(paragraph, rel_id, text, style);
  dirty_ = true;
  return Paragraph(text, alignment, style, false, {Run(text, style)}, {}, -1, {},
                   {HyperlinkInfo{text, url}}, bind_body_paragraph(*xml_, dirty_, paragraph_index));
}

Paragraph Document::add_hyperlink_to_table_cell(std::size_t table_index, std::size_t row_index,
                                                std::size_t col_index, const std::string& text,
                                                const std::string& url, const RunStyle& style) {
  const std::string rel_id =
      register_external_relationship(package_, kHyperlinkRelationshipType, url);
  pugi::xml_node table = require_table_node(*xml_, table_index);
  pugi::xml_node row = require_row_node(table, row_index);
  pugi::xml_node cell = require_cell_node(row, col_index);
  const std::size_t cell_paragraph_index = paragraph_index_in_cell(cell);
  pugi::xml_node paragraph = cell.append_child("w:p");
  append_hyperlink(paragraph, rel_id, text, style);
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

  ensure_comments_part(package_);
  pugi::xml_document comments_xml = parse_package_xml(package_, kCommentsEntry);
  pugi::xml_node comments = child_named(comments_xml, "w:comments");
  const std::size_t comment_id = next_comment_id(comments_xml);
  pugi::xml_node comment = comments.append_child("w:comment");
  comment.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());
  comment.append_attribute("w:author").set_value(author.c_str());
  comment.append_attribute("w:initials").set_value(initials.c_str());
  pugi::xml_node comment_paragraph = comment.append_child("w:p");
  set_paragraph_text(comment_paragraph, text);
  package_.set_entry(kCommentsEntry, xml_to_bytes(comments_xml));

  append_comment_markers(paragraph, comment_id);
  dirty_ = true;
  return comment_id;
}

std::vector<Paragraph> Document::paragraphs() const {
  std::vector<Paragraph> result;
  const RelTargetMap hyperlink_targets =
      relationship_targets_by_type(package_, kHyperlinkRelationshipType);
  std::size_t paragraph_index = 0;
  for (const pugi::xml_node& child : document_body(*xml_).children()) {
    if (std::string_view(child.name()) != "w:p") {
      continue;
    }
    std::string text;
    const std::string style_id = paragraph_style_id_from_xml(child);
    collect_text(child, text);
    result.emplace_back(text, paragraph_alignment_from_xml(child), first_run_style_from_xml(child),
                        paragraph_has_page_break(child), runs_from_xml(child), style_id,
                        heading_level_from_style_id(style_id), paragraph_format_from_xml(child),
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

PageSize Document::page_size_pt() const {
  PageSize size;
  const pugi::xml_node sect_pr = child_named(document_body(*xml_), "w:sectPr");
  if (!sect_pr) {
    return size;
  }
  const pugi::xml_node pg_sz = child_named(sect_pr, "w:pgSz");
  if (!pg_sz) {
    return size;
  }
  size.width_pt = static_cast<int>(pg_sz.attribute("w:w").as_int() / kTwipsPerPoint);
  size.height_pt = static_cast<int>(pg_sz.attribute("w:h").as_int() / kTwipsPerPoint);
  return size;
}

PageOrientation Document::page_orientation() const {
  const pugi::xml_node sect_pr = child_named(document_body(*xml_), "w:sectPr");
  if (!sect_pr) {
    return PageOrientation::Portrait;
  }
  const pugi::xml_node pg_sz = child_named(sect_pr, "w:pgSz");
  if (!pg_sz) {
    return PageOrientation::Portrait;
  }
  return orientation_from_value(pg_sz.attribute("w:orient").value());
}

PageMargins Document::page_margins_pt() const {
  PageMargins margins;
  const pugi::xml_node sect_pr = child_named(document_body(*xml_), "w:sectPr");
  if (!sect_pr) {
    return margins;
  }
  const pugi::xml_node pg_mar = child_named(sect_pr, "w:pgMar");
  if (!pg_mar) {
    return margins;
  }
  margins.top_pt = pg_mar.attribute("w:top").as_int() / kTwipsPerPoint;
  margins.right_pt = pg_mar.attribute("w:right").as_int() / kTwipsPerPoint;
  margins.bottom_pt = pg_mar.attribute("w:bottom").as_int() / kTwipsPerPoint;
  margins.left_pt = pg_mar.attribute("w:left").as_int() / kTwipsPerPoint;
  margins.header_pt = pg_mar.attribute("w:header").as_int() / kTwipsPerPoint;
  margins.footer_pt = pg_mar.attribute("w:footer").as_int() / kTwipsPerPoint;
  margins.gutter_pt = pg_mar.attribute("w:gutter").as_int() / kTwipsPerPoint;
  return margins;
}

std::vector<CommentInfo> Document::comments() const {
  std::vector<CommentInfo> result;
  if (!package_.has_entry(kCommentsEntry)) {
    return result;
  }
  const pugi::xml_document comments_xml = parse_package_xml(package_, kCommentsEntry);
  const pugi::xml_node comments = child_named(comments_xml, "w:comments");
  for (const pugi::xml_node& comment : comments.children("w:comment")) {
    CommentInfo info;
    info.id = static_cast<std::size_t>(comment.attribute("w:id").as_uint());
    info.author = comment.attribute("w:author").value();
    info.initials = comment.attribute("w:initials").value();
    for (const pugi::xml_node& paragraph : comment.children("w:p")) {
      std::string paragraph_text;
      collect_text(paragraph, paragraph_text);
      if (!info.text.empty()) {
        info.text.push_back('\n');
      }
      info.text += paragraph_text;
    }
    result.push_back(std::move(info));
  }
  return result;
}

std::size_t Document::comment_count() const { return comments().size(); }

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
  return count_named_descendants(*xml_, "a:blip");
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
