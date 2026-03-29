#include "internal/header_footer_support.hpp"

#include <algorithm>
#include <cctype>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <string_view>

#include "internal/paragraph_support.hpp"
#include "internal/style_support.hpp"

namespace docxcpp {

namespace {

constexpr const char* kDocumentRelationshipsEntry = "word/_rels/document.xml.rels";
constexpr const char* kContentTypesEntry = "[Content_Types].xml";
constexpr const char* kSettingsEntry = "word/settings.xml";
constexpr const char* kHeaderRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
constexpr const char* kFooterRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
constexpr const char* kHeaderContentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
constexpr const char* kFooterContentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

pugi::xml_node child_named_local(const pugi::xml_node& parent, const char* name) {
  for (const pugi::xml_node& child : parent.children()) {
    if (std::string_view(child.name()) == name) {
      return child;
    }
  }
  return {};
}

const char* reference_element_name_local(bool is_header) {
  return is_header ? "w:headerReference" : "w:footerReference";
}

const char* root_element_name_local(bool is_header) { return is_header ? "w:hdr" : "w:ftr"; }

const char* relationship_type_local(bool is_header) {
  return is_header ? kHeaderRelationshipType : kFooterRelationshipType;
}

const char* content_type_local(bool is_header) {
  return is_header ? kHeaderContentType : kFooterContentType;
}

const char* prefix_local(bool is_header) { return is_header ? "header" : "footer"; }

const char* header_footer_type_value_local(HeaderFooterType type) {
  switch (type) {
    case HeaderFooterType::FirstPage:
      return "first";
    case HeaderFooterType::EvenPage:
      return "even";
    case HeaderFooterType::Default:
    default:
      return "default";
  }
}

std::vector<std::uint8_t> xml_to_bytes_local(const pugi::xml_document& xml) {
  std::ostringstream stream;
  xml.save(stream, "  ", pugi::format_default, pugi::encoding_utf8);
  const std::string content = stream.str();
  return std::vector<std::uint8_t>(content.begin(), content.end());
}

pugi::xml_document parse_package_xml_local(const OpcPackage& package, const char* entry_name) {
  pugi::xml_document xml;
  const auto& bytes = package.entry(entry_name);
  const auto result =
      xml.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!result) {
    throw std::runtime_error(std::string("failed to parse package xml: ") + entry_name);
  }
  return xml;
}

std::size_t parse_numeric_suffix_local(std::string_view value, std::string_view prefix,
                                       std::string_view suffix) {
  if (value.size() < prefix.size() + suffix.size()) {
    return 0;
  }
  if (value.substr(0, prefix.size()) != prefix ||
      value.substr(value.size() - suffix.size()) != suffix) {
    return 0;
  }
  std::size_t number = 0;
  for (const char ch : value.substr(prefix.size(), value.size() - prefix.size() - suffix.size())) {
    if (!std::isdigit(static_cast<unsigned char>(ch))) {
      return 0;
    }
    number = number * 10 + static_cast<std::size_t>(ch - '0');
  }
  return number;
}

std::size_t next_relationship_id_local(const pugi::xml_document& rels_xml) {
  std::size_t max_id = 0;
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    const std::string_view value = rel.attribute("Id").value();
    if (value.rfind("rId", 0) != 0) {
      continue;
    }
    std::size_t number = 0;
    for (const char ch : value.substr(3)) {
      if (!std::isdigit(static_cast<unsigned char>(ch))) {
        number = 0;
        break;
      }
      number = number * 10 + static_cast<std::size_t>(ch - '0');
    }
    max_id = std::max(max_id, number);
  }
  return max_id + 1;
}

void ensure_content_type_override_local(pugi::xml_document& content_types_xml,
                                        const std::string& part_name,
                                        const std::string& content_type) {
  pugi::xml_node types = child_named_local(content_types_xml, "Types");
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

std::vector<pugi::xml_node> section_nodes_local(pugi::xml_document& document_xml) {
  std::vector<pugi::xml_node> sections;
  pugi::xml_node document = child_named_local(document_xml, "w:document");
  pugi::xml_node body = child_named_local(document, "w:body");
  if (!body) {
    throw std::runtime_error("required DOCX node is missing: w:body");
  }
  for (const pugi::xml_node& child : body.children()) {
    if (std::string_view(child.name()) == "w:p") {
      if (const pugi::xml_node p_pr = child_named_local(child, "w:pPr")) {
        if (const pugi::xml_node sect_pr = child_named_local(p_pr, "w:sectPr")) {
          sections.push_back(sect_pr);
        }
      }
      continue;
    }
    if (std::string_view(child.name()) == "w:sectPr") {
      sections.push_back(child);
    }
  }
  if (sections.empty()) {
    sections.push_back(body.append_child("w:sectPr"));
  }
  return sections;
}

void clear_header_footer_references_local(pugi::xml_node sect_pr, const char* element_name,
                                          const char* type_value) {
  std::vector<pugi::xml_node> to_remove;
  for (const pugi::xml_node& child : sect_pr.children()) {
    if (std::string_view(child.name()) == element_name &&
        std::string_view(child.attribute("w:type").value()) == type_value) {
      to_remove.push_back(child);
    }
  }
  for (const pugi::xml_node& child : to_remove) {
    sect_pr.remove_child(child);
  }
}

pugi::xml_node header_footer_reference_local(pugi::xml_node sect_pr, const char* element_name,
                                             const char* type_value) {
  for (const pugi::xml_node& child : sect_pr.children()) {
    if (std::string_view(child.name()) == element_name &&
        std::string_view(child.attribute("w:type").value()) == type_value) {
      return child;
    }
  }
  return {};
}

std::string next_part_name_local(const OpcPackage& package, const char* prefix) {
  std::size_t max_id = 0;
  for (const std::string& entry_name : package.entry_names()) {
    max_id = std::max(max_id,
                      parse_numeric_suffix_local(entry_name, std::string("word/") + prefix, ".xml"));
  }
  return std::string(prefix) + std::to_string(max_id + 1) + ".xml";
}

std::vector<std::uint8_t> part_xml_bytes_local(const char* root_name, const std::string& text) {
  pugi::xml_document xml;
  pugi::xml_node root = xml.append_child(root_name);
  root.append_attribute("xmlns:r")
      .set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
  root.append_attribute("xmlns:w")
      .set_value("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
  pugi::xml_node paragraph = root.append_child("w:p");
  pugi::xml_node run = paragraph.append_child("w:r");
  pugi::xml_node text_node = run.append_child("w:t");
  text_node.text().set(text.c_str());
  return xml_to_bytes_local(xml);
}

pugi::xml_document empty_part_xml_local(const char* root_name) {
  pugi::xml_document xml;
  pugi::xml_node root = xml.append_child(root_name);
  root.append_attribute("xmlns:r")
      .set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
  root.append_attribute("xmlns:w")
      .set_value("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
  root.append_child("w:p");
  return xml;
}

ParagraphAlignment paragraph_alignment_from_xml_local(const pugi::xml_node& paragraph) {
  const pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  if (!p_pr) {
    return ParagraphAlignment::Inherit;
  }
  const pugi::xml_node jc = child_named_local(p_pr, "w:jc");
  if (!jc) {
    return ParagraphAlignment::Inherit;
  }
  const std::string_view value = jc.attribute("w:val").value();
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

void apply_alignment_local(pugi::xml_node paragraph, ParagraphAlignment alignment) {
  pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  if (alignment == ParagraphAlignment::Inherit) {
    if (!p_pr) {
      return;
    }
    if (pugi::xml_node jc = child_named_local(p_pr, "w:jc")) {
      p_pr.remove_child(jc);
    }
    return;
  }
  if (!p_pr) {
    p_pr = paragraph.prepend_child("w:pPr");
  }
  pugi::xml_node jc = child_named_local(p_pr, "w:jc");
  if (!jc) {
    jc = p_pr.append_child("w:jc");
  }
  const char* value = "left";
  switch (alignment) {
    case ParagraphAlignment::Center:
      value = "center";
      break;
    case ParagraphAlignment::Right:
      value = "right";
      break;
    case ParagraphAlignment::Justify:
      value = "both";
      break;
    case ParagraphAlignment::Left:
    case ParagraphAlignment::Inherit:
    default:
      value = "left";
      break;
  }
  if (jc.attribute("w:val")) {
    jc.attribute("w:val").set_value(value);
  } else {
    jc.append_attribute("w:val").set_value(value);
  }
}

void collect_text_from_node_local(const pugi::xml_node& node, std::string& text) {
  const std::string_view name = node.name();
  if (name == "w:t") {
    text += node.text().get();
  } else if (name == "w:tab") {
    text.push_back('\t');
  } else if (name == "w:br" || name == "w:cr") {
    text.push_back('\n');
  }
  for (const pugi::xml_node& child : node.children()) {
    collect_text_from_node_local(child, text);
  }
}

std::string collect_text_local(const pugi::xml_node& root) {
  std::string text;
  bool first_paragraph = true;
  for (const pugi::xml_node& paragraph : root.children("w:p")) {
    if (!first_paragraph) {
      text.push_back('\n');
    }
    first_paragraph = false;
    collect_text_from_node_local(paragraph, text);
  }
  return text;
}

std::string ensure_part_and_reference_local(OpcPackage& package, pugi::xml_node sect_pr,
                                            const char* element_name, const char* root_name,
                                            const char* relationship_type,
                                            const char* content_type, const char* prefix,
                                            const char* type_value, const std::string* text) {
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");

  std::string rel_id;
  std::string target_name;
  if (pugi::xml_node ref = header_footer_reference_local(sect_pr, element_name, type_value)) {
    rel_id = ref.attribute("r:id").value();
    for (const pugi::xml_node& rel : relationships.children("Relationship")) {
      if (std::string_view(rel.attribute("Id").value()) == rel_id) {
        target_name = rel.attribute("Target").value();
        break;
      }
    }
  }

  if (rel_id.empty() || target_name.empty()) {
    rel_id = "rId" + std::to_string(next_relationship_id_local(rels_xml));
    target_name = next_part_name_local(package, prefix);
    pugi::xml_node rel = relationships.append_child("Relationship");
    rel.append_attribute("Id").set_value(rel_id.c_str());
    rel.append_attribute("Type").set_value(relationship_type);
    rel.append_attribute("Target").set_value(target_name.c_str());
    clear_header_footer_references_local(sect_pr, element_name, type_value);
    pugi::xml_node ref = sect_pr.prepend_child(element_name);
    ref.append_attribute("w:type").set_value(type_value);
    ref.append_attribute("r:id").set_value(rel_id.c_str());
  }

  package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes_local(rels_xml));

  pugi::xml_document content_types_xml = parse_package_xml_local(package, kContentTypesEntry);
  ensure_content_type_override_local(content_types_xml, "/word/" + target_name, content_type);
  package.set_entry(kContentTypesEntry, xml_to_bytes_local(content_types_xml));

  if (text != nullptr) {
    package.set_entry("word/" + target_name, part_xml_bytes_local(root_name, *text));
  } else if (!package.has_entry("word/" + target_name)) {
    pugi::xml_document empty_xml = empty_part_xml_local(root_name);
    package.set_entry("word/" + target_name, xml_to_bytes_local(empty_xml));
  }
  return rel_id;
}

std::string part_target_name_local(OpcPackage& package, pugi::xml_node sect_pr, const char* element_name,
                                   const char* root_name, const char* relationship_type,
                                   const char* content_type, const char* prefix,
                                   const char* type_value) {
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  if (pugi::xml_node ref = header_footer_reference_local(sect_pr, element_name, type_value)) {
    const std::string rel_id = ref.attribute("r:id").value();
    for (const pugi::xml_node& rel : relationships.children("Relationship")) {
      if (std::string_view(rel.attribute("Id").value()) == rel_id) {
        return rel.attribute("Target").value();
      }
    }
  }
  ensure_part_and_reference_local(package, sect_pr, element_name, root_name, relationship_type,
                                  content_type, prefix, type_value, nullptr);
  rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  relationships = child_named_local(rels_xml, "Relationships");
  const pugi::xml_node ref = header_footer_reference_local(sect_pr, element_name, type_value);
  const std::string rel_id = ref.attribute("r:id").value();
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Id").value()) == rel_id) {
      return rel.attribute("Target").value();
    }
  }
  throw std::runtime_error("failed to resolve header/footer target");
}

void clear_part_content_local(OpcPackage& package, const std::string& entry_name, const char* root_name) {
  pugi::xml_document xml = empty_part_xml_local(root_name);
  package.set_entry(entry_name, xml_to_bytes_local(xml));
}

Paragraph append_part_paragraph_local(OpcPackage& package, const std::string& entry_name,
                                      const char* root_name, const std::string& text,
                                      ParagraphAlignment alignment) {
  pugi::xml_document xml;
  if (package.has_entry(entry_name)) {
    xml = parse_package_xml_local(package, entry_name.c_str());
  } else {
    xml = empty_part_xml_local(root_name);
  }
  pugi::xml_node root = xml.document_element();
  if (!root) {
    xml = empty_part_xml_local(root_name);
    root = xml.document_element();
  }
  if (root.first_child() && !root.first_child().next_sibling()) {
    const pugi::xml_node only_paragraph = root.first_child();
    std::string existing_text = collect_text_local(root);
    if (std::string_view(only_paragraph.name()) == "w:p" && existing_text.empty()) {
      root.remove_child(only_paragraph);
    }
  }
  pugi::xml_node paragraph = root.append_child("w:p");
  apply_alignment_local(paragraph, alignment);
  set_paragraph_text_in_xml(paragraph, text);
  package.set_entry(entry_name, xml_to_bytes_local(xml));
  return Paragraph(text, alignment, {}, false, {Run(text, {})}, {}, -1, {}, {});
}

Paragraph append_part_runs_local(OpcPackage& package, const std::string& entry_name,
                                 const char* root_name, const std::vector<Run>& runs,
                                 ParagraphAlignment alignment) {
  pugi::xml_document xml;
  if (package.has_entry(entry_name)) {
    xml = parse_package_xml_local(package, entry_name.c_str());
  } else {
    xml = empty_part_xml_local(root_name);
  }
  pugi::xml_node root = xml.document_element();
  if (!root) {
    xml = empty_part_xml_local(root_name);
    root = xml.document_element();
  }
  if (root.first_child() && !root.first_child().next_sibling()) {
    const pugi::xml_node only_paragraph = root.first_child();
    std::string existing_text = collect_text_local(root);
    if (std::string_view(only_paragraph.name()) == "w:p" && existing_text.empty()) {
      root.remove_child(only_paragraph);
    }
  }
  pugi::xml_node paragraph = root.append_child("w:p");
  apply_alignment_local(paragraph, alignment);
  const StyleCatalog catalog = load_style_catalog(package);
  set_paragraph_runs_in_xml(paragraph, runs, &catalog);
  package.set_entry(entry_name, xml_to_bytes_local(xml));
  Paragraph result = paragraph_from_runs(runs, alignment);
  return Paragraph(result.text(), result.alignment(), result.first_run_style(), result.has_page_break(),
                   result.runs(), result.style_id(), result.heading_level(), result.format(), {});
}

void set_on_off_node_local(pugi::xml_node parent, const char* name, bool enabled) {
  pugi::xml_node node = child_named_local(parent, name);
  if (enabled) {
    if (!node) {
      node = parent.append_child(name);
    }
    if (node.attribute("w:val")) {
      node.remove_attribute("w:val");
    }
  } else if (node) {
    parent.remove_child(node);
  }
}

std::vector<std::string> read_section_part_texts_local(const OpcPackage& package,
                                                       const pugi::xml_document& document_xml,
                                                       const char* element_name,
                                                       const char* type_value) {
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  std::unordered_map<std::string, std::string> targets;
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    targets.emplace(rel.attribute("Id").value(), rel.attribute("Target").value());
  }

  pugi::xml_document mutable_copy;
  mutable_copy.reset(document_xml);
  std::vector<pugi::xml_node> sections = section_nodes_local(mutable_copy);
  std::vector<std::string> result;
  result.reserve(sections.size());
  for (const pugi::xml_node& sect_pr : sections) {
    const pugi::xml_node ref = header_footer_reference_local(sect_pr, element_name, type_value);
    if (!ref) {
      result.push_back("");
      continue;
    }
    const auto it = targets.find(ref.attribute("r:id").value());
    if (it == targets.end() || !package.has_entry("word/" + it->second)) {
      result.push_back("");
      continue;
    }
    pugi::xml_document part_xml = parse_package_xml_local(package, ("word/" + it->second).c_str());
    result.push_back(collect_text_local(part_xml.document_element()));
  }
  return result;
}

}  // namespace

void set_even_and_odd_headers_in_settings(OpcPackage& package, bool enabled) {
  pugi::xml_document settings_xml = parse_package_xml_local(package, kSettingsEntry);
  pugi::xml_node settings = child_named_local(settings_xml, "w:settings");
  if (!settings) {
    throw std::runtime_error("required DOCX node is missing: w:settings");
  }
  set_on_off_node_local(settings, "w:evenAndOddHeaders", enabled);
  package.set_entry(kSettingsEntry, xml_to_bytes_local(settings_xml));
}

bool read_even_and_odd_headers_from_settings(const OpcPackage& package) {
  pugi::xml_document settings_xml = parse_package_xml_local(package, kSettingsEntry);
  pugi::xml_node settings = child_named_local(settings_xml, "w:settings");
  if (!settings) {
    throw std::runtime_error("required DOCX node is missing: w:settings");
  }
  return static_cast<bool>(child_named_local(settings, "w:evenAndOddHeaders"));
}

void set_section_different_first_page_in_xml(pugi::xml_document& document_xml,
                                             std::size_t section_index, bool enabled) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  set_on_off_node_local(sections[section_index], "w:titlePg", enabled);
}

bool read_section_different_first_page_from_xml(const pugi::xml_document& document_xml,
                                                std::size_t section_index) {
  pugi::xml_document mutable_copy;
  mutable_copy.reset(document_xml);
  std::vector<pugi::xml_node> sections = section_nodes_local(mutable_copy);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  return static_cast<bool>(child_named_local(sections[section_index], "w:titlePg"));
}

void set_section_header_text(OpcPackage& package, pugi::xml_document& document_xml,
                             std::size_t section_index, const std::string& text,
                             HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  ensure_part_and_reference_local(package, sections[section_index], reference_element_name_local(true),
                                  root_element_name_local(true), relationship_type_local(true),
                                  content_type_local(true), prefix_local(true),
                                  header_footer_type_value_local(type), &text);
}

void set_section_footer_text(OpcPackage& package, pugi::xml_document& document_xml,
                             std::size_t section_index, const std::string& text,
                             HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  ensure_part_and_reference_local(package, sections[section_index], reference_element_name_local(false),
                                  root_element_name_local(false), relationship_type_local(false),
                                  content_type_local(false), prefix_local(false),
                                  header_footer_type_value_local(type), &text);
}

void clear_section_header(OpcPackage& package, pugi::xml_document& document_xml,
                          std::size_t section_index, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(true), root_element_name_local(true),
      relationship_type_local(true), content_type_local(true), prefix_local(true),
      header_footer_type_value_local(type));
  clear_part_content_local(package, "word/" + target, root_element_name_local(true));
}

void clear_section_footer(OpcPackage& package, pugi::xml_document& document_xml,
                          std::size_t section_index, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(false), root_element_name_local(false),
      relationship_type_local(false), content_type_local(false), prefix_local(false),
      header_footer_type_value_local(type));
  clear_part_content_local(package, "word/" + target, root_element_name_local(false));
}

Paragraph add_section_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::string& text,
                                       ParagraphAlignment alignment, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(true), root_element_name_local(true),
      relationship_type_local(true), content_type_local(true), prefix_local(true),
      header_footer_type_value_local(type));
  return append_part_paragraph_local(package, "word/" + target, root_element_name_local(true), text,
                                     alignment);
}

Paragraph add_section_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::vector<Run>& runs,
                                       ParagraphAlignment alignment, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(true), root_element_name_local(true),
      relationship_type_local(true), content_type_local(true), prefix_local(true),
      header_footer_type_value_local(type));
  return append_part_runs_local(package, "word/" + target, root_element_name_local(true), runs,
                                alignment);
}

Paragraph add_section_styled_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                              std::size_t section_index, const std::string& text,
                                              const RunStyle& style, ParagraphAlignment alignment,
                                              HeaderFooterType type) {
  return add_section_header_paragraph(package, document_xml, section_index,
                                      std::vector<Run>{Run(text, style)}, alignment, type);
}

Paragraph add_section_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::string& text,
                                       ParagraphAlignment alignment, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(false), root_element_name_local(false),
      relationship_type_local(false), content_type_local(false), prefix_local(false),
      header_footer_type_value_local(type));
  return append_part_paragraph_local(package, "word/" + target, root_element_name_local(false), text,
                                     alignment);
}

Paragraph add_section_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::vector<Run>& runs,
                                       ParagraphAlignment alignment, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections = section_nodes_local(document_xml);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const std::string target = part_target_name_local(
      package, sections[section_index], reference_element_name_local(false), root_element_name_local(false),
      relationship_type_local(false), content_type_local(false), prefix_local(false),
      header_footer_type_value_local(type));
  return append_part_runs_local(package, "word/" + target, root_element_name_local(false), runs,
                                alignment);
}

Paragraph add_section_styled_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                              std::size_t section_index, const std::string& text,
                                              const RunStyle& style, ParagraphAlignment alignment,
                                              HeaderFooterType type) {
  return add_section_footer_paragraph(package, document_xml, section_index,
                                      std::vector<Run>{Run(text, style)}, alignment, type);
}

std::vector<std::string> read_section_headers(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              HeaderFooterType type) {
  return read_section_part_texts_local(package, document_xml, reference_element_name_local(true),
                                       header_footer_type_value_local(type));
}

std::vector<std::string> read_section_footers(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              HeaderFooterType type) {
  return read_section_part_texts_local(package, document_xml, reference_element_name_local(false),
                                       header_footer_type_value_local(type));
}

std::vector<Paragraph> read_header_paragraphs(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              std::size_t section_index, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections;
  pugi::xml_document mutable_copy;
  mutable_copy.reset(document_xml);
  sections = section_nodes_local(mutable_copy);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const pugi::xml_node ref =
      header_footer_reference_local(sections[section_index], reference_element_name_local(true),
                                    header_footer_type_value_local(type));
  if (!ref) {
    return {};
  }
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  std::string target;
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Id").value()) == ref.attribute("r:id").value()) {
      target = rel.attribute("Target").value();
      break;
    }
  }
  if (target.empty()) {
    return {};
  }
  pugi::xml_document part_xml = parse_package_xml_local(package, ("word/" + target).c_str());
  const StyleCatalog catalog = load_style_catalog(package);
  std::vector<Paragraph> result;
  for (const pugi::xml_node& paragraph : part_xml.document_element().children("w:p")) {
    const std::string style_id = read_paragraph_style_id_from_xml(paragraph);
    std::string text;
    collect_text_from_node_local(paragraph, text);
    const auto runs = read_runs_from_xml(paragraph, &catalog);
    result.emplace_back(text, paragraph_alignment_from_xml_local(paragraph),
                        read_first_run_style_from_xml(paragraph, &catalog), false, runs, style_id,
                        read_heading_level_from_style_id(style_id),
                        read_paragraph_format_from_xml(paragraph), std::vector<HyperlinkInfo>{}, nullptr);
  }
  return result;
}

std::vector<Paragraph> read_footer_paragraphs(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              std::size_t section_index, HeaderFooterType type) {
  std::vector<pugi::xml_node> sections;
  pugi::xml_document mutable_copy;
  mutable_copy.reset(document_xml);
  sections = section_nodes_local(mutable_copy);
  if (section_index >= sections.size()) {
    throw std::out_of_range("section_index out of range");
  }
  const pugi::xml_node ref =
      header_footer_reference_local(sections[section_index], reference_element_name_local(false),
                                    header_footer_type_value_local(type));
  if (!ref) {
    return {};
  }
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  std::string target;
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Id").value()) == ref.attribute("r:id").value()) {
      target = rel.attribute("Target").value();
      break;
    }
  }
  if (target.empty()) {
    return {};
  }
  pugi::xml_document part_xml = parse_package_xml_local(package, ("word/" + target).c_str());
  const StyleCatalog catalog = load_style_catalog(package);
  std::vector<Paragraph> result;
  for (const pugi::xml_node& paragraph : part_xml.document_element().children("w:p")) {
    const std::string style_id = read_paragraph_style_id_from_xml(paragraph);
    std::string text;
    collect_text_from_node_local(paragraph, text);
    const auto runs = read_runs_from_xml(paragraph, &catalog);
    result.emplace_back(text, paragraph_alignment_from_xml_local(paragraph),
                        read_first_run_style_from_xml(paragraph, &catalog), false, runs, style_id,
                        read_heading_level_from_style_id(style_id),
                        read_paragraph_format_from_xml(paragraph), std::vector<HyperlinkInfo>{}, nullptr);
  }
  return result;
}

}  // namespace docxcpp
