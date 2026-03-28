#include "internal/hyperlink_comment_support.hpp"

#include <algorithm>
#include <cctype>
#include <sstream>
#include <stdexcept>
#include <string_view>
#include <utility>
#include <vector>

#include "internal/document_model_support.hpp"

namespace docxcpp {

namespace {

constexpr const char* kDocumentRelationshipsEntry = "word/_rels/document.xml.rels";
constexpr const char* kContentTypesEntry = "[Content_Types].xml";
constexpr const char* kCommentsEntry = "word/comments.xml";
constexpr const char* kHyperlinkRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
constexpr const char* kCommentsRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
constexpr const char* kCommentsContentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

pugi::xml_node child_named_local(const pugi::xml_node& parent, const char* name) {
  for (const pugi::xml_node& child : parent.children()) {
    if (std::string_view(child.name()) == name) {
      return child;
    }
  }
  return {};
}

void collect_text_local(const pugi::xml_node& node, std::string& text) {
  const std::string_view name = node.name();
  if (name == "w:t") {
    text += node.text().get();
  } else if (name == "w:tab") {
    text.push_back('\t');
  } else if (name == "w:br" || name == "w:cr") {
    text.push_back('\n');
  }
  for (const pugi::xml_node& child : node.children()) {
    collect_text_local(child, text);
  }
}

bool needs_preserve_space_local(const std::string& text) {
  if (text.empty()) {
    return false;
  }
  if (text.front() == ' ' || text.back() == ' ') {
    return true;
  }
  return text.find("  ") != std::string::npos;
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

std::size_t parse_numeric_suffix_local(std::string_view value, std::string_view prefix) {
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

std::size_t next_relationship_id_local(const pugi::xml_document& rels_xml) {
  std::size_t max_id = 0;
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    max_id = std::max(max_id, parse_numeric_suffix_local(rel.attribute("Id").value(), "rId"));
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

const char* default_comments_xml_local() {
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

void ensure_comments_part_local(OpcPackage& package) {
  if (!package.has_entry(kCommentsEntry)) {
    const std::string xml = default_comments_xml_local();
    package.set_entry(kCommentsEntry, std::vector<std::uint8_t>(xml.begin(), xml.end()));
  }

  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  bool has_comments_rel = false;
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) == kCommentsRelationshipType) {
      has_comments_rel = true;
      break;
    }
  }
  if (!has_comments_rel) {
    const std::string rel_id = "rId" + std::to_string(next_relationship_id_local(rels_xml));
    pugi::xml_node rel = relationships.append_child("Relationship");
    rel.append_attribute("Id").set_value(rel_id.c_str());
    rel.append_attribute("Type").set_value(kCommentsRelationshipType);
    rel.append_attribute("Target").set_value("comments.xml");
    package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes_local(rels_xml));
  }

  pugi::xml_document content_types_xml = parse_package_xml_local(package, kContentTypesEntry);
  ensure_content_type_override_local(content_types_xml, "/word/comments.xml", kCommentsContentType);
  package.set_entry(kContentTypesEntry, xml_to_bytes_local(content_types_xml));
}

std::size_t next_comment_id_local(const pugi::xml_document& comments_xml) {
  std::size_t max_id = 0;
  const pugi::xml_node comments = child_named_local(comments_xml, "w:comments");
  for (const pugi::xml_node& comment : comments.children("w:comment")) {
    max_id = std::max(max_id, static_cast<std::size_t>(comment.attribute("w:id").as_uint()));
  }
  return comments.first_child() ? max_id + 1 : 0;
}

void append_comment_markers_local(pugi::xml_node paragraph, std::size_t comment_id) {
  pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  pugi::xml_node start = p_pr ? paragraph.insert_child_after("w:commentRangeStart", p_pr)
                              : paragraph.prepend_child("w:commentRangeStart");
  start.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());

  pugi::xml_node end = paragraph.append_child("w:commentRangeEnd");
  end.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());

  pugi::xml_node ref_run = paragraph.append_child("w:r");
  pugi::xml_node ref = ref_run.append_child("w:commentReference");
  ref.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());
}

void append_text_run_local(pugi::xml_node paragraph, const std::string& text) {
  pugi::xml_node run = paragraph.append_child("w:r");
  std::size_t start = 0;
  bool wrote_text = false;
  while (start <= text.size()) {
    const std::size_t newline = text.find('\n', start);
    const std::string segment =
        newline == std::string::npos ? text.substr(start) : text.substr(start, newline - start);

    if (!segment.empty() || !wrote_text) {
      pugi::xml_node text_node = run.append_child("w:t");
      text_node.text().set(segment.c_str());
      if (needs_preserve_space_local(segment)) {
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

}  // namespace

RelTargetMap hyperlink_relationship_targets(const OpcPackage& package) {
  RelTargetMap targets;
  const pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  const pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) != kHyperlinkRelationshipType) {
      continue;
    }
    targets.emplace(rel.attribute("Id").value(), rel.attribute("Target").value());
  }
  return targets;
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
    collect_text_local(child, info.text);
    result.push_back(std::move(info));
  }
  return result;
}

std::string register_external_hyperlink_relationship(OpcPackage& package, const std::string& target) {
  pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) == kHyperlinkRelationshipType &&
        std::string_view(rel.attribute("Target").value()) == target &&
        std::string_view(rel.attribute("TargetMode").value()) == "External") {
      return rel.attribute("Id").value();
    }
  }
  const std::string rel_id = "rId" + std::to_string(next_relationship_id_local(rels_xml));
  pugi::xml_node rel = relationships.append_child("Relationship");
  rel.append_attribute("Id").set_value(rel_id.c_str());
  rel.append_attribute("Type").set_value(kHyperlinkRelationshipType);
  rel.append_attribute("Target").set_value(target.c_str());
  rel.append_attribute("TargetMode").set_value("External");
  package.set_entry(kDocumentRelationshipsEntry, xml_to_bytes_local(rels_xml));
  return rel_id;
}

void append_external_hyperlink(pugi::xml_node paragraph, const std::string& rel_id,
                               const std::string& text, const RunStyle& style) {
  pugi::xml_node hyperlink = paragraph.append_child("w:hyperlink");
  hyperlink.append_attribute("r:id").set_value(rel_id.c_str());
  hyperlink.append_attribute("w:history").set_value("1");
  pugi::xml_node run = hyperlink.append_child("w:r");
  pugi::xml_node properties = child_named_local(run, "w:rPr");
  if (!properties) {
    properties = run.prepend_child("w:rPr");
  }
  if (!child_named_local(properties, "w:rStyle")) {
    pugi::xml_node style_node = properties.append_child("w:rStyle");
    style_node.append_attribute("w:val").set_value("Hyperlink");
  }
  apply_run_style_for_model(run, style);
  append_run_text_for_model(run, text);
}

std::size_t add_comment_to_paragraph(OpcPackage& package, pugi::xml_node paragraph,
                                     const std::string& text, const std::string& author,
                                     const std::string& initials) {
  ensure_comments_part_local(package);
  pugi::xml_document comments_xml = parse_package_xml_local(package, kCommentsEntry);
  pugi::xml_node comments = child_named_local(comments_xml, "w:comments");
  const std::size_t comment_id = next_comment_id_local(comments_xml);
  pugi::xml_node comment = comments.append_child("w:comment");
  comment.append_attribute("w:id").set_value(std::to_string(comment_id).c_str());
  comment.append_attribute("w:author").set_value(author.c_str());
  comment.append_attribute("w:initials").set_value(initials.c_str());
  pugi::xml_node comment_paragraph = comment.append_child("w:p");
  append_text_run_local(comment_paragraph, text);
  package.set_entry(kCommentsEntry, xml_to_bytes_local(comments_xml));

  append_comment_markers_local(paragraph, comment_id);
  return comment_id;
}

std::vector<CommentInfo> read_comments(const OpcPackage& package) {
  std::vector<CommentInfo> result;
  if (!package.has_entry(kCommentsEntry)) {
    return result;
  }
  const pugi::xml_document comments_xml = parse_package_xml_local(package, kCommentsEntry);
  const pugi::xml_node comments = child_named_local(comments_xml, "w:comments");
  for (const pugi::xml_node& comment : comments.children("w:comment")) {
    CommentInfo info;
    info.id = static_cast<std::size_t>(comment.attribute("w:id").as_uint());
    info.author = comment.attribute("w:author").value();
    info.initials = comment.attribute("w:initials").value();
    for (const pugi::xml_node& paragraph : comment.children("w:p")) {
      std::string paragraph_text;
      collect_text_local(paragraph, paragraph_text);
      if (!info.text.empty()) {
        info.text.push_back('\n');
      }
      info.text += paragraph_text;
    }
    result.push_back(std::move(info));
  }
  return result;
}

}  // namespace docxcpp
