#include "internal/paragraph_support.hpp"

#include <optional>
#include <stdexcept>
#include <string_view>
#include <utility>

#include "internal/document_model_support.hpp"

namespace docxcpp {

namespace {

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

void clear_children_local(pugi::xml_node node) {
  while (node.first_child()) {
    node.remove_child(node.first_child());
  }
}

std::optional<int> twips_attr_to_pt_local(const pugi::xml_node& node, const char* attr_name) {
  const pugi::xml_attribute attr = node.attribute(attr_name);
  if (!attr) {
    return std::nullopt;
  }
  return attr.as_int() / 20;
}

std::optional<bool> on_off_property_from_xml_local(const pugi::xml_node& parent, const char* name) {
  const pugi::xml_node node = child_named_local(parent, name);
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

RunStyle run_style_from_xml_local(const pugi::xml_node& run) {
  const pugi::xml_node r_pr = child_named_local(run, "w:rPr");
  if (!r_pr) {
    return {};
  }

  RunStyle style;
  style.bold = static_cast<bool>(child_named_local(r_pr, "w:b"));
  style.italic = static_cast<bool>(child_named_local(r_pr, "w:i"));
  style.underline = static_cast<bool>(child_named_local(r_pr, "w:u"));
  if (const pugi::xml_node size = child_named_local(r_pr, "w:sz")) {
    style.font_size_pt = size.attribute("w:val").as_int() / 2;
  }
  if (const pugi::xml_node color = child_named_local(r_pr, "w:color")) {
    style.color_hex = color.attribute("w:val").value();
  }
  if (const pugi::xml_node fonts = child_named_local(r_pr, "w:rFonts")) {
    style.font_name = fonts.attribute("w:ascii").value();
    if (style.font_name.empty()) {
      style.font_name = fonts.attribute("w:hAnsi").value();
    }
  }
  if (const pugi::xml_node highlight = child_named_local(r_pr, "w:highlight")) {
    style.highlight = highlight.attribute("w:val").value();
  }
  return style;
}

std::string combined_text_local(const std::vector<Run>& runs) {
  std::string text;
  for (const Run& run : runs) {
    text += run.text();
  }
  return text;
}

}  // namespace

ParagraphFormat read_paragraph_format_from_xml(const pugi::xml_node& paragraph) {
  ParagraphFormat format;
  const pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  if (!p_pr) {
    return format;
  }

  if (const pugi::xml_node ind = child_named_local(p_pr, "w:ind")) {
    format.left_indent_pt = twips_attr_to_pt_local(ind, "w:left");
    format.right_indent_pt = twips_attr_to_pt_local(ind, "w:right");
    if (const pugi::xml_attribute first_line = ind.attribute("w:firstLine")) {
      format.first_line_indent_pt = first_line.as_int() / 20;
    } else if (const pugi::xml_attribute hanging = ind.attribute("w:hanging")) {
      format.first_line_indent_pt = -(hanging.as_int() / 20);
    }
  }

  if (const pugi::xml_node spacing = child_named_local(p_pr, "w:spacing")) {
    format.space_before_pt = twips_attr_to_pt_local(spacing, "w:before");
    format.space_after_pt = twips_attr_to_pt_local(spacing, "w:after");
    if (const pugi::xml_attribute line = spacing.attribute("w:line")) {
      format.line_spacing_pt = line.as_int() / 20;
    }
  }

  format.keep_together = on_off_property_from_xml_local(p_pr, "w:keepLines");
  format.keep_with_next = on_off_property_from_xml_local(p_pr, "w:keepNext");
  format.page_break_before = on_off_property_from_xml_local(p_pr, "w:pageBreakBefore");

  return format;
}

std::string read_paragraph_style_id_from_xml(const pugi::xml_node& paragraph) {
  const pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  if (!p_pr) {
    return {};
  }
  const pugi::xml_node p_style = child_named_local(p_pr, "w:pStyle");
  if (!p_style) {
    return {};
  }
  return p_style.attribute("w:val").value();
}

int read_heading_level_from_style_id(std::string_view style_id) {
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

RunStyle read_first_run_style_from_xml(const pugi::xml_node& paragraph) {
  for (const pugi::xml_node& child : paragraph.children()) {
    if (std::string_view(child.name()) == "w:r") {
      return run_style_from_xml_local(child);
    }
    if (std::string_view(child.name()) == "w:hyperlink") {
      for (const pugi::xml_node& run : child.children("w:r")) {
        return run_style_from_xml_local(run);
      }
    }
  }
  return {};
}

std::vector<Run> read_runs_from_xml(const pugi::xml_node& paragraph) {
  std::vector<Run> runs;
  for (const pugi::xml_node& child : paragraph.children()) {
    if (std::string_view(child.name()) == "w:r") {
      std::string text;
      collect_text_local(child, text);
      runs.emplace_back(text, run_style_from_xml_local(child));
      continue;
    }
    if (std::string_view(child.name()) == "w:hyperlink") {
      for (const pugi::xml_node& run : child.children("w:r")) {
        std::string text;
        collect_text_local(run, text);
        runs.emplace_back(text, run_style_from_xml_local(run));
      }
    }
  }
  return runs;
}

void set_paragraph_text_in_xml(pugi::xml_node paragraph, const std::string& text,
                               const RunStyle& style) {
  pugi::xml_node existing_properties = child_named_local(paragraph, "w:pPr");
  pugi::xml_document preserved_properties_doc;
  if (existing_properties) {
    preserved_properties_doc.append_copy(existing_properties);
  }
  clear_children_local(paragraph);
  if (preserved_properties_doc.first_child()) {
    paragraph.append_copy(preserved_properties_doc.first_child());
  }
  pugi::xml_node run = paragraph.append_child("w:r");
  apply_run_style_for_model(run, style);
  append_run_text_for_model(run, text);
}

void set_paragraph_runs_in_xml(pugi::xml_node paragraph, const std::vector<Run>& runs) {
  pugi::xml_node existing_properties = child_named_local(paragraph, "w:pPr");
  pugi::xml_document preserved_properties_doc;
  if (existing_properties) {
    preserved_properties_doc.append_copy(existing_properties);
  }
  clear_children_local(paragraph);
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
    apply_run_style_for_model(run, run_value.style());
    append_run_text_for_model(run, run_value.text());
  }
}

Paragraph paragraph_from_runs(const std::vector<Run>& runs, ParagraphAlignment alignment,
                              std::string style_id, int heading_level) {
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
  return Paragraph(combined_text_local(runs), alignment, first_style, has_page_break, runs,
                   std::move(style_id), heading_level, {}, {}, {});
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

}  // namespace docxcpp
