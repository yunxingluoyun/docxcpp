#include "docxcpp/document.hpp"

#include <cstdlib>
#include <optional>
#include <stdexcept>
#include <string_view>

#include "pugixml.hpp"

namespace docxcpp {

namespace {

constexpr double kAutoLineUnit = 240.0;

const char* tab_alignment_value_local(TabAlignment alignment) {
  switch (alignment) {
    case TabAlignment::Center:
      return "center";
    case TabAlignment::Right:
      return "right";
    case TabAlignment::Decimal:
      return "decimal";
    case TabAlignment::Bar:
      return "bar";
    case TabAlignment::Left:
    default:
      return "left";
  }
}

const char* tab_leader_value_local(TabLeader leader) {
  switch (leader) {
    case TabLeader::Dots:
      return "dot";
    case TabLeader::Dashes:
      return "hyphen";
    case TabLeader::Lines:
      return "underscore";
    case TabLeader::Heavy:
      return "heavy";
    case TabLeader::MiddleDot:
      return "middleDot";
    case TabLeader::Spaces:
    default:
      return "none";
  }
}

pugi::xml_node child_named_local(const pugi::xml_node& parent, const char* name) {
  for (const pugi::xml_node& child : parent.children()) {
    if (std::string_view(child.name()) == name) {
      return child;
    }
  }
  return {};
}

pugi::xml_node document_body_local(const pugi::xml_document& xml) {
  pugi::xml_node document = child_named_local(xml, "w:document");
  if (!document) {
    throw std::runtime_error("required DOCX node is missing: w:document");
  }
  pugi::xml_node body = child_named_local(document, "w:body");
  if (!body) {
    throw std::runtime_error("required DOCX node is missing: w:body");
  }
  return body;
}

pugi::xml_node nth_child_named_local(pugi::xml_node parent, const char* name, std::size_t index) {
  std::size_t current = 0;
  for (pugi::xml_node child : parent.children(name)) {
    if (current == index) {
      return child;
    }
    ++current;
  }
  return {};
}

pugi::xml_node get_or_add_paragraph_properties_local(pugi::xml_node paragraph) {
  pugi::xml_node p_pr = child_named_local(paragraph, "w:pPr");
  if (!p_pr) {
    p_pr = paragraph.prepend_child("w:pPr");
  }
  return p_pr;
}

void cleanup_if_empty_local(pugi::xml_node parent, pugi::xml_node child) {
  if (child && !child.first_attribute() && !child.first_child()) {
    parent.remove_child(child);
  }
}

void apply_indent_attr_local(pugi::xml_node ind, const char* name, std::optional<int> pt_value) {
  if (!pt_value.has_value()) {
    if (ind.attribute(name)) {
      ind.remove_attribute(name);
    }
    return;
  }
  const auto twips = std::to_string(*pt_value * 20);
  if (ind.attribute(name)) {
    ind.attribute(name).set_value(twips.c_str());
  } else {
    ind.append_attribute(name).set_value(twips.c_str());
  }
}

void set_on_off_property_local(pugi::xml_node parent, const char* name, std::optional<bool> value) {
  pugi::xml_node node = child_named_local(parent, name);
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

pugi::xml_node paragraph_by_index_or_throw(pugi::xml_document& xml, std::size_t paragraph_index) {
  pugi::xml_node paragraph = nth_child_named_local(document_body_local(xml), "w:p", paragraph_index);
  if (!paragraph) {
    throw std::out_of_range("paragraph_index out of range");
  }
  return paragraph;
}

const char* alignment_value_local(ParagraphAlignment alignment) {
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

void apply_alignment_local(pugi::xml_node paragraph, ParagraphAlignment alignment) {
  pugi::xml_node properties = child_named_local(paragraph, "w:pPr");
  if (alignment == ParagraphAlignment::Inherit) {
    if (!properties) {
      return;
    }
    pugi::xml_node jc = child_named_local(properties, "w:jc");
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
  pugi::xml_node jc = child_named_local(properties, "w:jc");
  if (!jc) {
    jc = properties.append_child("w:jc");
  }
  if (jc.attribute("w:val")) {
    jc.attribute("w:val").set_value(alignment_value_local(alignment));
  } else {
    jc.append_attribute("w:val").set_value(alignment_value_local(alignment));
  }
}

}  // namespace

void Document::set_paragraph_alignment(std::size_t paragraph_index, ParagraphAlignment alignment) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);
  apply_alignment_local(paragraph, alignment);
  dirty_ = true;
}

void Document::set_paragraph_indentation_pt(std::size_t paragraph_index,
                                            std::optional<int> left_indent_pt,
                                            std::optional<int> right_indent_pt,
                                            std::optional<int> first_line_indent_pt) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);
  if ((left_indent_pt && *left_indent_pt < 0) || (right_indent_pt && *right_indent_pt < 0)) {
    throw std::invalid_argument("left/right paragraph indent cannot be negative");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  pugi::xml_node ind = child_named_local(p_pr, "w:ind");
  if (!ind && (left_indent_pt || right_indent_pt || first_line_indent_pt)) {
    ind = p_pr.append_child("w:ind");
  }
  if (ind) {
    apply_indent_attr_local(ind, "w:left", left_indent_pt);
    apply_indent_attr_local(ind, "w:right", right_indent_pt);
    if (ind.attribute("w:firstLine")) {
      ind.remove_attribute("w:firstLine");
    }
    if (ind.attribute("w:hanging")) {
      ind.remove_attribute("w:hanging");
    }
    if (first_line_indent_pt.has_value()) {
      const auto twips = std::to_string(std::abs(*first_line_indent_pt) * 20);
      const char* attr_name = *first_line_indent_pt >= 0 ? "w:firstLine" : "w:hanging";
      ind.append_attribute(attr_name).set_value(twips.c_str());
    }
    cleanup_if_empty_local(p_pr, ind);
    cleanup_if_empty_local(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_spacing_pt(std::size_t paragraph_index,
                                        std::optional<int> space_before_pt,
                                        std::optional<int> space_after_pt) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);
  if ((space_before_pt && *space_before_pt < 0) || (space_after_pt && *space_after_pt < 0)) {
    throw std::invalid_argument("paragraph spacing cannot be negative");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  pugi::xml_node spacing = child_named_local(p_pr, "w:spacing");
  if (!spacing && (space_before_pt || space_after_pt)) {
    spacing = p_pr.append_child("w:spacing");
  }
  if (spacing) {
    apply_indent_attr_local(spacing, "w:before", space_before_pt);
    apply_indent_attr_local(spacing, "w:after", space_after_pt);
    cleanup_if_empty_local(p_pr, spacing);
    cleanup_if_empty_local(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_line_spacing_pt(std::size_t paragraph_index,
                                             std::optional<int> line_spacing_pt) {
  set_paragraph_line_spacing(
      paragraph_index,
      line_spacing_pt ? std::optional<double>(static_cast<double>(*line_spacing_pt)) : std::nullopt,
      LineSpacingMode::Exact);
}

void Document::set_paragraph_line_spacing(std::size_t paragraph_index, std::optional<double> value,
                                          LineSpacingMode mode) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);
  if (value && *value <= 0) {
    throw std::invalid_argument("line spacing must be greater than zero");
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  pugi::xml_node spacing = child_named_local(p_pr, "w:spacing");
  if (!spacing && value) {
    spacing = p_pr.append_child("w:spacing");
  }
  if (spacing) {
    if (spacing.attribute("w:line")) {
      spacing.remove_attribute("w:line");
    }
    if (spacing.attribute("w:lineRule")) {
      spacing.remove_attribute("w:lineRule");
    }
    if (value.has_value()) {
      std::string line_value;
      const char* rule_value = nullptr;
      switch (mode) {
        case LineSpacingMode::Exact:
          line_value = std::to_string(static_cast<int>(*value * 20.0));
          rule_value = "exact";
          break;
        case LineSpacingMode::AtLeast:
          line_value = std::to_string(static_cast<int>(*value * 20.0));
          rule_value = "atLeast";
          break;
        case LineSpacingMode::Multiple:
          line_value = std::to_string(static_cast<int>(*value * kAutoLineUnit));
          rule_value = "auto";
          break;
      }
      spacing.append_attribute("w:line").set_value(line_value.c_str());
      spacing.append_attribute("w:lineRule").set_value(rule_value);
    }
    cleanup_if_empty_local(p_pr, spacing);
    cleanup_if_empty_local(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_tab_stops(std::size_t paragraph_index,
                                       const std::vector<TabStop>& tab_stops) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);
  for (const TabStop& tab_stop : tab_stops) {
    if (tab_stop.position_pt <= 0) {
      throw std::invalid_argument("tab stop position must be greater than zero");
    }
  }

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  pugi::xml_node tabs = child_named_local(p_pr, "w:tabs");
  if (!tabs && !tab_stops.empty()) {
    tabs = p_pr.append_child("w:tabs");
  }
  if (tabs) {
    while (tabs.first_child()) {
      tabs.remove_child(tabs.first_child());
    }
    for (const TabStop& tab_stop : tab_stops) {
      pugi::xml_node tab = tabs.append_child("w:tab");
      tab.append_attribute("w:val").set_value(tab_alignment_value_local(tab_stop.alignment));
      tab.append_attribute("w:leader").set_value(tab_leader_value_local(tab_stop.leader));
      tab.append_attribute("w:pos").set_value(std::to_string(tab_stop.position_pt * 20).c_str());
    }
    cleanup_if_empty_local(p_pr, tabs);
    cleanup_if_empty_local(paragraph, p_pr);
  }
  dirty_ = true;
}

void Document::set_paragraph_pagination(std::size_t paragraph_index,
                                        std::optional<bool> keep_together,
                                        std::optional<bool> keep_with_next,
                                        std::optional<bool> page_break_before) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  set_on_off_property_local(p_pr, "w:keepLines", keep_together);
  set_on_off_property_local(p_pr, "w:keepNext", keep_with_next);
  set_on_off_property_local(p_pr, "w:pageBreakBefore", page_break_before);
  cleanup_if_empty_local(paragraph, p_pr);
  dirty_ = true;
}

void Document::set_paragraph_widow_control(std::size_t paragraph_index,
                                           std::optional<bool> widow_control) {
  pugi::xml_node paragraph = paragraph_by_index_or_throw(*xml_, paragraph_index);

  pugi::xml_node p_pr = get_or_add_paragraph_properties_local(paragraph);
  set_on_off_property_local(p_pr, "w:widowControl", widow_control);
  cleanup_if_empty_local(paragraph, p_pr);
  dirty_ = true;
}

}  // namespace docxcpp
