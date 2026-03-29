#include "internal/layout_support.hpp"

#include <stdexcept>
#include <string>
#include <string_view>

namespace docxcpp {

namespace {

constexpr std::int64_t kTwipsPerPoint = 20;

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

pugi::xml_node document_sect_pr_local(pugi::xml_document& xml) {
  pugi::xml_node body = document_body_local(xml);
  pugi::xml_node sect_pr = child_named_local(body, "w:sectPr");
  if (!sect_pr) {
    sect_pr = body.append_child("w:sectPr");
  }
  return sect_pr;
}

pugi::xml_node get_or_add_child_local(pugi::xml_node parent, const char* name) {
  pugi::xml_node child = child_named_local(parent, name);
  if (!child) {
    child = parent.append_child(name);
  }
  return child;
}

const char* orientation_value_local(PageOrientation orientation) {
  switch (orientation) {
    case PageOrientation::Landscape:
      return "landscape";
    case PageOrientation::Portrait:
    default:
      return "portrait";
  }
}

PageOrientation orientation_from_value_local(std::string_view value) {
  if (value == "landscape") {
    return PageOrientation::Landscape;
  }
  return PageOrientation::Portrait;
}

Section section_from_sect_pr_local(const pugi::xml_node& sect_pr) {
  PageSize size;
  if (const pugi::xml_node pg_sz = child_named_local(sect_pr, "w:pgSz")) {
    size.width_pt = static_cast<int>(pg_sz.attribute("w:w").as_int() / kTwipsPerPoint);
    size.height_pt = static_cast<int>(pg_sz.attribute("w:h").as_int() / kTwipsPerPoint);
  }

  PageOrientation orientation = PageOrientation::Portrait;
  if (const pugi::xml_node pg_sz = child_named_local(sect_pr, "w:pgSz")) {
    orientation = orientation_from_value_local(pg_sz.attribute("w:orient").value());
  }

  PageMargins margins;
  if (const pugi::xml_node pg_mar = child_named_local(sect_pr, "w:pgMar")) {
    margins.top_pt = pg_mar.attribute("w:top").as_int() / kTwipsPerPoint;
    margins.right_pt = pg_mar.attribute("w:right").as_int() / kTwipsPerPoint;
    margins.bottom_pt = pg_mar.attribute("w:bottom").as_int() / kTwipsPerPoint;
    margins.left_pt = pg_mar.attribute("w:left").as_int() / kTwipsPerPoint;
    margins.header_pt = pg_mar.attribute("w:header").as_int() / kTwipsPerPoint;
    margins.footer_pt = pg_mar.attribute("w:footer").as_int() / kTwipsPerPoint;
    margins.gutter_pt = pg_mar.attribute("w:gutter").as_int() / kTwipsPerPoint;
  }

  return Section(size, orientation, margins);
}

}  // namespace

std::vector<Section> read_sections_from_xml(const pugi::xml_document& xml) {
  std::vector<Section> sections;
  const pugi::xml_node body = document_body_local(xml);
  for (const pugi::xml_node& child : body.children()) {
    if (std::string_view(child.name()) == "w:p") {
      if (const pugi::xml_node p_pr = child_named_local(child, "w:pPr")) {
        if (const pugi::xml_node sect_pr = child_named_local(p_pr, "w:sectPr")) {
          sections.push_back(section_from_sect_pr_local(sect_pr));
        }
      }
      continue;
    }
    if (std::string_view(child.name()) == "w:sectPr") {
      sections.push_back(section_from_sect_pr_local(child));
    }
  }
  if (sections.empty()) {
    sections.emplace_back();
  }
  return sections;
}

void set_section_properties_in_xml(pugi::xml_node sect_pr, const Section& section) {
  pugi::xml_node pg_sz = get_or_add_child_local(sect_pr, "w:pgSz");
  const auto width_twips = std::to_string(section.page_size_pt().width_pt * kTwipsPerPoint);
  const auto height_twips = std::to_string(section.page_size_pt().height_pt * kTwipsPerPoint);
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
  const char* orientation = orientation_value_local(section.page_orientation());
  if (pg_sz.attribute("w:orient")) {
    pg_sz.attribute("w:orient").set_value(orientation);
  } else {
    pg_sz.append_attribute("w:orient").set_value(orientation);
  }

  pugi::xml_node pg_mar = get_or_add_child_local(sect_pr, "w:pgMar");
  auto set_margin_attr = [&](const char* name, int pt_value) {
    const auto twips = std::to_string(pt_value * kTwipsPerPoint);
    if (pg_mar.attribute(name)) {
      pg_mar.attribute(name).set_value(twips.c_str());
    } else {
      pg_mar.append_attribute(name).set_value(twips.c_str());
    }
  };

  const PageMargins& margins = section.page_margins_pt();
  set_margin_attr("w:top", margins.top_pt);
  set_margin_attr("w:right", margins.right_pt);
  set_margin_attr("w:bottom", margins.bottom_pt);
  set_margin_attr("w:left", margins.left_pt);
  set_margin_attr("w:header", margins.header_pt);
  set_margin_attr("w:footer", margins.footer_pt);
  set_margin_attr("w:gutter", margins.gutter_pt);
}

void set_page_size_pt_in_xml(pugi::xml_document& xml, int width_pt, int height_pt) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  PageSize size = section.page_size_pt();
  size.width_pt = width_pt;
  size.height_pt = height_pt;
  set_section_properties_in_xml(sect_pr,
                                Section(size, section.page_orientation(), section.page_margins_pt()));
}

void set_page_orientation_in_xml(pugi::xml_document& xml, PageOrientation orientation) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  set_section_properties_in_xml(sect_pr,
                                Section(section.page_size_pt(), orientation, section.page_margins_pt()));
}

void set_page_margins_pt_in_xml(pugi::xml_document& xml, const PageMargins& margins) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  set_section_properties_in_xml(sect_pr,
                                Section(section.page_size_pt(), section.page_orientation(), margins));
}

PageSize read_page_size_pt_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_size_pt();
}

PageOrientation read_page_orientation_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_orientation();
}

PageMargins read_page_margins_pt_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_margins_pt();
}

}  // namespace docxcpp
