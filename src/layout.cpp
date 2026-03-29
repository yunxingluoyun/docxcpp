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
    size.width = Length::from_twips(pg_sz.attribute("w:w").as_llong());
    size.height = Length::from_twips(pg_sz.attribute("w:h").as_llong());
  }

  PageOrientation orientation = PageOrientation::Portrait;
  if (const pugi::xml_node pg_sz = child_named_local(sect_pr, "w:pgSz")) {
    orientation = orientation_from_value_local(pg_sz.attribute("w:orient").value());
  }

  PageMargins margins;
  if (const pugi::xml_node pg_mar = child_named_local(sect_pr, "w:pgMar")) {
    margins.top = Length::from_twips(pg_mar.attribute("w:top").as_llong());
    margins.right = Length::from_twips(pg_mar.attribute("w:right").as_llong());
    margins.bottom = Length::from_twips(pg_mar.attribute("w:bottom").as_llong());
    margins.left = Length::from_twips(pg_mar.attribute("w:left").as_llong());
    margins.header = Length::from_twips(pg_mar.attribute("w:header").as_llong());
    margins.footer = Length::from_twips(pg_mar.attribute("w:footer").as_llong());
    margins.gutter = Length::from_twips(pg_mar.attribute("w:gutter").as_llong());
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
  const auto width_twips = std::to_string(section.page_size().width.twips());
  const auto height_twips = std::to_string(section.page_size().height.twips());
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
  auto set_margin_attr = [&](const char* name, Length value) {
    const auto twips = std::to_string(value.twips());
    if (pg_mar.attribute(name)) {
      pg_mar.attribute(name).set_value(twips.c_str());
    } else {
      pg_mar.append_attribute(name).set_value(twips.c_str());
    }
  };

  const PageMargins& margins = section.page_margins();
  set_margin_attr("w:top", margins.top);
  set_margin_attr("w:right", margins.right);
  set_margin_attr("w:bottom", margins.bottom);
  set_margin_attr("w:left", margins.left);
  set_margin_attr("w:header", margins.header);
  set_margin_attr("w:footer", margins.footer);
  set_margin_attr("w:gutter", margins.gutter);
}

void set_page_size_in_xml(pugi::xml_document& xml, const PageSize& page_size) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  set_section_properties_in_xml(sect_pr, Section(page_size, section.page_orientation(), section.page_margins()));
}

void set_page_orientation_in_xml(pugi::xml_document& xml, PageOrientation orientation) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  set_section_properties_in_xml(sect_pr, Section(section.page_size(), orientation, section.page_margins()));
}

void set_page_margins_in_xml(pugi::xml_document& xml, const PageMargins& margins) {
  pugi::xml_node sect_pr = document_sect_pr_local(xml);
  Section section = section_from_sect_pr_local(sect_pr);
  set_section_properties_in_xml(sect_pr, Section(section.page_size(), section.page_orientation(), margins));
}

PageSize read_page_size_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_size();
}

PageOrientation read_page_orientation_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_orientation();
}

PageMargins read_page_margins_from_xml(const pugi::xml_document& xml) {
  return read_sections_from_xml(xml).back().page_margins();
}

}  // namespace docxcpp
