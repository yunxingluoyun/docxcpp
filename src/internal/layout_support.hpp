#pragma once

#include "pugixml.hpp"

#include "docxcpp/document.hpp"

namespace docxcpp {

std::vector<Section> read_sections_from_xml(const pugi::xml_document& xml);
void set_page_size_pt_in_xml(pugi::xml_document& xml, int width_pt, int height_pt);
void set_page_orientation_in_xml(pugi::xml_document& xml, PageOrientation orientation);
void set_page_margins_pt_in_xml(pugi::xml_document& xml, const PageMargins& margins);
PageSize read_page_size_pt_from_xml(const pugi::xml_document& xml);
PageOrientation read_page_orientation_from_xml(const pugi::xml_document& xml);
PageMargins read_page_margins_pt_from_xml(const pugi::xml_document& xml);

}  // namespace docxcpp
