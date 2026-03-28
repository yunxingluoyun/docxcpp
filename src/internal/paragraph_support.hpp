#pragma once

#include <string>
#include <vector>

#include "pugixml.hpp"

#include "docxcpp/paragraph.hpp"

namespace docxcpp {

ParagraphFormat read_paragraph_format_from_xml(const pugi::xml_node& paragraph);
std::string read_paragraph_style_id_from_xml(const pugi::xml_node& paragraph);
int read_heading_level_from_style_id(std::string_view style_id);
RunStyle read_first_run_style_from_xml(const pugi::xml_node& paragraph);
std::vector<Run> read_runs_from_xml(const pugi::xml_node& paragraph);
void set_paragraph_text_in_xml(pugi::xml_node paragraph, const std::string& text,
                               const RunStyle& style = {});
void set_paragraph_runs_in_xml(pugi::xml_node paragraph, const std::vector<Run>& runs);
Paragraph paragraph_from_runs(const std::vector<Run>& runs, ParagraphAlignment alignment,
                              std::string style_id = {}, int heading_level = -1);
std::string paragraph_style_id_for_level(int level);

}  // namespace docxcpp
