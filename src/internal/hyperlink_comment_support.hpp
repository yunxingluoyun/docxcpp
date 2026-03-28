#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include "pugixml.hpp"

#include "docxcpp/document.hpp"

namespace docxcpp {

using RelTargetMap = std::unordered_map<std::string, std::string>;

RelTargetMap hyperlink_relationship_targets(const OpcPackage& package);
std::vector<HyperlinkInfo> hyperlinks_from_xml(const pugi::xml_node& paragraph,
                                               const RelTargetMap& hyperlink_targets);
std::string register_external_hyperlink_relationship(OpcPackage& package, const std::string& target);
void append_external_hyperlink(pugi::xml_node paragraph, const std::string& rel_id,
                               const std::string& text, const RunStyle& style);
std::size_t add_comment_to_paragraph(OpcPackage& package, pugi::xml_node paragraph,
                                     const std::string& text, const std::string& author,
                                     const std::string& initials);
std::vector<CommentInfo> read_comments(const OpcPackage& package);

}  // namespace docxcpp
