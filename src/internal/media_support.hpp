#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <unordered_map>
#include <vector>

#include "pugixml.hpp"

#include "docxcpp/document.hpp"

namespace docxcpp {

struct ImageInfo {
  std::string extension;
  std::string content_type;
  std::string filename;
  std::vector<std::uint8_t> bytes;
  std::int64_t cx_emu;
  std::int64_t cy_emu;
};

using RelTargetMap = std::unordered_map<std::string, std::string>;

ImageInfo load_image_info(const std::filesystem::path& image_path);
ImageInfo load_image_info(const std::vector<std::uint8_t>& bytes, std::string extension,
                          std::string filename);
std::int64_t scaled_emu(std::int64_t original_primary, std::int64_t original_secondary,
                        int requested_primary_pt, int requested_secondary_pt, bool primary_axis);
void append_picture_run(pugi::xml_document& xml, pugi::xml_node paragraph, const std::string& rel_id,
                        const ImageInfo& image, std::int64_t cx_emu, std::int64_t cy_emu);
PictureSize picture_size_from_inline(const pugi::xml_node& inline_node);
RelTargetMap image_relationship_targets(const OpcPackage& package);
void collect_pictures_from_paragraph(const pugi::xml_node& paragraph, std::size_t paragraph_index,
                                     const RelTargetMap& rel_targets,
                                     std::vector<PictureInfo>& pictures);
std::size_t image_count_in_xml(const pugi::xml_document& xml);

}  // namespace docxcpp
