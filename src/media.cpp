#include "internal/media_support.hpp"

#include <algorithm>
#include <cctype>
#include <fstream>
#include <sstream>
#include <stdexcept>
#include <string_view>

namespace docxcpp {

namespace {

constexpr const char* kDocumentRelationshipsEntry = "word/_rels/document.xml.rels";
constexpr const char* kImageRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kPictureUri =
    "http://schemas.openxmlformats.org/drawingml/2006/picture";
constexpr std::int64_t kEmuPerPoint = 12700;

std::vector<std::uint8_t> read_binary_file(const std::filesystem::path& path) {
  std::ifstream stream(path, std::ios::binary);
  if (!stream) {
    throw std::runtime_error("failed to open file: " + path.string());
  }
  stream.seekg(0, std::ios::end);
  const auto size = static_cast<std::size_t>(stream.tellg());
  stream.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(size);
  if (size > 0) {
    stream.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(size));
  }
  return bytes;
}

std::uint32_t read_be32(const std::vector<std::uint8_t>& bytes, std::size_t offset) {
  return (static_cast<std::uint32_t>(bytes[offset]) << 24U) |
         (static_cast<std::uint32_t>(bytes[offset + 1]) << 16U) |
         (static_cast<std::uint32_t>(bytes[offset + 2]) << 8U) |
         static_cast<std::uint32_t>(bytes[offset + 3]);
}

std::uint16_t read_be16(const std::vector<std::uint8_t>& bytes, std::size_t offset) {
  return static_cast<std::uint16_t>((bytes[offset] << 8U) | bytes[offset + 1]);
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
  return child_named_local(child_named_local(xml, "w:document"), "w:body");
}

void ensure_picture_namespaces_local(pugi::xml_document& xml) {
  pugi::xml_node document = child_named_local(xml, "w:document");
  auto ensure_attr = [&](const char* name, const char* value) {
    if (!document.attribute(name)) {
      document.append_attribute(name).set_value(value);
    }
  };
  ensure_attr("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
  ensure_attr("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
}

std::size_t next_docpr_id_local(const pugi::xml_node& node, std::size_t current_max = 0) {
  if (std::string_view(node.name()) == "wp:docPr") {
    current_max = std::max(current_max, static_cast<std::size_t>(node.attribute("id").as_uint()));
  }
  for (const pugi::xml_node& child : node.children()) {
    current_max = next_docpr_id_local(child, current_max);
  }
  return current_max;
}

std::size_t count_named_descendants_local(const pugi::xml_node& node, std::string_view name) {
  std::size_t count = std::string_view(node.name()) == name ? 1 : 0;
  for (const pugi::xml_node& child : node.children()) {
    count += count_named_descendants_local(child, name);
  }
  return count;
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

}  // namespace

ImageInfo load_image_info(const std::filesystem::path& image_path) {
  const std::vector<std::uint8_t> bytes = read_binary_file(image_path);
  std::string extension = image_path.extension().string();
  std::transform(extension.begin(), extension.end(), extension.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  if (!extension.empty() && extension.front() == '.') {
    extension.erase(extension.begin());
  }
  return load_image_info(bytes, extension, image_path.filename().string());
}

ImageInfo load_image_info(const std::vector<std::uint8_t>& bytes, std::string extension,
                          std::string filename) {
  ImageInfo info;
  info.filename = std::move(filename);
  info.bytes = bytes;
  std::transform(extension.begin(), extension.end(), extension.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  if (!extension.empty() && extension.front() == '.') {
    extension.erase(extension.begin());
  }
  info.extension = extension;

  std::uint32_t width = 0;
  std::uint32_t height = 0;
  double dpi_x = 96.0;
  double dpi_y = 96.0;

  if (extension == "png") {
    if (info.bytes.size() < 24 || read_be32(info.bytes, 12) != 0x49484452U) {
      throw std::runtime_error("unsupported or invalid PNG: " + info.filename);
    }
    info.content_type = "image/png";
    width = read_be32(info.bytes, 16);
    height = read_be32(info.bytes, 20);

    std::size_t offset = 8;
    while (offset + 12 <= info.bytes.size()) {
      const std::uint32_t chunk_length = read_be32(info.bytes, offset);
      const std::uint32_t chunk_type = read_be32(info.bytes, offset + 4);
      if (chunk_type == 0x70485973U && offset + 8 + chunk_length <= info.bytes.size()) {
        const std::uint32_t x_ppm = read_be32(info.bytes, offset + 8);
        const std::uint32_t y_ppm = read_be32(info.bytes, offset + 12);
        const std::uint8_t unit = info.bytes[offset + 16];
        if (unit == 1 && x_ppm > 0 && y_ppm > 0) {
          dpi_x = static_cast<double>(x_ppm) * 0.0254;
          dpi_y = static_cast<double>(y_ppm) * 0.0254;
        }
        break;
      }
      offset += static_cast<std::size_t>(chunk_length) + 12;
    }
  } else if (extension == "jpg" || extension == "jpeg") {
    info.content_type = "image/jpeg";
    std::size_t offset = 2;
    while (offset + 4 < info.bytes.size()) {
      if (info.bytes[offset] != 0xFF) {
        ++offset;
        continue;
      }

      const std::uint8_t marker = info.bytes[offset + 1];
      if (marker == 0xD8 || marker == 0xD9) {
        offset += 2;
        continue;
      }

      const std::uint16_t segment_length = read_be16(info.bytes, offset + 2);
      if (segment_length < 2 || offset + 2 + segment_length > info.bytes.size()) {
        break;
      }

      if (marker == 0xE0 && segment_length >= 14 &&
          std::string(reinterpret_cast<const char*>(&info.bytes[offset + 4]), 5) == "JFIF\0") {
        const std::uint8_t units = info.bytes[offset + 11];
        const std::uint16_t x_density = read_be16(info.bytes, offset + 12);
        const std::uint16_t y_density = read_be16(info.bytes, offset + 14);
        if (x_density > 0 && y_density > 0) {
          if (units == 1) {
            dpi_x = x_density;
            dpi_y = y_density;
          } else if (units == 2) {
            dpi_x = x_density * 2.54;
            dpi_y = y_density * 2.54;
          }
        }
      }

      const bool is_sof =
          (marker >= 0xC0 && marker <= 0xC3) || (marker >= 0xC5 && marker <= 0xC7) ||
          (marker >= 0xC9 && marker <= 0xCB) || (marker >= 0xCD && marker <= 0xCF);
      if (is_sof && segment_length >= 7) {
        height = read_be16(info.bytes, offset + 5);
        width = read_be16(info.bytes, offset + 7);
        break;
      }

      offset += 2 + segment_length;
    }

    if (width == 0 || height == 0) {
      throw std::runtime_error("unsupported or invalid JPEG: " + info.filename);
    }
  } else {
    throw std::runtime_error("only png/jpg/jpeg images are supported right now");
  }

  info.cx_emu = static_cast<std::int64_t>(914400.0 * static_cast<double>(width) / dpi_x);
  info.cy_emu = static_cast<std::int64_t>(914400.0 * static_cast<double>(height) / dpi_y);
  return info;
}

std::int64_t scaled_emu(std::int64_t original_primary, std::int64_t original_secondary,
                        int requested_primary_pt, int requested_secondary_pt, bool primary_axis) {
  const std::int64_t requested_primary =
      requested_primary_pt > 0 ? static_cast<std::int64_t>(requested_primary_pt) * kEmuPerPoint : 0;
  const std::int64_t requested_secondary =
      requested_secondary_pt > 0 ? static_cast<std::int64_t>(requested_secondary_pt) * kEmuPerPoint : 0;

  if (requested_primary > 0 && requested_secondary > 0) {
    return primary_axis ? requested_primary : requested_secondary;
  }
  if (requested_primary > 0) {
    if (primary_axis) {
      return requested_primary;
    }
    return static_cast<std::int64_t>(
        (static_cast<long double>(original_secondary) * requested_primary) / original_primary);
  }
  if (requested_secondary > 0) {
    if (!primary_axis) {
      return requested_secondary;
    }
    return static_cast<std::int64_t>(
        (static_cast<long double>(original_primary) * requested_secondary) / original_secondary);
  }
  return primary_axis ? original_primary : original_secondary;
}

void append_picture_run(pugi::xml_document& xml, pugi::xml_node paragraph, const std::string& rel_id,
                        const ImageInfo& image, std::int64_t cx_emu, std::int64_t cy_emu) {
  ensure_picture_namespaces_local(xml);
  pugi::xml_node run = paragraph.append_child("w:r");
  pugi::xml_node drawing = run.append_child("w:drawing");
  pugi::xml_node inline_node = drawing.append_child("wp:inline");
  inline_node.append_attribute("distT").set_value("0");
  inline_node.append_attribute("distB").set_value("0");
  inline_node.append_attribute("distL").set_value("0");
  inline_node.append_attribute("distR").set_value("0");

  pugi::xml_node extent = inline_node.append_child("wp:extent");
  extent.append_attribute("cx").set_value(std::to_string(cx_emu).c_str());
  extent.append_attribute("cy").set_value(std::to_string(cy_emu).c_str());

  const std::size_t doc_pr_id = next_docpr_id_local(xml, 0) + 1;
  pugi::xml_node doc_pr = inline_node.append_child("wp:docPr");
  doc_pr.append_attribute("id").set_value(std::to_string(doc_pr_id).c_str());
  doc_pr.append_attribute("name").set_value(("Picture " + std::to_string(doc_pr_id)).c_str());

  pugi::xml_node graphic_frame = inline_node.append_child("wp:cNvGraphicFramePr");
  graphic_frame.append_child("a:graphicFrameLocks")
      .append_attribute("noChangeAspect")
      .set_value("1");

  pugi::xml_node graphic = inline_node.append_child("a:graphic");
  pugi::xml_node graphic_data = graphic.append_child("a:graphicData");
  graphic_data.append_attribute("uri").set_value(kPictureUri);

  pugi::xml_node pic = graphic_data.append_child("pic:pic");
  pugi::xml_node nv_pic_pr = pic.append_child("pic:nvPicPr");
  pugi::xml_node c_nv_pr = nv_pic_pr.append_child("pic:cNvPr");
  c_nv_pr.append_attribute("id").set_value("0");
  c_nv_pr.append_attribute("name").set_value(image.filename.c_str());
  nv_pic_pr.append_child("pic:cNvPicPr");

  pugi::xml_node blip_fill = pic.append_child("pic:blipFill");
  blip_fill.append_child("a:blip").append_attribute("r:embed").set_value(rel_id.c_str());
  pugi::xml_node stretch = blip_fill.append_child("a:stretch");
  stretch.append_child("a:fillRect");

  pugi::xml_node sp_pr = pic.append_child("pic:spPr");
  pugi::xml_node xfrm = sp_pr.append_child("a:xfrm");
  pugi::xml_node off = xfrm.append_child("a:off");
  off.append_attribute("x").set_value("0");
  off.append_attribute("y").set_value("0");
  pugi::xml_node ext = xfrm.append_child("a:ext");
  ext.append_attribute("cx").set_value(std::to_string(cx_emu).c_str());
  ext.append_attribute("cy").set_value(std::to_string(cy_emu).c_str());
  sp_pr.append_child("a:prstGeom").append_attribute("prst").set_value("rect");
}

PictureSize picture_size_from_inline(const pugi::xml_node& inline_node) {
  const pugi::xml_node extent = child_named_local(inline_node, "wp:extent");
  PictureSize size;
  size.width_pt = static_cast<int>(extent.attribute("cx").as_llong() / kEmuPerPoint);
  size.height_pt = static_cast<int>(extent.attribute("cy").as_llong() / kEmuPerPoint);
  return size;
}

RelTargetMap image_relationship_targets(const OpcPackage& package) {
  RelTargetMap targets;
  const pugi::xml_document rels_xml = parse_package_xml_local(package, kDocumentRelationshipsEntry);
  const pugi::xml_node relationships = child_named_local(rels_xml, "Relationships");
  for (const pugi::xml_node& rel : relationships.children("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) != kImageRelationshipType) {
      continue;
    }
    targets.emplace(rel.attribute("Id").value(), rel.attribute("Target").value());
  }
  return targets;
}

void collect_pictures_from_paragraph(const pugi::xml_node& paragraph, std::size_t paragraph_index,
                                     const RelTargetMap& rel_targets,
                                     std::vector<PictureInfo>& pictures) {
  for (const pugi::xml_node& node : paragraph.children()) {
    for (const pugi::xml_node& child : node.children()) {
      if (std::string_view(child.name()) != "w:drawing") {
        continue;
      }
      for (const pugi::xml_node& inline_node : child.children("wp:inline")) {
        PictureInfo info;
        info.size_pt = picture_size_from_inline(inline_node);
        info.paragraph_index = paragraph_index;
        const pugi::xml_node blip =
            child_named_local(child_named_local(child_named_local(child_named_local(inline_node, "a:graphic"),
                                                                  "a:graphicData"),
                                                "pic:pic"),
                              "pic:blipFill");
        pugi::xml_node embed_holder = child_named_local(blip, "a:blip");
        info.relationship_id = embed_holder.attribute("r:embed").value();
        const auto it = rel_targets.find(info.relationship_id);
        if (it != rel_targets.end()) {
          info.target = it->second;
        }
        pictures.push_back(std::move(info));
      }
    }
  }
}

std::size_t image_count_in_xml(const pugi::xml_document& xml) {
  return count_named_descendants_local(xml, "a:blip");
}

}  // namespace docxcpp
