#pragma once

#include <algorithm>
#include <cctype>
#include <string>
#include <string_view>
#include <unordered_map>

#include "pugixml.hpp"

#include "docxcpp/opc_package.hpp"
#include "docxcpp/paragraph.hpp"

namespace docxcpp {

enum class NamedStyleType {
  Character,
  Table,
};

struct NamedStyleInfo {
  std::string id;
  std::string name;
};

struct StyleCatalog {
  std::unordered_map<std::string, std::string> character_name_by_id;
  std::unordered_map<std::string, std::string> table_name_by_id;
  std::unordered_map<std::string, std::string> character_lookup_to_id;
  std::unordered_map<std::string, std::string> table_lookup_to_id;
};

inline std::string normalized_style_lookup_key(std::string_view value) {
  std::string normalized;
  normalized.reserve(value.size());
  for (const unsigned char ch : value) {
    if (std::isalnum(ch) != 0) {
      normalized.push_back(static_cast<char>(std::tolower(ch)));
    }
  }
  return normalized;
}

inline pugi::xml_node child_named_style_local(const pugi::xml_node& parent, const char* name) {
  for (const pugi::xml_node& child : parent.children()) {
    if (std::string_view(child.name()) == name) {
      return child;
    }
  }
  return {};
}

inline void register_named_style_local(std::unordered_map<std::string, std::string>& name_by_id,
                                       std::unordered_map<std::string, std::string>& lookup_to_id,
                                       const std::string& id, const std::string& name) {
  if (id.empty()) {
    return;
  }
  if (!name.empty()) {
    name_by_id[id] = name;
  }

  lookup_to_id[normalized_style_lookup_key(id)] = id;
  if (!name.empty()) {
    lookup_to_id[normalized_style_lookup_key(name)] = id;
  }
}

inline StyleCatalog load_style_catalog(const OpcPackage& package) {
  StyleCatalog catalog;
  if (!package.has_entry("word/styles.xml")) {
    return catalog;
  }

  pugi::xml_document xml;
  const auto& bytes = package.entry("word/styles.xml");
  const auto result =
      xml.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!result) {
    return catalog;
  }

  const pugi::xml_node styles = child_named_style_local(xml, "w:styles");
  for (const pugi::xml_node& style : styles.children("w:style")) {
    const std::string type = style.attribute("w:type").value();
    const std::string id = style.attribute("w:styleId").value();
    const std::string name = child_named_style_local(style, "w:name").attribute("w:val").value();
    if (type == "character") {
      register_named_style_local(catalog.character_name_by_id, catalog.character_lookup_to_id, id,
                                 name);
    } else if (type == "table") {
      register_named_style_local(catalog.table_name_by_id, catalog.table_lookup_to_id, id, name);
    }
  }
  return catalog;
}

inline const std::unordered_map<std::string, std::string>& style_name_map(const StyleCatalog& catalog,
                                                                           NamedStyleType type) {
  return type == NamedStyleType::Character ? catalog.character_name_by_id : catalog.table_name_by_id;
}

inline const std::unordered_map<std::string, std::string>& style_lookup_map(const StyleCatalog& catalog,
                                                                             NamedStyleType type) {
  return type == NamedStyleType::Character ? catalog.character_lookup_to_id
                                           : catalog.table_lookup_to_id;
}

inline NamedStyleInfo style_info_from_id(const StyleCatalog& catalog, NamedStyleType type,
                                         std::string_view style_id) {
  if (style_id.empty()) {
    return {};
  }
  NamedStyleInfo info;
  info.id = std::string(style_id);
  const auto& name_by_id = style_name_map(catalog, type);
  if (const auto it = name_by_id.find(info.id); it != name_by_id.end()) {
    info.name = it->second;
  }
  return info;
}

inline NamedStyleInfo resolve_style_reference(const StyleCatalog& catalog, NamedStyleType type,
                                              std::string_view style_id_or_name) {
  if (style_id_or_name.empty()) {
    return {};
  }

  const std::string key = normalized_style_lookup_key(style_id_or_name);
  const auto& lookup_to_id = style_lookup_map(catalog, type);
  if (const auto it = lookup_to_id.find(key); it != lookup_to_id.end()) {
    return style_info_from_id(catalog, type, it->second);
  }
  return style_info_from_id(catalog, type, style_id_or_name);
}

inline RunStyle resolve_run_style_reference(const RunStyle& style,
                                            const StyleCatalog* style_catalog) {
  RunStyle resolved = style;
  if (style_catalog == nullptr) {
    return resolved;
  }

  if (resolved.character_style_id.empty() && !resolved.character_style_name.empty()) {
    const NamedStyleInfo info =
        resolve_style_reference(*style_catalog, NamedStyleType::Character,
                                resolved.character_style_name);
    if (!info.id.empty()) {
      resolved.character_style_id = info.id;
    }
    if (!info.name.empty()) {
      resolved.character_style_name = info.name;
    }
  } else if (!resolved.character_style_id.empty() && resolved.character_style_name.empty()) {
    resolved.character_style_name =
        style_info_from_id(*style_catalog, NamedStyleType::Character,
                           resolved.character_style_id)
            .name;
  }

  return resolved;
}

}  // namespace docxcpp
