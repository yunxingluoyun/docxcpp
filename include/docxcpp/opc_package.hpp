#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <unordered_map>
#include <vector>

namespace docxcpp {

class OpcPackage {
public:
  using EntryMap = std::unordered_map<std::string, std::vector<std::uint8_t>>;

  static OpcPackage open(const std::filesystem::path& path);
  static OpcPackage from_directory(const std::filesystem::path& directory);

  const std::vector<std::uint8_t>& entry(const std::string& name) const;
  bool has_entry(const std::string& name) const noexcept;
  std::vector<std::string> entry_names() const;
  void set_entry(std::string name, std::vector<std::uint8_t> bytes);
  void save(const std::filesystem::path& path) const;

private:
  EntryMap entries_;
};

}  // namespace docxcpp
