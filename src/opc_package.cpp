#include "docxcpp/opc_package.hpp"

#include <fstream>
#include <stdexcept>
#include <string>
#include <utility>

#include "miniz.h"

namespace docxcpp {

namespace {

std::vector<std::uint8_t> read_file(const std::filesystem::path& path) {
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

void write_file(const std::filesystem::path& path, const std::vector<std::uint8_t>& bytes) {
  std::ofstream stream(path, std::ios::binary);
  if (!stream) {
    throw std::runtime_error("failed to open output file: " + path.string());
  }

  if (!bytes.empty()) {
    stream.write(reinterpret_cast<const char*>(bytes.data()),
                 static_cast<std::streamsize>(bytes.size()));
  }
}

std::string normalize_entry_name(const std::filesystem::path& root,
                                 const std::filesystem::path& file) {
  return std::filesystem::relative(file, root).generic_string();
}

}  // namespace

OpcPackage OpcPackage::open(const std::filesystem::path& path) {
  OpcPackage package;
  const std::vector<std::uint8_t> archive_bytes = read_file(path);

  mz_zip_archive archive{};
  if (!mz_zip_reader_init_mem(&archive, archive_bytes.data(), archive_bytes.size(), 0)) {
    throw std::runtime_error("failed to open zip archive: " + path.string());
  }

  const mz_uint file_count = mz_zip_reader_get_num_files(&archive);
  for (mz_uint index = 0; index < file_count; ++index) {
    mz_zip_archive_file_stat stat{};
    if (!mz_zip_reader_file_stat(&archive, index, &stat)) {
      mz_zip_reader_end(&archive);
      throw std::runtime_error("failed to stat zip entry in: " + path.string());
    }
    if (mz_zip_reader_is_file_a_directory(&archive, index)) {
      continue;
    }

    size_t size = 0;
    void* entry_data = mz_zip_reader_extract_to_heap(&archive, index, &size, 0);
    if (entry_data == nullptr) {
      mz_zip_reader_end(&archive);
      throw std::runtime_error("failed to extract zip entry: " + std::string(stat.m_filename));
    }

    std::vector<std::uint8_t> bytes(static_cast<std::uint8_t*>(entry_data),
                                    static_cast<std::uint8_t*>(entry_data) + size);
    mz_free(entry_data);
    package.entries_.emplace(stat.m_filename, std::move(bytes));
  }

  mz_zip_reader_end(&archive);
  return package;
}

OpcPackage OpcPackage::from_directory(const std::filesystem::path& directory) {
  if (!std::filesystem::exists(directory)) {
    throw std::runtime_error("template directory does not exist: " + directory.string());
  }

  OpcPackage package;
  for (const auto& entry : std::filesystem::recursive_directory_iterator(directory)) {
    if (!entry.is_regular_file()) {
      continue;
    }
    package.entries_.emplace(normalize_entry_name(directory, entry.path()), read_file(entry.path()));
  }
  return package;
}

const std::vector<std::uint8_t>& OpcPackage::entry(const std::string& name) const {
  const auto it = entries_.find(name);
  if (it == entries_.end()) {
    throw std::runtime_error("missing OPC entry: " + name);
  }
  return it->second;
}

bool OpcPackage::has_entry(const std::string& name) const noexcept {
  return entries_.find(name) != entries_.end();
}

std::vector<std::string> OpcPackage::entry_names() const {
  std::vector<std::string> names;
  names.reserve(entries_.size());
  for (const auto& [name, _] : entries_) {
    names.push_back(name);
  }
  return names;
}

void OpcPackage::set_entry(std::string name, std::vector<std::uint8_t> bytes) {
  entries_[std::move(name)] = std::move(bytes);
}

void OpcPackage::save(const std::filesystem::path& path) const {
  mz_zip_archive archive{};
  if (!mz_zip_writer_init_heap(&archive, 0, 0)) {
    throw std::runtime_error("failed to initialize zip writer");
  }

  for (const auto& [name, bytes] : entries_) {
    if (!mz_zip_writer_add_mem(&archive, name.c_str(), bytes.data(), bytes.size(),
                               MZ_BEST_COMPRESSION)) {
      mz_zip_writer_end(&archive);
      throw std::runtime_error("failed to add zip entry: " + name);
    }
  }

  void* output = nullptr;
  size_t size = 0;
  if (!mz_zip_writer_finalize_heap_archive(&archive, &output, &size)) {
    mz_zip_writer_end(&archive);
    throw std::runtime_error("failed to finalize zip archive");
  }

  std::vector<std::uint8_t> buffer(static_cast<std::uint8_t*>(output),
                                   static_cast<std::uint8_t*>(output) + size);
  mz_free(output);
  mz_zip_writer_end(&archive);

  write_file(path, buffer);
}

}  // namespace docxcpp
