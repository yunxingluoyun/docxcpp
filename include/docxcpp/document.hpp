#pragma once

#include <filesystem>
#include <memory>
#include <string>
#include <vector>

#include "pugixml.hpp"

#include "docxcpp/opc_package.hpp"
#include "docxcpp/paragraph.hpp"
#include "docxcpp/table.hpp"

namespace docxcpp {

struct PictureSize {
  int width_pt{0};
  int height_pt{0};
};

struct PictureInfo {
  PictureSize size_pt;
  std::string relationship_id;
  std::string target;
  bool in_table_cell{false};
  std::size_t paragraph_index{0};
  std::size_t table_index{0};
  std::size_t row_index{0};
  std::size_t col_index{0};
};

struct PageSize {
  int width_pt{0};
  int height_pt{0};
};

struct PageMargins {
  int top_pt{0};
  int right_pt{0};
  int bottom_pt{0};
  int left_pt{0};
};

class Document {
public:
  Document();
  explicit Document(const std::filesystem::path& path);

  static Document open(const std::filesystem::path& path);

  Paragraph add_paragraph(const std::string& text);
  Paragraph add_paragraph(const std::vector<Run>& runs,
                          ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  Paragraph add_page_break();
  Paragraph add_styled_paragraph(const std::string& text, const RunStyle& style,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  Paragraph add_heading(const std::string& text, int level = 1);
  Paragraph add_heading(const std::vector<Run>& runs, int level = 1,
                        ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  Paragraph add_styled_heading(const std::string& text, int level, const RunStyle& style,
                               ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  Table add_table(std::size_t rows, std::size_t cols);
  void set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                      const std::string& text);
  void set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                      const std::vector<Run>& runs);
  Paragraph add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, const std::string& text);
  Paragraph add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, const std::vector<Run>& runs);
  void merge_table_cells(std::size_t table_index, std::size_t start_row, std::size_t start_col,
                         std::size_t end_row, std::size_t end_col);
  void set_paragraph_alignment(std::size_t paragraph_index, ParagraphAlignment alignment);
  void set_page_size_pt(int width_pt, int height_pt);
  void set_page_margins_pt(const PageMargins& margins);
  void add_picture(const std::filesystem::path& image_path);
  void add_picture(const std::filesystem::path& image_path, const PictureSize& size);
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path);
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path, const PictureSize& size);
  std::vector<Paragraph> paragraphs() const;
  std::vector<Table> tables() const;
  PageSize page_size_pt() const;
  PageMargins page_margins_pt() const;
  std::vector<PictureSize> picture_sizes_pt() const;
  std::vector<PictureInfo> pictures() const;
  std::size_t image_count() const;
  void save(const std::filesystem::path& path);

private:
  explicit Document(OpcPackage package);

  void load_document_xml();
  void store_document_xml();

  OpcPackage package_;
  std::unique_ptr<pugi::xml_document> xml_;
  mutable bool dirty_{false};
};

}  // namespace docxcpp
