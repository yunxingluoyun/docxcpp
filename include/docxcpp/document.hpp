#pragma once

#include <cstdint>
#include <filesystem>
#include <memory>
#include <optional>
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

struct CommentInfo {
  std::size_t id{0};
  std::string text;
  std::string author;
  std::string initials;
};

struct PageSize {
  int width_pt{0};
  int height_pt{0};
};

enum class PageOrientation {
  Portrait,
  Landscape,
};

struct PageMargins {
  int top_pt{0};
  int right_pt{0};
  int bottom_pt{0};
  int left_pt{0};
  int header_pt{0};
  int footer_pt{0};
  int gutter_pt{0};
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
  void set_paragraph_indentation_pt(std::size_t paragraph_index, std::optional<int> left_indent_pt,
                                    std::optional<int> right_indent_pt,
                                    std::optional<int> first_line_indent_pt);
  void set_paragraph_spacing_pt(std::size_t paragraph_index, std::optional<int> space_before_pt,
                                std::optional<int> space_after_pt);
  void set_paragraph_line_spacing_pt(std::size_t paragraph_index,
                                     std::optional<int> line_spacing_pt);
  void set_paragraph_pagination(std::size_t paragraph_index, std::optional<bool> keep_together,
                                std::optional<bool> keep_with_next,
                                std::optional<bool> page_break_before);
  void set_page_size_pt(int width_pt, int height_pt);
  void set_page_orientation(PageOrientation orientation);
  void set_page_margins_pt(const PageMargins& margins);
  void add_picture(const std::filesystem::path& image_path);
  void add_picture(const std::filesystem::path& image_path, const PictureSize& size);
  void add_picture_data(const std::vector<std::uint8_t>& image_bytes, const std::string& extension);
  void add_picture_data(const std::vector<std::uint8_t>& image_bytes, const std::string& extension,
                        const PictureSize& size, const std::string& filename_hint = "image");
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path);
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path, const PictureSize& size);
  void add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                      std::size_t col_index,
                                      const std::vector<std::uint8_t>& image_bytes,
                                      const std::string& extension, const std::string& filename_hint = "image");
  void add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                      std::size_t col_index,
                                      const std::vector<std::uint8_t>& image_bytes,
                                      const std::string& extension, const PictureSize& size,
                                      const std::string& filename_hint = "image");
  Paragraph add_hyperlink(const std::string& text, const std::string& url,
                          ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                          const RunStyle& style = {});
  Paragraph add_hyperlink_to_table_cell(std::size_t table_index, std::size_t row_index,
                                        std::size_t col_index, const std::string& text,
                                        const std::string& url, const RunStyle& style = {});
  std::size_t add_comment(std::size_t paragraph_index, const std::string& text,
                          const std::string& author = {}, const std::string& initials = {});
  std::vector<Paragraph> paragraphs() const;
  std::vector<Table> tables() const;
  PageSize page_size_pt() const;
  PageOrientation page_orientation() const;
  PageMargins page_margins_pt() const;
  std::vector<CommentInfo> comments() const;
  std::size_t comment_count() const;
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
