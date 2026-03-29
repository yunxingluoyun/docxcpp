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
#include "docxcpp/section.hpp"
#include "docxcpp/table.hpp"

namespace docxcpp {

/**
 * @brief 图片尺寸，单位为 pt。
 */
struct PictureSize {
  int width_pt{0};
  int height_pt{0};
};

/**
 * @brief 图片读取结果，包含位置与关系信息。
 */
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

/**
 * @brief 批注读取结果。
 */
struct CommentInfo {
  std::size_t id{0};
  std::string text;
  std::string author;
  std::string initials;
};

class Document {
public:
  /**
   * @brief 创建一个基于默认模板的空文档。
   */
  Document();
  /**
   * @brief 从现有 docx 文件加载文档。
   * @param path docx 文件路径。
   */
  explicit Document(const std::filesystem::path& path);

  /**
   * @brief 以静态工厂形式打开文档。
   * @param path docx 文件路径。
   * @return 已加载的文档对象。
   */
  static Document open(const std::filesystem::path& path);

  /**
   * @brief 追加一个纯文本段落。
   * @param text 段落文本。
   * @return 新增段落。
   */
  Paragraph add_paragraph(const std::string& text);
  /**
   * @brief 追加一个由多个 runs 组成的段落。
   * @param runs 段落 runs。
   * @param alignment 对齐方式。
   * @return 新增段落。
   */
  Paragraph add_paragraph(const std::vector<Run>& runs,
                          ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  /**
   * @brief 追加一个显式分页符段落。
   * @return 新增段落。
   */
  Paragraph add_page_break();
  /**
   * @brief 追加一个带统一字符样式的段落。
   * @param text 段落文本。
   * @param style 字符样式。
   * @param alignment 对齐方式。
   * @return 新增段落。
   */
  Paragraph add_styled_paragraph(const std::string& text, const RunStyle& style,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  /**
   * @brief 追加一个标题段落。
   * @param text 标题文本。
   * @param level 标题级别。
   * @return 新增段落。
   */
  Paragraph add_heading(const std::string& text, int level = 1);
  /**
   * @brief 追加一个由多个 runs 构成的标题段落。
   * @param runs 标题 runs。
   * @param level 标题级别。
   * @param alignment 对齐方式。
   * @return 新增段落。
   */
  Paragraph add_heading(const std::vector<Run>& runs, int level = 1,
                        ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  /**
   * @brief 追加一个带统一字符样式的标题段落。
   * @param text 标题文本。
   * @param level 标题级别。
   * @param style 字符样式。
   * @param alignment 对齐方式。
   * @return 新增段落。
   */
  Paragraph add_styled_heading(const std::string& text, int level, const RunStyle& style,
                               ParagraphAlignment alignment = ParagraphAlignment::Inherit);
  /**
   * @brief 追加一个顶层表格。
   * @param rows 行数。
   * @param cols 列数。
   * @return 新增表格。
   */
  Table add_table(std::size_t rows, std::size_t cols);
  /**
   * @brief 在指定单元格中追加一个嵌套表格。
   * @param table_index 顶层表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param rows 嵌套表格行数。
   * @param cols 嵌套表格列数。
   * @return 新增嵌套表格。
   */
  Table add_nested_table(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                         std::size_t rows, std::size_t cols);
  /**
   * @brief 为顶层表格末尾追加一行。
   * @param table_index 顶层表格索引。
   */
  void add_table_row(std::size_t table_index);
  /**
   * @brief 为顶层表格末尾追加一列。
   * @param table_index 顶层表格索引。
   */
  void add_table_column(std::size_t table_index);
  /**
   * @brief 设置顶层表格单元格的文本内容。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param text 单元格文本。
   */
  void set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                      const std::string& text);
  /**
   * @brief 设置顶层表格单元格的 run 内容。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param runs 单元格 runs。
   */
  void set_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                      const std::vector<Run>& runs);
  /**
   * @brief 设置嵌套表格单元格的文本内容。
   * @param table_index 顶层表格索引。
   * @param row_index 顶层行索引。
   * @param col_index 顶层列索引。
   * @param nested_table_index 嵌套表格索引。
   * @param nested_row_index 嵌套行索引。
   * @param nested_col_index 嵌套列索引。
   * @param text 单元格文本。
   */
  void set_nested_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                             std::size_t nested_table_index, std::size_t nested_row_index,
                             std::size_t nested_col_index, const std::string& text);
  /**
   * @brief 设置嵌套表格单元格的 run 内容。
   * @param table_index 顶层表格索引。
   * @param row_index 顶层行索引。
   * @param col_index 顶层列索引。
   * @param nested_table_index 嵌套表格索引。
   * @param nested_row_index 嵌套行索引。
   * @param nested_col_index 嵌套列索引。
   * @param runs 单元格 runs。
   */
  void set_nested_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                             std::size_t nested_table_index, std::size_t nested_row_index,
                             std::size_t nested_col_index, const std::vector<Run>& runs);
  /**
   * @brief 在顶层表格单元格末尾追加一个纯文本段落。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param text 段落文本。
   * @return 新增段落。
   */
  Paragraph add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, const std::string& text);
  /**
   * @brief 在顶层表格单元格末尾追加一个由 runs 组成的段落。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param runs 段落 runs。
   * @return 新增段落。
   */
  Paragraph add_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                     std::size_t col_index, const std::vector<Run>& runs);
  /**
   * @brief 在嵌套表格单元格末尾追加一个纯文本段落。
   * @param table_index 顶层表格索引。
   * @param row_index 顶层行索引。
   * @param col_index 顶层列索引。
   * @param nested_table_index 嵌套表格索引。
   * @param nested_row_index 嵌套行索引。
   * @param nested_col_index 嵌套列索引。
   * @param text 段落文本。
   * @return 新增段落。
   */
  Paragraph add_nested_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                            std::size_t col_index, std::size_t nested_table_index,
                                            std::size_t nested_row_index,
                                            std::size_t nested_col_index, const std::string& text);
  /**
   * @brief 在嵌套表格单元格末尾追加一个由 runs 组成的段落。
   * @param table_index 顶层表格索引。
   * @param row_index 顶层行索引。
   * @param col_index 顶层列索引。
   * @param nested_table_index 嵌套表格索引。
   * @param nested_row_index 嵌套行索引。
   * @param nested_col_index 嵌套列索引。
   * @param runs 段落 runs。
   * @return 新增段落。
   */
  Paragraph add_nested_table_cell_paragraph(std::size_t table_index, std::size_t row_index,
                                            std::size_t col_index, std::size_t nested_table_index,
                                            std::size_t nested_row_index,
                                            std::size_t nested_col_index,
                                            const std::vector<Run>& runs);
  /**
   * @brief 合并顶层表格中的矩形单元格区域。
   * @param table_index 表格索引。
   * @param start_row 起始行。
   * @param start_col 起始列。
   * @param end_row 结束行。
   * @param end_col 结束列。
   */
  void merge_table_cells(std::size_t table_index, std::size_t start_row, std::size_t start_col,
                         std::size_t end_row, std::size_t end_col);
  /**
   * @brief 设置指定段落的对齐方式。
   * @param paragraph_index 段落索引。
   * @param alignment 对齐方式。
   */
  void set_paragraph_alignment(std::size_t paragraph_index, ParagraphAlignment alignment);
  /**
   * @brief 设置段落缩进，单位为 pt。
   * @param paragraph_index 段落索引。
   * @param left_indent_pt 左缩进。
   * @param right_indent_pt 右缩进。
   * @param first_line_indent_pt 首行缩进。
   */
  void set_paragraph_indentation_pt(std::size_t paragraph_index, std::optional<int> left_indent_pt,
                                    std::optional<int> right_indent_pt,
                                    std::optional<int> first_line_indent_pt);
  /**
   * @brief 设置段前段后间距，单位为 pt。
   * @param paragraph_index 段落索引。
   * @param space_before_pt 段前间距。
   * @param space_after_pt 段后间距。
   */
  void set_paragraph_spacing_pt(std::size_t paragraph_index, std::optional<int> space_before_pt,
                                std::optional<int> space_after_pt);
  /**
   * @brief 以固定值模式设置行距，单位为 pt。
   * @param paragraph_index 段落索引。
   * @param line_spacing_pt 行距值。
   */
  void set_paragraph_line_spacing_pt(std::size_t paragraph_index,
                                     std::optional<int> line_spacing_pt);
  /**
   * @brief 按指定模式设置段落行距。
   * @param paragraph_index 段落索引。
   * @param value 行距值。
   * @param mode 行距模式。
   */
  void set_paragraph_line_spacing(std::size_t paragraph_index, std::optional<double> value,
                                  LineSpacingMode mode = LineSpacingMode::Exact);
  /**
   * @brief 设置段落制表位。
   * @param paragraph_index 段落索引。
   * @param tab_stops 制表位列表。
   */
  void set_paragraph_tab_stops(std::size_t paragraph_index, const std::vector<TabStop>& tab_stops);
  /**
   * @brief 设置段落分页相关标记。
   * @param paragraph_index 段落索引。
   * @param keep_together 是否段内不拆分。
   * @param keep_with_next 是否与下一段同页。
   * @param page_break_before 是否段前分页。
   */
  void set_paragraph_pagination(std::size_t paragraph_index, std::optional<bool> keep_together,
                                std::optional<bool> keep_with_next,
                                std::optional<bool> page_break_before);
  /**
   * @brief 设置当前节页面尺寸，单位为 pt。
   * @param width_pt 页面宽度。
   * @param height_pt 页面高度。
   */
  void set_page_size_pt(int width_pt, int height_pt);
  /**
   * @brief 设置当前节页面方向。
   * @param orientation 页面方向。
   */
  void set_page_orientation(PageOrientation orientation);
  /**
   * @brief 设置当前节页边距与页眉页脚距离。
   * @param margins 页边距配置。
   */
  void set_page_margins_pt(const PageMargins& margins);
  /**
   * @brief 在文档末尾插入一个 section break，并继承当前节设置。
   */
  void add_section_break();
  /**
   * @brief 在文档末尾插入一个 section break，并指定下一节设置。
   * @param next_section 下一节配置。
   */
  void add_section_break(const Section& next_section);
  /**
   * @brief 设置文档是否启用奇偶页不同页眉页脚。
   * @param enabled 是否启用。
   */
  void set_even_and_odd_headers(bool enabled);
  /**
   * @brief 返回文档是否启用奇偶页不同页眉页脚。
   * @return 启用时为 true。
   */
  bool even_and_odd_headers() const;
  /**
   * @brief 设置某个节是否启用首页不同页眉页脚。
   * @param section_index 节索引。
   * @param enabled 是否启用。
   */
  void set_section_different_first_page(std::size_t section_index, bool enabled);
  /**
   * @brief 返回某个节是否启用首页不同页眉页脚。
   * @param section_index 节索引。
   * @return 启用时为 true。
   */
  bool section_different_first_page(std::size_t section_index) const;
  /**
   * @brief 清空某个节的默认页眉内容。
   * @param section_index 节索引。
   */
  void clear_header(std::size_t section_index);
  /**
   * @brief 清空某个节的默认页脚内容。
   * @param section_index 节索引。
   */
  void clear_footer(std::size_t section_index);
  /**
   * @brief 设置某个节指定类型页眉的纯文本内容。
   * @param section_index 节索引。
   * @param text 页眉文本。
   * @param type 页眉类型。
   */
  void set_header_text(std::size_t section_index, const std::string& text,
                       HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 设置某个节指定类型页脚的纯文本内容。
   * @param section_index 节索引。
   * @param text 页脚文本。
   * @param type 页脚类型。
   */
  void set_footer_text(std::size_t section_index, const std::string& text,
                       HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页眉末尾追加一个纯文本段落。
   * @param section_index 节索引。
   * @param text 段落文本。
   * @param alignment 对齐方式。
   * @param type 页眉类型。
   * @return 新增段落。
   */
  Paragraph add_header_paragraph(std::size_t section_index, const std::string& text,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                 HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页眉末尾追加一个由 runs 组成的段落。
   * @param section_index 节索引。
   * @param runs 段落 runs。
   * @param alignment 对齐方式。
   * @param type 页眉类型。
   * @return 新增段落。
   */
  Paragraph add_header_paragraph(std::size_t section_index, const std::vector<Run>& runs,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                 HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页眉末尾追加一个带统一字符样式的段落。
   * @param section_index 节索引。
   * @param text 段落文本。
   * @param style 字符样式。
   * @param alignment 对齐方式。
   * @param type 页眉类型。
   * @return 新增段落。
   */
  Paragraph add_styled_header_paragraph(std::size_t section_index, const std::string& text,
                                        const RunStyle& style,
                                        ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                        HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页脚末尾追加一个纯文本段落。
   * @param section_index 节索引。
   * @param text 段落文本。
   * @param alignment 对齐方式。
   * @param type 页脚类型。
   * @return 新增段落。
   */
  Paragraph add_footer_paragraph(std::size_t section_index, const std::string& text,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                 HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页脚末尾追加一个由 runs 组成的段落。
   * @param section_index 节索引。
   * @param runs 段落 runs。
   * @param alignment 对齐方式。
   * @param type 页脚类型。
   * @return 新增段落。
   */
  Paragraph add_footer_paragraph(std::size_t section_index, const std::vector<Run>& runs,
                                 ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                 HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向页脚末尾追加一个带统一字符样式的段落。
   * @param section_index 节索引。
   * @param text 段落文本。
   * @param style 字符样式。
   * @param alignment 对齐方式。
   * @param type 页脚类型。
   * @return 新增段落。
   */
  Paragraph add_styled_footer_paragraph(std::size_t section_index, const std::string& text,
                                        const RunStyle& style,
                                        ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                                        HeaderFooterType type = HeaderFooterType::Default);
  /**
   * @brief 向文档正文追加图片。
   * @param image_path 图片路径。
   */
  void add_picture(const std::filesystem::path& image_path);
  /**
   * @brief 向文档正文追加指定尺寸的图片。
   * @param image_path 图片路径。
   * @param size 图片尺寸。
   */
  void add_picture(const std::filesystem::path& image_path, const PictureSize& size);
  /**
   * @brief 以字节流形式向文档正文追加图片。
   * @param image_bytes 图片字节流。
   * @param extension 文件扩展名。
   */
  void add_picture_data(const std::vector<std::uint8_t>& image_bytes, const std::string& extension);
  /**
   * @brief 以字节流形式向文档正文追加指定尺寸的图片。
   * @param image_bytes 图片字节流。
   * @param extension 文件扩展名。
   * @param size 图片尺寸。
   * @param filename_hint 文件名提示。
   */
  void add_picture_data(const std::vector<std::uint8_t>& image_bytes, const std::string& extension,
                        const PictureSize& size, const std::string& filename_hint = "image");
  /**
   * @brief 向顶层表格单元格追加图片。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param image_path 图片路径。
   */
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path);
  /**
   * @brief 向顶层表格单元格追加指定尺寸的图片。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param image_path 图片路径。
   * @param size 图片尺寸。
   */
  void add_picture_to_table_cell(std::size_t table_index, std::size_t row_index, std::size_t col_index,
                                 const std::filesystem::path& image_path, const PictureSize& size);
  /**
   * @brief 以字节流形式向顶层表格单元格追加图片。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param image_bytes 图片字节流。
   * @param extension 文件扩展名。
   * @param filename_hint 文件名提示。
   */
  void add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                      std::size_t col_index,
                                      const std::vector<std::uint8_t>& image_bytes,
                                      const std::string& extension, const std::string& filename_hint = "image");
  /**
   * @brief 以字节流形式向顶层表格单元格追加指定尺寸的图片。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param image_bytes 图片字节流。
   * @param extension 文件扩展名。
   * @param size 图片尺寸。
   * @param filename_hint 文件名提示。
   */
  void add_picture_data_to_table_cell(std::size_t table_index, std::size_t row_index,
                                      std::size_t col_index,
                                      const std::vector<std::uint8_t>& image_bytes,
                                      const std::string& extension, const PictureSize& size,
                                      const std::string& filename_hint = "image");
  /**
   * @brief 向正文追加一个超链接段落。
   * @param text 显示文本。
   * @param url 目标地址。
   * @param alignment 对齐方式。
   * @param style 链接样式。
   * @return 新增段落。
   */
  Paragraph add_hyperlink(const std::string& text, const std::string& url,
                          ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                          const RunStyle& style = {});
  /**
   * @brief 向顶层表格单元格追加一个超链接段落。
   * @param table_index 表格索引。
   * @param row_index 行索引。
   * @param col_index 列索引。
   * @param text 显示文本。
   * @param url 目标地址。
   * @param style 链接样式。
   * @return 新增段落。
   */
  Paragraph add_hyperlink_to_table_cell(std::size_t table_index, std::size_t row_index,
                                        std::size_t col_index, const std::string& text,
                                        const std::string& url, const RunStyle& style = {});
  /**
   * @brief 为指定正文段落添加批注。
   * @param paragraph_index 段落索引。
   * @param text 批注文本。
   * @param author 作者。
   * @param initials 首字母。
   * @return 新增批注的 ID。
   */
  std::size_t add_comment(std::size_t paragraph_index, const std::string& text,
                          const std::string& author = {}, const std::string& initials = {});
  /**
   * @brief 返回正文段落列表。
   * @return 段落列表。
   */
  std::vector<Paragraph> paragraphs() const;
  /**
   * @brief 返回顶层表格列表。
   * @return 顶层表格列表。
   */
  std::vector<Table> tables() const;
  /**
   * @brief 返回文档中的 section 列表。
   * @return section 列表。
   */
  std::vector<Section> sections() const;
  /**
   * @brief 返回最后一个 section 的页面尺寸。
   * @return 页面尺寸。
   */
  PageSize page_size_pt() const;
  /**
   * @brief 返回最后一个 section 的页面方向。
   * @return 页面方向。
   */
  PageOrientation page_orientation() const;
  /**
   * @brief 返回最后一个 section 的页边距配置。
   * @return 页边距配置。
   */
  PageMargins page_margins_pt() const;
  /**
   * @brief 返回文档中所有批注。
   * @return 批注列表。
   */
  std::vector<CommentInfo> comments() const;
  /**
   * @brief 返回批注数量。
   * @return 批注数量。
   */
  std::size_t comment_count() const;
  /**
   * @brief 返回各节指定类型页眉的纯文本汇总。
   * @param type 页眉类型。
   * @return 页眉文本列表。
   */
  std::vector<std::string> headers(HeaderFooterType type = HeaderFooterType::Default) const;
  /**
   * @brief 返回各节指定类型页脚的纯文本汇总。
   * @param type 页脚类型。
   * @return 页脚文本列表。
   */
  std::vector<std::string> footers(HeaderFooterType type = HeaderFooterType::Default) const;
  /**
   * @brief 返回某个节指定类型页眉中的段落列表。
   * @param section_index 节索引。
   * @param type 页眉类型。
   * @return 段落列表。
   */
  std::vector<Paragraph> header_paragraphs(std::size_t section_index,
                                           HeaderFooterType type = HeaderFooterType::Default) const;
  /**
   * @brief 返回某个节指定类型页脚中的段落列表。
   * @param section_index 节索引。
   * @param type 页脚类型。
   * @return 段落列表。
   */
  std::vector<Paragraph> footer_paragraphs(std::size_t section_index,
                                           HeaderFooterType type = HeaderFooterType::Default) const;
  /**
   * @brief 返回文档中所有图片尺寸。
   * @return 图片尺寸列表。
   */
  std::vector<PictureSize> picture_sizes_pt() const;
  /**
   * @brief 返回图片的详细位置信息。
   * @return 图片信息列表。
   */
  std::vector<PictureInfo> pictures() const;
  /**
   * @brief 返回图片数量。
   * @return 图片数量。
   */
  std::size_t image_count() const;
  /**
   * @brief 将文档保存到指定路径。
   * @param path 输出路径。
   */
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
