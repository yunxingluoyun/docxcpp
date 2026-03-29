#include <cstdint>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <optional>
#include <stdexcept>
#include <string>
#include <vector>

#include "docxcpp/document.hpp"

namespace {

// 读取图片文件，供正文和表格单元格中的图片示例复用。
std::vector<std::uint8_t> read_bytes(const std::filesystem::path& path) {
  std::ifstream stream(path, std::ios::binary);
  if (!stream) {
    throw std::runtime_error("failed to open image file: " + path.string());
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

// 返回最后一个段落的索引，便于给刚写入的段落继续配置格式或批注。
std::size_t last_paragraph_index(const docxcpp::Document& document) {
  const auto paragraphs = document.paragraphs();
  if (paragraphs.empty()) {
    throw std::runtime_error("document has no paragraphs");
  }
  return paragraphs.size() - 1;
}

}  // namespace

int main(int argc, char** argv) {
  namespace fs = std::filesystem;

  const fs::path output =
      argc > 1 ? fs::path(argv[1]) : fs::path("full-feature-5-pages.docx");
  const fs::path image_path = fs::path("cpp") / "image.jpeg";
  const auto image_bytes = read_bytes(image_path);

  docxcpp::Document document;

  // 第 1 页：页面设置与首页内容。
  document.set_page_size(docxcpp::StandardPageSize::Letter);
  document.set_page_orientation(docxcpp::PageOrientation::Landscape);

  docxcpp::PageMargins margins;
  margins.top = docxcpp::Length::from_pt(72);
  margins.right = docxcpp::Length::from_pt(72);
  margins.bottom = docxcpp::Length::from_pt(72);
  margins.left = docxcpp::Length::from_pt(72);
  margins.header = docxcpp::Length::from_pt(30);
  margins.footer = docxcpp::Length::from_pt(30);
  margins.gutter = docxcpp::Length::from_pt(12);
  document.set_page_margins(margins);

  docxcpp::RunStyle title_style;
  title_style.bold = true;
  title_style.font_size_pt = 22;
  title_style.color_hex = "0F4C81";
  title_style.font_name = "Georgia";
  title_style.highlight = "yellow";
  document.add_styled_heading("DocxCPP 全功能中文示例", 0, title_style,
                              docxcpp::ParagraphAlignment::Center);

  document.add_paragraph(
      "第 1 页从页面设置开始，覆盖标题、混合 runs、可变段落、超链接和批注等能力。");

  docxcpp::RunStyle strong_style;
  strong_style.bold = true;
  strong_style.color_hex = "8A1538";

  docxcpp::RunStyle code_style;
  code_style.font_name = "Courier New";
  code_style.underline = true;

  std::vector<docxcpp::Run> mixed_runs{
      docxcpp::Run("这是", strong_style),
      docxcpp::Run("混合", code_style),
      docxcpp::Run(" run 段落，位于第 1 页。"),
  };
  document.add_paragraph(mixed_runs, docxcpp::ParagraphAlignment::Left);

  auto mutable_paragraph = document.add_paragraph("这是可变段落");
  mutable_paragraph.add_run("，后面追加了一段中文 run", strong_style);
  document.add_comment(last_paragraph_index(document), "这是针对可变段落的批注", "DocxCPP", "DX");

  docxcpp::RunStyle hyperlink_style;
  hyperlink_style.underline = true;
  hyperlink_style.color_hex = "0000EE";
  document.add_hyperlink("项目主页", "https://example.com/docxcpp-demo",
                         docxcpp::ParagraphAlignment::Right, hyperlink_style);

  // 第 2 页：段落格式、分页控制和样式化正文。
  document.add_page_break();

  std::vector<docxcpp::Run> heading_runs{
      docxcpp::Run("第 2 页", strong_style),
      docxcpp::Run(" 段落格式", code_style),
  };
  document.add_heading(heading_runs, 2, docxcpp::ParagraphAlignment::Center);

  docxcpp::RunStyle body_style;
  body_style.italic = true;
  body_style.font_size_pt = 12;
  body_style.color_hex = "336B87";
  body_style.highlight = "green";
  document.add_styled_paragraph(
      "这个段落用于验证带样式的中文文本写入，以及读取时的对齐和正文提取。",
      body_style, docxcpp::ParagraphAlignment::Justify);

  auto formatted = document.add_paragraph("这个段落用于验证缩进、段前段后间距和分页控制。");
  formatted.add_run(" 它还演示了在现有段落后继续追加中文 run。", code_style);
  const std::size_t formatted_index = last_paragraph_index(document);
  document.set_paragraph_alignment(formatted_index, docxcpp::ParagraphAlignment::Center);
  document.set_paragraph_indentation_pt(formatted_index, 24, 18, -12);
  document.set_paragraph_spacing_pt(formatted_index, 10, 14);
  document.set_paragraph_line_spacing_pt(formatted_index, 20);
  document.set_paragraph_pagination(formatted_index, true, false, std::nullopt);

  // 第 3 页：表格、超链接、正文图片、单元格图片。
  document.add_page_break();

  document.add_heading("第 3 页 表格与图片", 1);
  document.add_table(2, 3);
  document.set_table_cell(0, 0, 0, "合并表头");
  document.set_table_cell(0, 0, 1, "将被合并隐藏");
  document.set_table_cell(0, 0, 2, "图片占位单元格");
  document.set_table_cell(0, 1, 0, mixed_runs);
  document.add_table_cell_paragraph(0, 1, 0, "单元格中的第二段文字");
  document.add_table_cell_paragraph(0, 1, 1, heading_runs);
  document.add_hyperlink_to_table_cell(0, 1, 1, "单元格超链接", "https://example.com/table-link",
                                       hyperlink_style);

  docxcpp::PictureSize wide_picture;
  wide_picture.width_pt = 160;
  wide_picture.height_pt = 120;
  document.add_picture(image_path, wide_picture);

  docxcpp::PictureSize medium_picture;
  medium_picture.width_pt = 144;
  medium_picture.height_pt = 108;
  document.add_picture_data(image_bytes, "jpeg", medium_picture, "buffer-inline");

  docxcpp::PictureSize cell_picture;
  cell_picture.width_pt = 72;
  cell_picture.height_pt = 54;
  document.add_picture_to_table_cell(0, 1, 2, image_path, cell_picture);
  document.add_picture_data_to_table_cell(0, 0, 2, image_bytes, "jpeg", cell_picture,
                                          "buffer-cell");
  document.merge_table_cells(0, 0, 0, 0, 1);

  // 第 4 页：page_break_before 和普通正文的组合。
  auto page4_intro =
      document.add_paragraph("第 4 页通过当前段落上的 page_break_before 生成。");
  page4_intro.add_run(" 这能验证分页相关 API 在中文文本下也能正常工作。", body_style);
  const std::size_t page4_intro_index = last_paragraph_index(document);
  document.set_paragraph_pagination(page4_intro_index, false, true, true);

  document.add_paragraph(
      "这一页继续追加普通正文，用来确认强制分页后的内容顺序保持正确。");

  // 第 5 页：显式分页后的总结页面。
  document.add_page_break();

  document.add_heading("第 5 页 总结", 1);
  document.add_paragraph(
      "第 5 页用于收尾，并同时验证显式分页符与 page_break_before 的组合效果。");

  std::vector<docxcpp::Run> closing_runs{
      docxcpp::Run("已覆盖功能：", strong_style),
      docxcpp::Run("标题、段落、runs、段落格式、表格、图片、超链接、批注和页面设置。",
                   body_style),
  };
  document.add_paragraph(closing_runs, docxcpp::ParagraphAlignment::Left);

  if (output.has_parent_path()) {
    fs::create_directories(output.parent_path());
  }
  document.save(output);

  std::cout << "saved=" << output << '\n';
  std::cout << "paragraphs=" << document.paragraphs().size() << '\n';
  std::cout << "tables=" << document.tables().size() << '\n';
  std::cout << "comments=" << document.comment_count() << '\n';
  std::cout << "images=" << document.image_count() << '\n';
  std::cout << "sections=" << document.sections().size() << '\n';

  return 0;
}
