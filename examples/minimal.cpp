#include <filesystem>
#include <fstream>
#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>

#include "docxcpp/document.hpp"

namespace {

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

  const fs::path output = argc > 1 ? fs::path(argv[1]) : fs::path("docxcpp-minimal-demo.docx");
  const fs::path image_path = fs::path("cpp") / "image.jpeg";
  const auto image_bytes = read_bytes(image_path);

  docxcpp::Document document;

  document.set_page_size_pt(595, 842);
  document.set_page_orientation(docxcpp::PageOrientation::Portrait);

  docxcpp::PageMargins margins;
  margins.top_pt = 72;
  margins.right_pt = 54;
  margins.bottom_pt = 72;
  margins.left_pt = 54;
  margins.header_pt = 24;
  margins.footer_pt = 24;
  document.set_page_margins_pt(margins);

  docxcpp::RunStyle title_style;
  title_style.bold = true;
  title_style.font_size_pt = 20;
  title_style.color_hex = "1F4E79";
  title_style.font_name = "Georgia";
  document.add_styled_heading("DocxCPP 快速示例", 0, title_style,
                              docxcpp::ParagraphAlignment::Center);

  document.add_paragraph("这个示例展示段落、可变 run、超链接、表格、嵌套表格、图片、批注和 section。");

  docxcpp::RunStyle emphasis_style;
  emphasis_style.bold = true;
  emphasis_style.color_hex = "8A1538";
  docxcpp::RunStyle code_style;
  code_style.font_name = "Courier New";
  code_style.underline = true;

  std::vector<docxcpp::Run> intro_runs{
      docxcpp::Run("这一段包含 ", {}),
      docxcpp::Run("加粗中文", emphasis_style),
      docxcpp::Run(" 与 ", {}),
      docxcpp::Run("带下划线字体", code_style),
  };
  document.add_paragraph(intro_runs, docxcpp::ParagraphAlignment::Left);

  auto mutable_paragraph = document.add_paragraph("这是一个可变段落");
  mutable_paragraph.add_run("，后面继续追加 run。", emphasis_style);
  document.add_comment(last_paragraph_index(document), "示例中的批注", "DocxCPP", "DX");

  docxcpp::RunStyle link_style;
  link_style.underline = true;
  link_style.color_hex = "0000EE";
  document.add_hyperlink("项目主页", "https://example.com/docxcpp-demo",
                         docxcpp::ParagraphAlignment::Right, link_style);

  const auto paragraph_index = last_paragraph_index(document) - 1;
  document.set_paragraph_indentation_pt(paragraph_index, 24, 0, -12);
  document.set_paragraph_spacing_pt(paragraph_index, 8, 12);
  document.set_paragraph_line_spacing_pt(paragraph_index, 18);

  document.add_heading("表格示例", 1);
  document.add_table(2, 3);
  document.set_table_cell(0, 0, 0, "表头");
  document.set_table_cell(0, 0, 1, "将被合并");
  document.set_table_cell(0, 0, 2, "插图");
  document.set_table_cell(0, 1, 0, "正文单元格");
  document.add_table_cell_paragraph(0, 1, 0, "同一单元格中的第二段。");
  document.add_hyperlink_to_table_cell(0, 1, 1, "单元格链接", "https://example.com/table",
                                       link_style);
  document.merge_table_cells(0, 0, 0, 0, 1);
  document.add_table_row(0);
  document.add_table_column(0);
  document.set_table_cell(0, 2, 0, "第三行第一列");
  document.set_table_cell(0, 2, 1, "嵌套表格");
  document.set_table_cell(0, 2, 2, "扩展列");
  document.set_table_cell(0, 1, 3, "新列内容");

  document.add_nested_table(0, 2, 1, 1, 2);
  document.set_nested_table_cell(0, 2, 1, 0, 0, 0, "子表 A");
  document.set_nested_table_cell(0, 2, 1, 0, 0, 1, "子表 B");

  docxcpp::PictureSize body_image_size;
  body_image_size.width_pt = 120;
  body_image_size.height_pt = 90;
  document.add_picture(image_path, body_image_size);

  docxcpp::PictureSize cell_image_size;
  cell_image_size.width_pt = 54;
  cell_image_size.height_pt = 40;
  document.add_picture_data_to_table_cell(0, 0, 2, image_bytes, "jpeg", cell_image_size,
                                          "table-buffer");

  document.add_section_break();
  document.set_header_text(0, "默认页眉");
  document.set_footer_text(0, "默认页脚");
  document.set_even_and_odd_headers(true);
  document.set_section_different_first_page(0, true);
  document.set_header_text(0, "首页页眉", docxcpp::HeaderFooterType::FirstPage);
  document.set_footer_text(0, "偶数页页脚", docxcpp::HeaderFooterType::EvenPage);
  document.add_paragraph("第二节继续追加内容，用于演示 section break 之后的普通写入。");

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
