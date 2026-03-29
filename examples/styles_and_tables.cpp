#include <filesystem>
#include <iostream>
#include <vector>

#include "docxcpp/document.hpp"

int main(int argc, char** argv) {
  namespace fs = std::filesystem;

  // 这个示例专门演示字符样式引用、表格样式引用和嵌套表格。
  const fs::path output =
      argc > 1 ? fs::path(argv[1]) : fs::path("docxcpp-styles-and-tables.docx");

  docxcpp::Document document;

  // 标题 run 通过字符样式名称引用内置样式。
  docxcpp::RunStyle title_style;
  title_style.bold = true;
  title_style.font_size_pt = 18;
  title_style.character_style_name = "Emphasis";
  document.add_styled_heading("样式与表格示例", 0, title_style,
                              docxcpp::ParagraphAlignment::Center);

  // 同一段中混用普通 run、字符样式 ID 和其他字符格式。
  docxcpp::RunStyle strong_style;
  strong_style.character_style_id = "Strong";
  strong_style.color_hex = "8A1538";

  docxcpp::RunStyle code_style;
  code_style.font_name = "Courier New";
  code_style.underline = true;

  std::vector<docxcpp::Run> runs{
      docxcpp::Run("这一段展示 ", {}),
      docxcpp::Run("字符样式引用", strong_style),
      docxcpp::Run(" 与 ", {}),
      docxcpp::Run("普通字符格式", code_style),
  };
  auto paragraph = document.add_paragraph(runs);

  // 已存在段落前插入一段，演示可变段落 API。
  paragraph.insert_paragraph_before("这是一段插入到前面的说明文字。");

  // 创建一个带样式的表格，并在后续修改为另一种表格样式。
  document.add_table(3, 3, "Light Shading Accent 1");
  document.set_table_style(0, "Medium Grid 1 Accent 1");

  document.set_table_cell(0, 0, 0, "列 A");
  document.set_table_cell(0, 0, 1, "列 B");
  document.set_table_cell(0, 0, 2, "列 C");
  document.set_table_cell(0, 1, 0, "普通文本");
  document.set_table_cell(0, 1, 1, runs);
  document.set_table_cell(0, 1, 2, "准备嵌套表格");
  document.set_table_cell(0, 2, 0, "第三行");
  document.set_table_cell(0, 2, 1, "合并区域");
  document.set_table_cell(0, 2, 2, "结束");

  // 合并两个单元格，便于观察读取时的 grid span 结果。
  document.merge_table_cells(0, 2, 1, 2, 2);

  // 在单元格中继续插入嵌套表格，并给子表应用另一种样式。
  document.add_nested_table(0, 1, 2, 1, 2, "Table Grid");
  document.set_nested_table_cell(0, 1, 2, 0, 0, 0, "子表 A");
  document.set_nested_table_cell(0, 1, 2, 0, 0, 1, "子表 B");

  if (output.has_parent_path()) {
    fs::create_directories(output.parent_path());
  }
  document.save(output);

  const auto paragraphs = document.paragraphs();
  const auto tables = document.tables();
  std::cout << "saved=" << output << '\n';
  std::cout << "paragraphs=" << paragraphs.size() << '\n';
  std::cout << "tables=" << tables.size() << '\n';
  if (!tables.empty()) {
    std::cout << "table[0].style_id=" << tables[0].style_id() << '\n';
    std::cout << "table[0].style_name=" << tables[0].style_name() << '\n';
  }

  return 0;
}
