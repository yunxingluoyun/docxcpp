#include <filesystem>
#include <iostream>
#include <string>

#include "docxcpp/document.hpp"

int main(int argc, char** argv) {
  namespace fs = std::filesystem;

  // 读取示例需要传入一个现有 docx 路径。
  if (argc < 2) {
    std::cerr << "用法: " << argv[0] << " <input.docx>\n";
    return 1;
  }

  const fs::path input = fs::path(argv[1]);
  docxcpp::Document document(input);

  // 读取文档中的主要统计信息。
  const auto paragraphs = document.paragraphs();
  const auto tables = document.tables();
  const auto sections = document.sections();
  const auto comments = document.comments();
  const auto pictures = document.pictures();

  std::cout << "input=" << input << '\n';
  std::cout << "paragraphs=" << paragraphs.size() << '\n';
  std::cout << "tables=" << tables.size() << '\n';
  std::cout << "sections=" << sections.size() << '\n';
  std::cout << "comments=" << comments.size() << '\n';
  std::cout << "images=" << pictures.size() << '\n';

  // 打印前几个段落，便于快速查看正文和样式信息。
  const std::size_t preview_count = std::min<std::size_t>(paragraphs.size(), 5);
  for (std::size_t i = 0; i < preview_count; ++i) {
    const auto& paragraph = paragraphs[i];
    std::cout << "paragraph[" << i << "].text=" << paragraph.text() << '\n';
    std::cout << "paragraph[" << i << "].style_id=" << paragraph.style_id() << '\n';
    std::cout << "paragraph[" << i << "].runs=" << paragraph.runs().size() << '\n';
  }

  // 如果有表格，输出第一个表格的行列信息和样式信息。
  if (!tables.empty()) {
    const auto& table = tables.front();
    std::cout << "table[0].rows=" << table.row_count() << '\n';
    std::cout << "table[0].cols=" << table.column_count() << '\n';
    std::cout << "table[0].style_id=" << table.style_id() << '\n';
    std::cout << "table[0].style_name=" << table.style_name() << '\n';
  }

  // 如果有批注，打印第一条批注的作者和内容。
  if (!comments.empty()) {
    std::cout << "comment[0].author=" << comments.front().author << '\n';
    std::cout << "comment[0].text=" << comments.front().text << '\n';
  }

  // 如果有图片，打印第一张图片的关联关系和位置。
  if (!pictures.empty()) {
    const auto& picture = pictures.front();
    std::cout << "image[0].target=" << picture.target << '\n';
    std::cout << "image[0].relationship_id=" << picture.relationship_id << '\n';
    std::cout << "image[0].in_table_cell=" << picture.in_table_cell << '\n';
  }

  return 0;
}
