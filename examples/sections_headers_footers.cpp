#include <filesystem>
#include <iostream>

#include "docxcpp/document.hpp"

int main(int argc, char** argv) {
  namespace fs = std::filesystem;

  // 这个示例专门演示页面设置、分节和页眉页脚。
  const fs::path output =
      argc > 1 ? fs::path(argv[1]) : fs::path("docxcpp-sections-headers-footers.docx");

  docxcpp::Document document;

  // 第一节使用 A4 竖版，边距混合使用英寸、厘米和点。
  document.set_page_size(docxcpp::StandardPageSize::A4);
  document.set_page_orientation(docxcpp::PageOrientation::Portrait);
  document.set_page_margins(docxcpp::PageMargins{
      docxcpp::Inches(1.0),
      docxcpp::Cm(2.0),
      docxcpp::Inches(1.0),
      docxcpp::Cm(2.0),
      docxcpp::Pt(24),
      docxcpp::Pt(24),
      docxcpp::Twips(0),
  });

  // 开启奇偶页不同和首页不同。
  document.set_even_and_odd_headers(true);
  document.set_section_different_first_page(0, true);

  // 写入第一页的页眉页脚内容。
  document.set_header_text(0, "第一节默认页眉");
  document.set_footer_text(0, "第一节默认页脚");
  document.set_header_text(0, "第一节首页页眉", docxcpp::HeaderFooterType::FirstPage);
  document.set_footer_text(0, "第一节偶数页页脚", docxcpp::HeaderFooterType::EvenPage);

  docxcpp::RunStyle header_style;
  header_style.bold = true;
  header_style.color_hex = "1F4E79";
  document.add_styled_header_paragraph(0, "第一节页眉中的样式化段落", header_style,
                                      docxcpp::ParagraphAlignment::Center);
  document.add_footer_paragraph(0, "第一节页脚中的附加段落",
                                docxcpp::ParagraphAlignment::Right);

  document.add_heading("第一节", 1);
  document.add_paragraph("这一节适合展示页面设置、首页不同和奇偶页不同。");

  // 第二节切换为 Letter 横版，演示不同节之间的页面设置差异。
  document.add_section_break(docxcpp::Section(
      docxcpp::StandardPageSize::Letter, docxcpp::PageOrientation::Landscape,
      docxcpp::PageMargins{
          docxcpp::Inches(1.25),
          docxcpp::Inches(0.75),
          docxcpp::Inches(1.0),
          docxcpp::Inches(0.75),
          docxcpp::Pt(18),
          docxcpp::Pt(18),
          docxcpp::Twips(180),
      }));

  document.set_header_text(1, "第二节默认页眉");
  document.set_footer_text(1, "第二节默认页脚");
  document.set_header_text(1, "第二节偶数页页眉", docxcpp::HeaderFooterType::EvenPage);

  document.add_heading("第二节", 1);
  document.add_paragraph("第二节改为横向页面，适合放宽表格或报表。");
  document.add_paragraph("示例结束后，可以用 Word 或 docxcpp 自己的读取接口检查 section 信息。");

  if (output.has_parent_path()) {
    fs::create_directories(output.parent_path());
  }
  document.save(output);

  std::cout << "saved=" << output << '\n';
  std::cout << "sections=" << document.sections().size() << '\n';
  std::cout << "headers=" << document.headers().size() << '\n';
  std::cout << "footers=" << document.footers().size() << '\n';

  return 0;
}
