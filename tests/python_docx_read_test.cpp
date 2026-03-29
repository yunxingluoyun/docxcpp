#include <algorithm>
#include <cassert>
#include <cmath>
#include <filesystem>
#include <string>
#include <vector>

#include "docxcpp/document.hpp"

int main() {
  namespace fs = std::filesystem;

  const fs::path artifact_dir = DOCXCPP_TEST_ARTIFACT_DIR;
  const fs::path input = artifact_dir / "python-docx-generated.docx";

  docxcpp::Document document = docxcpp::Document::open(input);

  const auto paragraphs = document.paragraphs();
  const auto tables = document.tables();
  const auto sections = document.sections();
  const auto comments = document.comments();
  const auto headers = document.headers();
  const auto footers = document.footers();
  const auto even_headers = document.headers(docxcpp::HeaderFooterType::EvenPage);
  const auto even_footers = document.footers(docxcpp::HeaderFooterType::EvenPage);
  const auto first_headers = document.headers(docxcpp::HeaderFooterType::FirstPage);
  const auto first_footers = document.footers(docxcpp::HeaderFooterType::FirstPage);
  const auto header_paragraphs_0 = document.header_paragraphs(0);
  const auto header_paragraphs_1 = document.header_paragraphs(1);
  const auto footer_paragraphs_0 = document.footer_paragraphs(0);
  const auto footer_paragraphs_1 = document.footer_paragraphs(1);
  const auto pictures = document.pictures();
  const auto picture_sizes = document.picture_sizes_pt();

  assert(paragraphs.size() == 10);
  assert(paragraphs[0].text() == "PythonDocx Title");
  assert(paragraphs[0].alignment() == docxcpp::ParagraphAlignment::Center);
  assert(paragraphs[0].style_id() == "Title");
  assert(paragraphs[0].heading_level() == 0);

  assert(paragraphs[1].text() == "Styled body\nline 2");
  assert(paragraphs[1].alignment() == docxcpp::ParagraphAlignment::Justify);
  assert(paragraphs[1].format().left_indent_pt.has_value());
  assert(*paragraphs[1].format().left_indent_pt == 24);
  assert(paragraphs[1].format().right_indent_pt.has_value());
  assert(*paragraphs[1].format().right_indent_pt == 18);
  assert(paragraphs[1].format().first_line_indent_pt.has_value());
  assert(*paragraphs[1].format().first_line_indent_pt == -12);
  assert(paragraphs[1].format().space_before_pt.has_value());
  assert(*paragraphs[1].format().space_before_pt == 6);
  assert(paragraphs[1].format().space_after_pt.has_value());
  assert(*paragraphs[1].format().space_after_pt == 10);
  assert(paragraphs[1].format().line_spacing_pt.has_value());
  assert(*paragraphs[1].format().line_spacing_pt == 18);
  assert(paragraphs[1].format().line_spacing_mode.has_value());
  assert(*paragraphs[1].format().line_spacing_mode == docxcpp::LineSpacingMode::Exact);
  assert(paragraphs[1].format().tab_stops.size() == 2);
  assert(paragraphs[1].format().keep_together.has_value());
  assert(*paragraphs[1].format().keep_together);
  assert(paragraphs[1].format().keep_with_next.has_value());
  assert(*paragraphs[1].format().keep_with_next);
  assert(paragraphs[1].format().page_break_before.has_value());
  assert(*paragraphs[1].format().page_break_before);
  assert(paragraphs[1].format().widow_control.has_value());
  assert(!*paragraphs[1].format().widow_control);
  assert(paragraphs[1].runs().size() == 2);
  assert(paragraphs[1].runs()[0].text() == "Styled ");
  assert(paragraphs[1].runs()[0].style().character_style_id == "Strong");
  assert(paragraphs[1].runs()[0].style().character_style_name == "Strong");
  assert(paragraphs[1].runs()[0].style().bold);
  assert(paragraphs[1].runs()[0].style().font_size_pt == 16);
  assert(paragraphs[1].runs()[0].style().color_hex == "123456");
  assert(paragraphs[1].runs()[0].style().highlight == "yellow");
  assert(paragraphs[1].runs()[1].text() == "body\nline 2");
  assert(paragraphs[1].runs()[1].style().italic);
  assert(paragraphs[1].runs()[1].style().font_name == "Courier New");
  assert(paragraphs[1].runs()[1].style().font_size_pt == 14);
  assert(paragraphs[1].runs()[1].style().color_hex == "AA2211");
  assert(paragraphs[1].runs()[1].style().highlight == "green");

  assert(paragraphs[2].text() == "CapsSubSuper");
  assert(paragraphs[2].runs().size() == 3);
  assert(paragraphs[2].runs()[0].style().font_name == "Arial");
  assert(paragraphs[2].runs()[0].style().ascii_font_name == "Arial");
  assert(paragraphs[2].runs()[0].style().h_ansi_font_name == "Arial");
  assert(paragraphs[2].runs()[0].style().east_asia_font_name == "SimSun");
  assert(paragraphs[2].runs()[0].style().complex_script_font_name == "Amiri");
  assert(paragraphs[2].runs()[0].style().all_caps);
  assert(paragraphs[2].runs()[0].style().strike);
  assert(paragraphs[2].runs()[1].style().small_caps);
  assert(paragraphs[2].runs()[1].style().subscript);
  assert(paragraphs[2].runs()[2].style().double_strike);
  assert(paragraphs[2].runs()[2].style().superscript);

  assert(paragraphs[3].text() == "Alpha\tBeta\tGamma");
  assert(paragraphs[3].alignment() == docxcpp::ParagraphAlignment::Center);
  assert(paragraphs[3].format().tab_stops.size() == 2);

  assert(paragraphs[4].text() == "Python Link");
  assert(paragraphs[4].hyperlinks().size() == 1);
  assert(paragraphs[4].hyperlinks()[0].text == "Python Link");
  assert(paragraphs[4].hyperlinks()[0].url == "https://example.com/python-link");

  assert(paragraphs[5].text() == "Commented paragraph");
  assert(paragraphs[6].text() == "Body picture: ");
  assert(paragraphs[7].text() == "\n");
  assert(paragraphs[7].has_page_break());
  assert(paragraphs[8].text().empty());
  assert(paragraphs[8].runs().empty());
  assert(paragraphs[9].text() == "Second section paragraph");

  assert(comments.size() == 1);
  assert(comments[0].text == "python fixture note");
  assert(comments[0].author == "PyDocx");
  assert(comments[0].initials == "PD");
  assert(document.comment_count() == 1);

  assert(tables.size() == 1);
  assert(tables[0].style_id() == "LightShading-Accent1");
  assert(tables[0].style_name() == "Light Shading Accent 1");
  assert(tables[0].row_count() == 3);
  assert(tables[0].column_count() == 3);
  assert(tables[0].cell(0, 0).text() == "Merged Head");
  assert(tables[0].cell(0, 1).text() == "Merged Head");
  assert(tables[0].cell(1, 1).text() == "Outer Cell\nCell Link\n");
  assert(tables[0].cell(1, 1).nested_tables().empty());
  assert(tables[0].rows()[1].cells()[2].vertical_merge() == "restart");
  assert(tables[0].rows()[2].cells()[2].vertical_merge() == "continue");
  assert(tables[0].cell(1, 2).text() == "Vertical Merge");
  assert(tables[0].cell(2, 2).text() == "Vertical Merge");
  assert(tables[0].cell(2, 1).text() == "Nested Host\n");
  assert(tables[0].cell(2, 1).nested_tables().size() == 1);
  assert(tables[0].cell(2, 1).nested_tables()[0].style_id() == "TableGrid");
  assert(tables[0].cell(2, 1).nested_tables()[0].style_name() == "Table Grid");
  assert(tables[0].cell(2, 1).nested_tables()[0].row_count() == 1);
  assert(tables[0].cell(2, 1).nested_tables()[0].column_count() == 2);
  assert(tables[0].cell(2, 1).nested_tables()[0].cell(0, 0).text() == "Nested A");
  assert(tables[0].cell(2, 1).nested_tables()[0].cell(0, 1).text() == "Nested B");

  assert(sections.size() == 2);
  assert(sections[0].page_orientation() == docxcpp::PageOrientation::Landscape);
  assert(std::llround(sections[0].page_size().width.pt()) == 720);
  assert(std::llround(sections[0].page_size().height.pt()) == 540);
  assert(std::llround(sections[0].page_margins().top.pt()) == 72);
  assert(std::llround(sections[0].page_margins().right.pt()) == 54);
  assert(std::llround(sections[0].page_margins().bottom.pt()) == 72);
  assert(std::llround(sections[0].page_margins().left.pt()) == 54);
  assert(std::llround(sections[0].page_margins().header.pt()) == 30);
  assert(std::llround(sections[0].page_margins().footer.pt()) == 24);
  assert(std::llround(sections[0].page_margins().gutter.pt()) == 12);
  assert(sections[1].page_orientation() == docxcpp::PageOrientation::Portrait);
  assert(std::llround(sections[1].page_size().width.pt()) == 612);
  assert(std::llround(sections[1].page_size().height.pt()) == 792);
  assert(std::llround(sections[1].page_margins().top.pt()) == 90);
  assert(std::llround(sections[1].page_margins().right.pt()) == 36);
  assert(std::llround(sections[1].page_margins().bottom.pt()) == 54);
  assert(std::llround(sections[1].page_margins().left.pt()) == 36);
  assert(std::llround(sections[1].page_margins().header.pt()) == 18);
  assert(std::llround(sections[1].page_margins().footer.pt()) == 18);

  assert(document.even_and_odd_headers());
  assert(document.section_different_first_page(0));
  assert(!document.section_different_first_page(1));
  assert(headers.size() == 2);
  assert(headers[0] == "Py Default Header\nHeader Tail");
  assert(headers[1] == "Second Header");
  assert(footers.size() == 2);
  assert(footers[0] == "Py Default Footer");
  assert(footers[1] == "Second Footer");
  assert(first_headers.size() == 2);
  assert(first_headers[0] == "Py First Header");
  assert(first_headers[1].empty());
  assert(first_footers.size() == 2);
  assert(first_footers[0] == "Py First Footer");
  assert(first_footers[1].empty());
  assert(even_headers.size() == 2);
  assert(even_headers[0] == "Py Even Header");
  assert(even_headers[1].empty());
  assert(even_footers.size() == 2);
  assert(even_footers[0] == "Py Even Footer");
  assert(even_footers[1].empty());
  assert(header_paragraphs_0.size() == 2);
  assert(header_paragraphs_0[0].text() == "Py Default Header");
  assert(header_paragraphs_0[1].text() == "Header Tail");
  assert(header_paragraphs_1.size() == 1);
  assert(header_paragraphs_1[0].text() == "Second Header");
  assert(footer_paragraphs_0.size() == 1);
  assert(footer_paragraphs_0[0].text() == "Py Default Footer");
  assert(footer_paragraphs_1.size() == 1);
  assert(footer_paragraphs_1[0].text() == "Second Footer");

  assert(pictures.size() == 2);
  assert(picture_sizes.size() == 2);
  assert(!pictures[0].in_table_cell);
  assert(pictures[0].size_pt.width_pt == 64);
  assert(pictures[0].size_pt.height_pt == 32);
  assert(pictures[1].in_table_cell);
  assert(pictures[1].table_index == 0);
  assert(pictures[1].row_index == 1);
  assert(pictures[1].col_index == 1);
  assert(pictures[1].size_pt.width_pt == 24);
  assert(pictures[1].size_pt.height_pt == 24);

  return 0;
}
