#include <cassert>
#include <filesystem>
#include <fstream>
#include <string>
#include <vector>

#include "docxcpp/document.hpp"

namespace {

std::vector<std::uint8_t> read_bytes(const std::filesystem::path& path) {
  std::ifstream stream(path, std::ios::binary);
  stream.seekg(0, std::ios::end);
  const auto size = static_cast<std::size_t>(stream.tellg());
  stream.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(size);
  if (size > 0) {
    stream.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(size));
  }
  return bytes;
}

}  // namespace

int main() {
  namespace fs = std::filesystem;

  const fs::path source_dir = DOCXCPP_SOURCE_DIR;
  const fs::path artifact_dir = DOCXCPP_TEST_ARTIFACT_DIR;
  const fs::path image_path = source_dir / "cpp" / "image.jpeg";
  const auto image_bytes = read_bytes(image_path);
  fs::create_directories(artifact_dir);

  docxcpp::Document created;
  docxcpp::RunStyle title_style;
  title_style.bold = true;
  title_style.font_size_pt = 20;
  title_style.color_hex = "#336699";
  title_style.font_name = "Arial";
  title_style.highlight = "yellow";
  created.add_styled_heading("Document Title", 0, title_style,
                             docxcpp::ParagraphAlignment::Center);
  created.set_page_size_pt(595, 842);
  created.set_page_orientation(docxcpp::PageOrientation::Landscape);
  docxcpp::PageMargins margins;
  margins.top_pt = 72;
  margins.right_pt = 54;
  margins.bottom_pt = 72;
  margins.left_pt = 54;
  margins.header_pt = 36;
  margins.footer_pt = 24;
  margins.gutter_pt = 18;
  created.set_page_margins_pt(margins);
  created.add_paragraph("alpha");
  created.add_comment(1, "alpha note", "Codex", "CX");
  created.set_paragraph_indentation_pt(1, 24, 12, -18);
  created.set_paragraph_spacing_pt(1, 6, 10);
  created.set_paragraph_line_spacing_pt(1, 18);
  created.set_paragraph_pagination(1, true, true, true);
  created.add_hyperlink("OpenAI", "https://openai.com", docxcpp::ParagraphAlignment::Center);
  docxcpp::RunStyle mixed_bold;
  mixed_bold.bold = true;
  mixed_bold.color_hex = "008800";
  docxcpp::RunStyle mixed_plain;
  mixed_plain.font_name = "Times New Roman";
  std::vector<docxcpp::Run> mixed_runs{
      docxcpp::Run("mix-", mixed_bold),
      docxcpp::Run("ed", mixed_plain),
  };
  created.add_paragraph(mixed_runs, docxcpp::ParagraphAlignment::Left);
  created.add_page_break();
  docxcpp::RunStyle aside_style;
  aside_style.italic = true;
  aside_style.underline = true;
  aside_style.font_size_pt = 14;
  aside_style.color_hex = "AA2200";
  aside_style.font_name = "Courier New";
  aside_style.highlight = "green";
  created.add_styled_paragraph(" beta \nline 2", aside_style, docxcpp::ParagraphAlignment::Right);
  created.add_heading("Section", 1);
  created.add_table(2, 3);
  created.set_table_cell(0, 0, 0, "r1c1");
  std::vector<docxcpp::Run> cell_runs{
      docxcpp::Run("r2", mixed_bold),
      docxcpp::Run("c3", mixed_plain),
  };
  created.set_table_cell(0, 1, 2, cell_runs);
  created.add_table_cell_paragraph(0, 1, 2, "second paragraph");
  created.merge_table_cells(0, 0, 0, 0, 1);
  docxcpp::PictureSize image_size;
  image_size.width_pt = 72;
  image_size.height_pt = 36;
  created.add_picture_data(image_bytes, "jpeg", image_size, "inline");
  docxcpp::PictureSize cell_image_size;
  cell_image_size.width_pt = 36;
  cell_image_size.height_pt = 36;
  created.add_picture_to_table_cell(0, 1, 2, image_path, cell_image_size);

  const fs::path output = artifact_dir / "all-features-generated.docx";
  created.save(output);

  docxcpp::Document reopened(output);
  const auto paragraphs = reopened.paragraphs();
  const auto tables = reopened.tables();
  const auto page_size = reopened.page_size_pt();
  const auto page_orientation = reopened.page_orientation();
  const auto page_margins = reopened.page_margins_pt();
  const auto comments = reopened.comments();
  const auto picture_sizes = reopened.picture_sizes_pt();
  const auto pictures = reopened.pictures();

  assert(paragraphs.size() == 8);
  assert(paragraphs[0].text() == "Document Title");
  assert(paragraphs[0].alignment() == docxcpp::ParagraphAlignment::Center);
  assert(paragraphs[0].first_run_style().bold);
  assert(!paragraphs[0].first_run_style().italic);
  assert(!paragraphs[0].first_run_style().underline);
  assert(paragraphs[0].first_run_style().font_size_pt == 20);
  assert(paragraphs[0].first_run_style().font_name == "Arial");
  assert(paragraphs[0].first_run_style().color_hex == "336699");
  assert(paragraphs[0].first_run_style().highlight == "yellow");
  assert(!paragraphs[0].has_page_break());
  assert(paragraphs[0].style_id() == "Title");
  assert(paragraphs[0].heading_level() == 0);
  assert(paragraphs[0].runs().size() == 1);
  assert(paragraphs[0].runs()[0].text() == "Document Title");
  assert(paragraphs[0].runs()[0].style().bold);
  assert(paragraphs[0].runs()[0].style().font_size_pt == 20);
  assert(paragraphs[1].text() == "alpha");
  assert(paragraphs[1].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(!paragraphs[1].has_page_break());
  assert(paragraphs[1].style_id().empty());
  assert(paragraphs[1].heading_level() == -1);
  assert(paragraphs[1].format().left_indent_pt.has_value());
  assert(*paragraphs[1].format().left_indent_pt == 24);
  assert(paragraphs[1].format().right_indent_pt.has_value());
  assert(*paragraphs[1].format().right_indent_pt == 12);
  assert(paragraphs[1].format().first_line_indent_pt.has_value());
  assert(*paragraphs[1].format().first_line_indent_pt == -18);
  assert(paragraphs[1].format().space_before_pt.has_value());
  assert(*paragraphs[1].format().space_before_pt == 6);
  assert(paragraphs[1].format().space_after_pt.has_value());
  assert(*paragraphs[1].format().space_after_pt == 10);
  assert(paragraphs[1].format().line_spacing_pt.has_value());
  assert(*paragraphs[1].format().line_spacing_pt == 18);
  assert(paragraphs[1].format().keep_together.has_value());
  assert(*paragraphs[1].format().keep_together);
  assert(paragraphs[1].format().keep_with_next.has_value());
  assert(*paragraphs[1].format().keep_with_next);
  assert(paragraphs[1].format().page_break_before.has_value());
  assert(*paragraphs[1].format().page_break_before);
  assert(paragraphs[1].runs().size() == 2);
  assert(paragraphs[1].runs()[0].text() == "alpha");
  assert(!paragraphs[1].runs()[0].style().bold);
  assert(paragraphs[1].runs()[1].text().empty());
  assert(paragraphs[2].text() == "OpenAI");
  assert(paragraphs[2].alignment() == docxcpp::ParagraphAlignment::Center);
  assert(paragraphs[2].runs().size() == 1);
  assert(paragraphs[2].runs()[0].text() == "OpenAI");
  assert(paragraphs[3].text() == "mix-ed");
  assert(paragraphs[3].alignment() == docxcpp::ParagraphAlignment::Left);
  assert(!paragraphs[3].has_page_break());
  assert(paragraphs[3].runs().size() == 2);
  assert(paragraphs[3].runs()[0].text() == "mix-");
  assert(paragraphs[3].runs()[0].style().bold);
  assert(paragraphs[3].runs()[0].style().color_hex == "008800");
  assert(paragraphs[3].runs()[1].text() == "ed");
  assert(paragraphs[3].runs()[1].style().font_name == "Times New Roman");
  assert(paragraphs[4].text() == "\n");
  assert(paragraphs[4].has_page_break());
  assert(paragraphs[4].runs().size() == 1);
  assert(paragraphs[4].runs()[0].text() == "\n");
  assert(paragraphs[4].runs()[0].style().font_size_pt == 0);
  assert(paragraphs[5].text() == " beta \nline 2");
  assert(paragraphs[5].alignment() == docxcpp::ParagraphAlignment::Right);
  assert(!paragraphs[5].first_run_style().bold);
  assert(paragraphs[5].first_run_style().italic);
  assert(paragraphs[5].first_run_style().underline);
  assert(paragraphs[5].first_run_style().font_size_pt == 14);
  assert(paragraphs[5].first_run_style().font_name == "Courier New");
  assert(paragraphs[5].first_run_style().color_hex == "AA2200");
  assert(paragraphs[5].first_run_style().highlight == "green");
  assert(!paragraphs[5].has_page_break());
  assert(paragraphs[5].style_id().empty());
  assert(paragraphs[5].heading_level() == -1);
  assert(paragraphs[5].runs().size() == 1);
  assert(paragraphs[5].runs()[0].text() == " beta \nline 2");
  assert(paragraphs[5].runs()[0].style().italic);
  assert(paragraphs[5].runs()[0].style().underline);
  assert(paragraphs[5].runs()[0].style().font_name == "Courier New");
  assert(paragraphs[6].text() == "Section");
  assert(paragraphs[6].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(paragraphs[6].style_id() == "Heading1");
  assert(paragraphs[6].heading_level() == 1);
  assert(paragraphs[6].runs().size() == 1);
  assert(paragraphs[6].runs()[0].text() == "Section");
  assert(paragraphs[7].text().empty());
  assert(paragraphs[7].style_id().empty());
  assert(paragraphs[7].heading_level() == -1);
  assert(paragraphs[7].runs().size() == 1);
  assert(paragraphs[7].runs()[0].text().empty());
  assert(comments.size() == 1);
  assert(comments[0].id == 0);
  assert(comments[0].text == "alpha note");
  assert(comments[0].author == "Codex");
  assert(comments[0].initials == "CX");
  assert(reopened.comment_count() == 1);

  assert(tables.size() == 1);
  assert(tables[0].rows().size() == 2);
  assert(tables[0].rows()[0].cells().size() == 2);
  assert(tables[0].rows()[0].cells()[0].text() == "r1c1");
  assert(tables[0].rows()[0].cells()[0].grid_span() == 2);
  assert(tables[0].rows()[0].cells()[0].vertical_merge().empty());
  assert(tables[0].rows()[1].cells()[2].text() == "r2c3\nsecond paragraph\n");
  assert(tables[0].rows()[1].cells()[2].grid_span() == 1);
  assert(tables[0].rows()[1].cells()[2].vertical_merge().empty());
  assert(page_size.width_pt == 595);
  assert(page_size.height_pt == 842);
  assert(page_orientation == docxcpp::PageOrientation::Landscape);
  assert(page_margins.top_pt == 72);
  assert(page_margins.right_pt == 54);
  assert(page_margins.bottom_pt == 72);
  assert(page_margins.left_pt == 54);
  assert(page_margins.header_pt == 36);
  assert(page_margins.footer_pt == 24);
  assert(page_margins.gutter_pt == 18);
  assert(picture_sizes.size() == 2);
  assert(picture_sizes[0].width_pt == 36);
  assert(picture_sizes[0].height_pt == 36);
  assert(picture_sizes[1].width_pt == 72);
  assert(picture_sizes[1].height_pt == 36);
  assert(pictures.size() == 2);
  assert(pictures[0].in_table_cell);
  assert(pictures[0].table_index == 0);
  assert(pictures[0].row_index == 1);
  assert(pictures[0].col_index == 2);
  assert(pictures[0].target == "media/image2.jpeg");
  assert(pictures[0].size_pt.width_pt == 36);
  assert(pictures[0].size_pt.height_pt == 36);
  assert(!pictures[1].in_table_cell);
  assert(pictures[1].target == "media/image1.jpeg");
  assert(pictures[1].size_pt.width_pt == 72);
  assert(pictures[1].size_pt.height_pt == 36);
  assert(reopened.image_count() == 2);

  const auto package = docxcpp::OpcPackage::open(output);
  assert(package.has_entry("word/document.xml"));
  assert(package.has_entry("word/comments.xml"));
  assert(package.has_entry("[Content_Types].xml"));
  assert(package.has_entry("word/_rels/document.xml.rels"));
  assert(package.has_entry("word/media/image1.jpeg"));
  assert(package.has_entry("word/media/image2.jpeg"));
  const auto& image1 = package.entry("word/media/image1.jpeg");
  const auto& image2 = package.entry("word/media/image2.jpeg");
  assert(!image1.empty());
  assert(image1 == image2);

  const auto& document_xml = package.entry("word/document.xml");
  const std::string xml(document_xml.begin(), document_xml.end());
  const std::string rels_xml(package.entry("word/_rels/document.xml.rels").begin(),
                             package.entry("word/_rels/document.xml.rels").end());
  const std::string comments_xml(package.entry("word/comments.xml").begin(),
                                 package.entry("word/comments.xml").end());
  assert(xml.find("w:jc w:val=\"center\"") != std::string::npos);
  assert(xml.find("w:jc w:val=\"right\"") != std::string::npos);
  assert(xml.find("w:jc w:val=\"left\"") != std::string::npos);
  assert(xml.find("<w:b") != std::string::npos);
  assert(xml.find("<w:i") != std::string::npos);
  assert(xml.find("<w:u w:val=\"single\"") != std::string::npos);
  assert(xml.find("<w:br") != std::string::npos);
  assert(xml.find("w:br w:type=\"page\"") != std::string::npos);
  assert(xml.find("w:color w:val=\"336699\"") != std::string::npos);
  assert(xml.find("w:color w:val=\"AA2200\"") != std::string::npos);
  assert(xml.find("w:color w:val=\"008800\"") != std::string::npos);
  assert(xml.find("w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"") != std::string::npos);
  assert(xml.find("w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\"") != std::string::npos);
  assert(xml.find("w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"") !=
         std::string::npos);
  assert(xml.find("w:ind w:left=\"480\" w:right=\"240\" w:hanging=\"360\"") !=
         std::string::npos);
  assert(xml.find("w:spacing w:before=\"120\" w:after=\"200\" w:line=\"360\" w:lineRule=\"exact\"") !=
         std::string::npos);
  assert(xml.find("<w:keepLines") != std::string::npos);
  assert(xml.find("<w:keepNext") != std::string::npos);
  assert(xml.find("<w:pageBreakBefore") != std::string::npos);
  assert(xml.find("<w:hyperlink r:id=\"") != std::string::npos);
  assert(xml.find("<w:commentRangeStart w:id=\"0\"") != std::string::npos);
  assert(xml.find("<w:commentRangeEnd w:id=\"0\"") != std::string::npos);
  assert(xml.find("<w:commentReference w:id=\"0\"") != std::string::npos);
  assert(xml.find("w:highlight w:val=\"yellow\"") != std::string::npos);
  assert(xml.find("w:highlight w:val=\"green\"") != std::string::npos);
  assert(xml.find("w:gridSpan w:val=\"2\"") != std::string::npos);
  assert(xml.find("w:pgSz w:w=\"11900\" w:h=\"16840\" w:orient=\"landscape\"") !=
         std::string::npos);
  assert(xml.find(
             "w:pgMar w:top=\"1440\" w:right=\"1080\" w:bottom=\"1440\" w:left=\"1080\" w:header=\"720\" w:footer=\"480\" w:gutter=\"360\"") !=
         std::string::npos);
  assert(xml.find("wp:extent cx=\"914400\" cy=\"457200\"") != std::string::npos);
  assert(xml.find("a:ext cx=\"914400\" cy=\"457200\"") != std::string::npos);
  assert(xml.find("wp:extent cx=\"457200\" cy=\"457200\"") != std::string::npos);
  assert(xml.find("a:ext cx=\"457200\" cy=\"457200\"") != std::string::npos);
  assert(xml.find("w:sz w:val=\"40\"") != std::string::npos);
  assert(xml.find("w:sz w:val=\"28\"") != std::string::npos);
  assert(rels_xml.find("TargetMode=\"External\"") != std::string::npos);
  assert(rels_xml.find("Target=\"https://openai.com\"") != std::string::npos);
  assert(rels_xml.find("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments") !=
         std::string::npos);
  assert(comments_xml.find("w:comment w:id=\"0\"") != std::string::npos);
  assert(comments_xml.find("w:author=\"Codex\"") != std::string::npos);
  assert(comments_xml.find("alpha note") != std::string::npos);
  return 0;
}
