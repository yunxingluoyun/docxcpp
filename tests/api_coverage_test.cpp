#include <cassert>
#include <filesystem>
#include <stdexcept>
#include <string>

#include "docxcpp/document.hpp"

int main() {
  namespace fs = std::filesystem;

  const fs::path source_dir = DOCXCPP_SOURCE_DIR;
  const fs::path artifact_dir = DOCXCPP_TEST_ARTIFACT_DIR;
  const fs::path image_path = source_dir / "cpp" / "image.jpeg";
  fs::create_directories(artifact_dir);

  docxcpp::Document document;
  document.add_heading("Heading Wrapper", 2);
  document.add_paragraph("Plain paragraph");
  document.add_table(2, 2);
  document.set_table_cell(0, 0, 0, "A1");
  document.add_table_cell_paragraph(0, 0, 0, "A1 second");
  document.add_picture(image_path);
  document.add_picture_to_table_cell(0, 1, 1, image_path);

  const fs::path output = artifact_dir / "api-coverage-generated.docx";
  document.save(output);

  docxcpp::Document reopened(output);
  const auto paragraphs = reopened.paragraphs();
  const auto tables = reopened.tables();
  const auto page_size = reopened.page_size_pt();
  const auto page_margins = reopened.page_margins_pt();
  const auto picture_sizes = reopened.picture_sizes_pt();
  const auto pictures = reopened.pictures();

  assert(paragraphs.size() == 3);
  assert(paragraphs[0].text() == "Heading Wrapper");
  assert(paragraphs[0].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(!paragraphs[0].has_page_break());
  assert(paragraphs[0].style_id() == "Heading2");
  assert(paragraphs[0].heading_level() == 2);
  assert(paragraphs[0].runs().size() == 1);
  assert(paragraphs[0].runs()[0].text() == "Heading Wrapper");
  assert(paragraphs[1].text() == "Plain paragraph");
  assert(paragraphs[1].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(paragraphs[1].style_id().empty());
  assert(paragraphs[1].heading_level() == -1);
  assert(paragraphs[1].runs().size() == 1);
  assert(paragraphs[1].runs()[0].text() == "Plain paragraph");
  assert(paragraphs[2].text().empty());
  assert(!paragraphs[2].has_page_break());
  assert(paragraphs[2].style_id().empty());
  assert(paragraphs[2].heading_level() == -1);
  assert(paragraphs[2].runs().size() == 1);
  assert(paragraphs[2].runs()[0].text().empty());

  assert(tables.size() == 1);
  assert(tables[0].rows().size() == 2);
  assert(tables[0].rows()[0].cells().size() == 2);
  assert(tables[0].rows()[0].cells()[0].text() == "A1\nA1 second");
  assert(tables[0].rows()[0].cells()[0].grid_span() == 1);
  assert(tables[0].rows()[0].cells()[0].vertical_merge().empty());
  assert(tables[0].rows()[1].cells()[1].text().empty());
  assert(page_size.width_pt == 612);
  assert(page_size.height_pt == 792);
  assert(page_margins.top_pt == 72);
  assert(page_margins.right_pt == 90);
  assert(page_margins.bottom_pt == 72);
  assert(page_margins.left_pt == 90);
  assert(picture_sizes.size() == 2);
  assert(picture_sizes[0].width_pt == 500);
  assert(picture_sizes[0].height_pt == 375);
  assert(picture_sizes[1].width_pt == 500);
  assert(picture_sizes[1].height_pt == 375);
  assert(pictures.size() == 2);
  assert(pictures[0].in_table_cell);
  assert(pictures[0].table_index == 0);
  assert(pictures[0].row_index == 1);
  assert(pictures[0].col_index == 1);
  assert(!pictures[1].in_table_cell);
  assert(reopened.image_count() == 2);

  const auto package = docxcpp::OpcPackage::open(output);
  assert(package.has_entry("word/media/image1.jpeg"));
  assert(package.has_entry("word/media/image2.jpeg"));

  const std::string xml(package.entry("word/document.xml").begin(),
                        package.entry("word/document.xml").end());
  assert(xml.find("w:pStyle w:val=\"Heading2\"") != std::string::npos);
  assert(xml.find("wp:extent cx=\"6353175\" cy=\"4762500\"") != std::string::npos);
  assert(xml.find("a:ext cx=\"6353175\" cy=\"4762500\"") != std::string::npos);

  bool threw = false;

  try {
    document.add_table(0, 1);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.set_page_size_pt(0, 100);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    docxcpp::PageMargins margins;
    margins.left_pt = -1;
    document.set_page_margins_pt(margins);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.merge_table_cells(0, 1, 1, 0, 0);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.set_paragraph_alignment(99, docxcpp::ParagraphAlignment::Center);
  } catch (const std::out_of_range&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    docxcpp::RunStyle bad_style;
    bad_style.color_hex = "xyz";
    document.add_styled_paragraph("bad", bad_style);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    docxcpp::RunStyle bad_style;
    bad_style.highlight = "bad value";
    document.add_styled_paragraph("bad", bad_style);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  return 0;
}
