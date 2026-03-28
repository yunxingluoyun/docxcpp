#include <cassert>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

#include "docxcpp/document.hpp"
#include "docxcpp/opc_package.hpp"
#include "pugixml.hpp"

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

std::vector<std::uint8_t> xml_to_bytes(const pugi::xml_document& xml) {
  std::ostringstream stream;
  xml.save(stream, "  ", pugi::format_default, pugi::encoding_utf8);
  const std::string content = stream.str();
  return std::vector<std::uint8_t>(content.begin(), content.end());
}

}  // namespace

int main() {
  namespace fs = std::filesystem;

  const fs::path source_dir = DOCXCPP_SOURCE_DIR;
  const fs::path artifact_dir = DOCXCPP_TEST_ARTIFACT_DIR;
  const fs::path image_path = source_dir / "cpp" / "image.jpeg";
  const auto image_bytes = read_bytes(image_path);
  fs::create_directories(artifact_dir);

  docxcpp::Document document;
  docxcpp::RunStyle accent_style;
  accent_style.bold = true;
  accent_style.color_hex = "112233";
  docxcpp::RunStyle tail_style;
  tail_style.font_name = "Georgia";
  std::vector<docxcpp::Run> heading_runs{
      docxcpp::Run("Heading ", accent_style),
      docxcpp::Run("Wrapper", tail_style),
  };
  std::vector<docxcpp::Run> paragraph_runs{
      docxcpp::Run("Plain ", accent_style),
      docxcpp::Run("paragraph", tail_style),
  };
  document.add_heading("Heading Wrapper", 2);
  document.add_heading(heading_runs, 3, docxcpp::ParagraphAlignment::Center);
  document.add_paragraph("Plain paragraph");
  document.add_paragraph(paragraph_runs);
  auto mutable_paragraph = document.add_paragraph("Mutable");
  mutable_paragraph.add_run(" tail", accent_style);
  document.set_paragraph_pagination(2, false, false, std::nullopt);
  document.set_page_orientation(docxcpp::PageOrientation::Landscape);
  document.add_comment(2, "coverage note", "DocxCPP", "DC");
  docxcpp::RunStyle link_style;
  link_style.underline = true;
  link_style.color_hex = "0000EE";
  auto project_link = document.add_hyperlink("Project", "https://example.com/project",
                                             docxcpp::ParagraphAlignment::Right, link_style);
  assert(project_link.hyperlinks().size() == 1);
  assert(project_link.hyperlinks()[0].text == "Project");
  assert(project_link.hyperlinks()[0].url == "https://example.com/project");
  document.add_table(2, 2);
  document.set_table_cell(0, 0, 0, "A1");
  document.set_table_cell(0, 0, 1, paragraph_runs);
  document.add_table_cell_paragraph(0, 0, 0, "A1 second");
  document.add_table_cell_paragraph(0, 1, 0, heading_runs);
  document.add_hyperlink_to_table_cell(0, 1, 0, "CellLink", "https://example.com/cell");
  document.add_picture_data(image_bytes, "jpeg");
  document.add_picture_data_to_table_cell(0, 1, 1, image_bytes, "jpeg", "buffered");

  const fs::path output = artifact_dir / "api-coverage-generated.docx";
  document.save(output);

  docxcpp::Document reopened(output);
  const auto paragraphs = reopened.paragraphs();
  const auto tables = reopened.tables();
  const auto sections = reopened.sections();
  const auto page_size = reopened.page_size_pt();
  const auto page_orientation = reopened.page_orientation();
  const auto page_margins = reopened.page_margins_pt();
  const auto comments = reopened.comments();
  const auto picture_sizes = reopened.picture_sizes_pt();
  const auto pictures = reopened.pictures();

  assert(paragraphs.size() == 7);
  assert(paragraphs[0].text() == "Heading Wrapper");
  assert(paragraphs[0].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(!paragraphs[0].has_page_break());
  assert(paragraphs[0].style_id() == "Heading2");
  assert(paragraphs[0].heading_level() == 2);
  assert(paragraphs[0].runs().size() == 1);
  assert(paragraphs[0].runs()[0].text() == "Heading Wrapper");
  assert(paragraphs[1].text() == "Heading Wrapper");
  assert(paragraphs[1].alignment() == docxcpp::ParagraphAlignment::Center);
  assert(paragraphs[1].style_id() == "Heading3");
  assert(paragraphs[1].heading_level() == 3);
  assert(paragraphs[1].runs().size() == 2);
  assert(paragraphs[1].runs()[0].text() == "Heading ");
  assert(paragraphs[1].runs()[0].style().bold);
  assert(paragraphs[1].runs()[1].text() == "Wrapper");
  assert(paragraphs[1].runs()[1].style().font_name == "Georgia");
  assert(paragraphs[2].text() == "Plain paragraph");
  assert(paragraphs[2].alignment() == docxcpp::ParagraphAlignment::Inherit);
  assert(paragraphs[2].style_id().empty());
  assert(paragraphs[2].heading_level() == -1);
  assert(!paragraphs[2].format().left_indent_pt.has_value());
  assert(!paragraphs[2].format().space_before_pt.has_value());
  assert(!paragraphs[2].format().line_spacing_pt.has_value());
  assert(paragraphs[2].format().keep_together.has_value());
  assert(!*paragraphs[2].format().keep_together);
  assert(paragraphs[2].format().keep_with_next.has_value());
  assert(!*paragraphs[2].format().keep_with_next);
  assert(!paragraphs[2].format().page_break_before.has_value());
  assert(paragraphs[2].runs().size() == 2);
  assert(paragraphs[2].runs()[0].text() == "Plain paragraph");
  assert(paragraphs[2].runs()[1].text().empty());
  assert(paragraphs[3].text() == "Plain paragraph");
  assert(paragraphs[3].runs().size() == 2);
  assert(paragraphs[3].runs()[0].style().color_hex == "112233");
  assert(paragraphs[3].runs()[1].style().font_name == "Georgia");
  assert(paragraphs[4].text() == "Mutable tail");
  assert(paragraphs[4].runs().size() == 2);
  assert(paragraphs[4].runs()[0].text() == "Mutable");
  assert(paragraphs[4].runs()[1].text() == " tail");
  assert(paragraphs[4].runs()[1].style().bold);
  assert(paragraphs[4].runs()[1].style().color_hex == "112233");
  assert(paragraphs[4].hyperlinks().empty());
  assert(paragraphs[5].text() == "Project");
  assert(paragraphs[5].alignment() == docxcpp::ParagraphAlignment::Right);
  assert(paragraphs[5].style_id().empty());
  assert(paragraphs[5].heading_level() == -1);
  assert(paragraphs[5].runs().size() == 1);
  assert(paragraphs[5].runs()[0].text() == "Project");
  assert(paragraphs[5].runs()[0].style().underline);
  assert(paragraphs[5].runs()[0].style().color_hex == "0000EE");
  assert(paragraphs[5].hyperlinks().size() == 1);
  assert(paragraphs[5].hyperlinks()[0].text == "Project");
  assert(paragraphs[5].hyperlinks()[0].url == "https://example.com/project");
  assert(paragraphs[6].text().empty());
  assert(!paragraphs[6].has_page_break());
  assert(paragraphs[6].style_id().empty());
  assert(paragraphs[6].heading_level() == -1);
  assert(paragraphs[6].runs().size() == 1);
  assert(paragraphs[6].runs()[0].text().empty());

  assert(comments.size() == 1);
  assert(comments[0].id == 0);
  assert(comments[0].text == "coverage note");
  assert(comments[0].author == "DocxCPP");
  assert(comments[0].initials == "DC");
  assert(reopened.comment_count() == 1);

  assert(tables.size() == 1);
  assert(tables[0].rows().size() == 2);
  assert(tables[0].rows()[0].cells().size() == 2);
  assert(tables[0].rows()[0].cells()[0].text() == "A1\nA1 second");
  assert(tables[0].rows()[0].cells()[1].text() == "Plain paragraph");
  assert(tables[0].rows()[0].cells()[0].grid_span() == 1);
  assert(tables[0].rows()[0].cells()[0].vertical_merge().empty());
  assert(tables[0].rows()[1].cells()[0].text() == "Heading Wrapper\nCellLink");
  assert(tables[0].rows()[1].cells()[1].text().empty());
  assert(sections.size() == 1);
  assert(sections[0].page_size_pt().width_pt == 612);
  assert(sections[0].page_size_pt().height_pt == 792);
  assert(sections[0].page_orientation() == docxcpp::PageOrientation::Landscape);
  assert(sections[0].page_margins_pt().top_pt == 72);
  assert(sections[0].page_margins_pt().right_pt == 90);
  assert(sections[0].page_margins_pt().bottom_pt == 72);
  assert(sections[0].page_margins_pt().left_pt == 90);
  assert(sections[0].page_margins_pt().header_pt == 36);
  assert(sections[0].page_margins_pt().footer_pt == 36);
  assert(sections[0].page_margins_pt().gutter_pt == 0);
  assert(page_size.width_pt == 612);
  assert(page_size.height_pt == 792);
  assert(page_orientation == docxcpp::PageOrientation::Landscape);
  assert(page_margins.top_pt == 72);
  assert(page_margins.right_pt == 90);
  assert(page_margins.bottom_pt == 72);
  assert(page_margins.left_pt == 90);
  assert(page_margins.header_pt == 36);
  assert(page_margins.footer_pt == 36);
  assert(page_margins.gutter_pt == 0);
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
  assert(package.has_entry("word/comments.xml"));

  const std::string xml(package.entry("word/document.xml").begin(),
                        package.entry("word/document.xml").end());
  const std::string rels_xml(package.entry("word/_rels/document.xml.rels").begin(),
                             package.entry("word/_rels/document.xml.rels").end());
  const std::string comments_xml(package.entry("word/comments.xml").begin(),
                                 package.entry("word/comments.xml").end());
  assert(xml.find("w:pStyle w:val=\"Heading2\"") != std::string::npos);
  assert(xml.find("w:pStyle w:val=\"Heading3\"") != std::string::npos);
  assert(xml.find("<w:keepLines w:val=\"0\"") != std::string::npos);
  assert(xml.find("<w:keepNext w:val=\"0\"") != std::string::npos);
  assert(xml.find("w:pgSz w:w=\"12240\" w:h=\"15840\" w:orient=\"landscape\"") !=
         std::string::npos);
  assert(xml.find("<w:hyperlink r:id=\"") != std::string::npos);
  assert(xml.find("<w:commentRangeStart w:id=\"0\"") != std::string::npos);
  assert(xml.find("<w:commentReference w:id=\"0\"") != std::string::npos);
  assert(xml.find("wp:extent cx=\"6353175\" cy=\"4762500\"") != std::string::npos);
  assert(xml.find("a:ext cx=\"6353175\" cy=\"4762500\"") != std::string::npos);
  assert(xml.find("w:color w:val=\"112233\"") != std::string::npos);
  assert(xml.find("w:color w:val=\"0000EE\"") != std::string::npos);
  assert(xml.find("w:rFonts w:ascii=\"Georgia\" w:hAnsi=\"Georgia\"") != std::string::npos);
  assert(rels_xml.find("TargetMode=\"External\"") != std::string::npos);
  assert(rels_xml.find("Target=\"https://example.com/project\"") != std::string::npos);
  assert(rels_xml.find("Target=\"https://example.com/cell\"") != std::string::npos);
  assert(rels_xml.find("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments") !=
         std::string::npos);
  assert(comments_xml.find("coverage note") != std::string::npos);

  {
    auto package = docxcpp::OpcPackage::open(output);
    const auto& document_xml = package.entry("word/document.xml");
    pugi::xml_document xml;
    const pugi::xml_parse_result parse_result =
        xml.load_buffer(document_xml.data(), document_xml.size());
    assert(parse_result);

    const auto child_named = [](const pugi::xml_node& parent, const char* name) {
      for (const pugi::xml_node& child : parent.children()) {
        if (std::string(child.name()) == name) {
          return child;
        }
      }
      return pugi::xml_node();
    };

    pugi::xml_node body = child_named(child_named(xml, "w:document"), "w:body");
    assert(body);
    const pugi::xml_node trailing_sect_pr = child_named(body, "w:sectPr");
    assert(trailing_sect_pr);

    pugi::xml_node first_paragraph;
    for (const pugi::xml_node& child : body.children()) {
      if (std::string(child.name()) == "w:p") {
        first_paragraph = child;
        break;
      }
    }
    assert(first_paragraph);

    pugi::xml_node injected_paragraph = body.insert_copy_before(first_paragraph, trailing_sect_pr);
    pugi::xml_node injected_p_pr = injected_paragraph.prepend_child("w:pPr");
    pugi::xml_node injected_sect_pr = injected_p_pr.append_child("w:sectPr");

    pugi::xml_node injected_pg_sz = injected_sect_pr.append_child("w:pgSz");
    injected_pg_sz.append_attribute("w:w").set_value("20000");
    injected_pg_sz.append_attribute("w:h").set_value("10000");
    injected_pg_sz.append_attribute("w:orient").set_value("landscape");

    pugi::xml_node injected_pg_mar = injected_sect_pr.append_child("w:pgMar");
    injected_pg_mar.append_attribute("w:top").set_value("1440");
    injected_pg_mar.append_attribute("w:right").set_value("720");
    injected_pg_mar.append_attribute("w:bottom").set_value("2160");
    injected_pg_mar.append_attribute("w:left").set_value("1080");
    injected_pg_mar.append_attribute("w:header").set_value("540");
    injected_pg_mar.append_attribute("w:footer").set_value("360");
    injected_pg_mar.append_attribute("w:gutter").set_value("180");

    package.set_entry("word/document.xml", xml_to_bytes(xml));
    const fs::path multi_section_output = artifact_dir / "api-coverage-multi-section.docx";
    package.save(multi_section_output);

    docxcpp::Document multi_section_doc(multi_section_output);
    const auto multi_sections = multi_section_doc.sections();
    assert(multi_sections.size() == 2);

    assert(multi_sections[0].page_size_pt().width_pt == 1000);
    assert(multi_sections[0].page_size_pt().height_pt == 500);
    assert(multi_sections[0].page_orientation() == docxcpp::PageOrientation::Landscape);
    assert(multi_sections[0].page_margins_pt().top_pt == 72);
    assert(multi_sections[0].page_margins_pt().right_pt == 36);
    assert(multi_sections[0].page_margins_pt().bottom_pt == 108);
    assert(multi_sections[0].page_margins_pt().left_pt == 54);
    assert(multi_sections[0].page_margins_pt().header_pt == 27);
    assert(multi_sections[0].page_margins_pt().footer_pt == 18);
    assert(multi_sections[0].page_margins_pt().gutter_pt == 9);

    assert(multi_sections[1].page_size_pt().width_pt == 612);
    assert(multi_sections[1].page_size_pt().height_pt == 792);
    assert(multi_sections[1].page_orientation() == docxcpp::PageOrientation::Landscape);
    assert(multi_sections[1].page_margins_pt().top_pt == 72);
    assert(multi_sections[1].page_margins_pt().right_pt == 90);
    assert(multi_sections[1].page_margins_pt().bottom_pt == 72);
    assert(multi_sections[1].page_margins_pt().left_pt == 90);
    assert(multi_sections[1].page_margins_pt().header_pt == 36);
    assert(multi_sections[1].page_margins_pt().footer_pt == 36);
    assert(multi_sections[1].page_margins_pt().gutter_pt == 0);
  }

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
    docxcpp::PageMargins margins;
    margins.header_pt = -1;
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
    document.add_comment(99, "bad", "DocxCPP", "DC");
  } catch (const std::out_of_range&) {
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

  threw = false;
  try {
    document.set_paragraph_indentation_pt(0, -1, 0, 0);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.set_paragraph_spacing_pt(0, -1, 0);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.set_paragraph_line_spacing_pt(0, 0);
  } catch (const std::invalid_argument&) {
    threw = true;
  }
  assert(threw);

  threw = false;
  try {
    document.set_paragraph_pagination(99, true, true, true);
  } catch (const std::out_of_range&) {
    threw = true;
  }
  assert(threw);

  return 0;
}
