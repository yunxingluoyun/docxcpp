#pragma once

#include <string>
#include <vector>

#include "pugixml.hpp"

#include "docxcpp/document.hpp"

namespace docxcpp {

void set_even_and_odd_headers_in_settings(OpcPackage& package, bool enabled);
bool read_even_and_odd_headers_from_settings(const OpcPackage& package);
void set_section_different_first_page_in_xml(pugi::xml_document& document_xml,
                                             std::size_t section_index, bool enabled);
bool read_section_different_first_page_from_xml(const pugi::xml_document& document_xml,
                                                std::size_t section_index);
void set_section_header_text(OpcPackage& package, pugi::xml_document& document_xml,
                             std::size_t section_index, const std::string& text,
                             HeaderFooterType type);
void set_section_footer_text(OpcPackage& package, pugi::xml_document& document_xml,
                             std::size_t section_index, const std::string& text,
                             HeaderFooterType type);
void clear_section_header(OpcPackage& package, pugi::xml_document& document_xml,
                          std::size_t section_index, HeaderFooterType type);
void clear_section_footer(OpcPackage& package, pugi::xml_document& document_xml,
                          std::size_t section_index, HeaderFooterType type);
Paragraph add_section_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::string& text,
                                       ParagraphAlignment alignment, HeaderFooterType type);
Paragraph add_section_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::vector<Run>& runs,
                                       ParagraphAlignment alignment, HeaderFooterType type);
Paragraph add_section_styled_header_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                              std::size_t section_index, const std::string& text,
                                              const RunStyle& style, ParagraphAlignment alignment,
                                              HeaderFooterType type);
Paragraph add_section_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::string& text,
                                       ParagraphAlignment alignment, HeaderFooterType type);
Paragraph add_section_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                       std::size_t section_index, const std::vector<Run>& runs,
                                       ParagraphAlignment alignment, HeaderFooterType type);
Paragraph add_section_styled_footer_paragraph(OpcPackage& package, pugi::xml_document& document_xml,
                                              std::size_t section_index, const std::string& text,
                                              const RunStyle& style, ParagraphAlignment alignment,
                                              HeaderFooterType type);
std::vector<std::string> read_section_headers(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              HeaderFooterType type);
std::vector<std::string> read_section_footers(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              HeaderFooterType type);
std::vector<Paragraph> read_header_paragraphs(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              std::size_t section_index, HeaderFooterType type);
std::vector<Paragraph> read_footer_paragraphs(const OpcPackage& package,
                                              const pugi::xml_document& document_xml,
                                              std::size_t section_index, HeaderFooterType type);

}  // namespace docxcpp
