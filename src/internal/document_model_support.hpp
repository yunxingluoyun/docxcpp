#pragma once

#include <cstddef>
#include <memory>
#include <string>

#include "pugixml.hpp"

#include "docxcpp/paragraph.hpp"

namespace docxcpp {

struct StyleCatalog;

struct ParagraphBinding {
  pugi::xml_document* xml{nullptr};
  bool* dirty{nullptr};
  std::shared_ptr<StyleCatalog> style_catalog;
  int kind{0};
  std::size_t paragraph_index{0};
  std::size_t table_index{0};
  std::size_t row_index{0};
  std::size_t col_index{0};
  std::size_t cell_paragraph_index{0};
};

pugi::xml_node paragraph_node_for_model_binding(const ParagraphBinding& binding);
void insert_paragraph_before_binding(ParagraphBinding& binding, const std::vector<Run>& runs,
                                     ParagraphAlignment alignment, Paragraph& paragraph);
void apply_run_style_for_model(pugi::xml_node run, const RunStyle& style,
                               const StyleCatalog* style_catalog = nullptr);
void append_run_text_for_model(pugi::xml_node run, const std::string& text);

}  // namespace docxcpp
