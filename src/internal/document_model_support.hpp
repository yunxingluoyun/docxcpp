#pragma once

#include <cstddef>
#include <string>

#include "pugixml.hpp"

#include "docxcpp/paragraph.hpp"

namespace docxcpp {

struct ParagraphBinding {
  pugi::xml_document* xml{nullptr};
  bool* dirty{nullptr};
  int kind{0};
  std::size_t paragraph_index{0};
  std::size_t table_index{0};
  std::size_t row_index{0};
  std::size_t col_index{0};
  std::size_t cell_paragraph_index{0};
};

pugi::xml_node paragraph_node_for_model_binding(const ParagraphBinding& binding);
void apply_run_style_for_model(pugi::xml_node run, const RunStyle& style);
void append_run_text_for_model(pugi::xml_node run, const std::string& text);

}  // namespace docxcpp
