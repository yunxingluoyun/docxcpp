#include "docxcpp/paragraph.hpp"

#include <stdexcept>
#include <utility>

#include "internal/document_model_support.hpp"
#include "internal/paragraph_support.hpp"
#include "internal/style_support.hpp"

namespace docxcpp {

Run::Run(std::string text, RunStyle style)
    : text_(std::move(text)), style_(std::move(style)) {}

const std::string& Run::text() const noexcept { return text_; }

const RunStyle& Run::style() const noexcept { return style_; }

Paragraph::Paragraph(std::string text, ParagraphAlignment alignment, RunStyle first_run_style,
                     bool has_page_break, std::vector<Run> runs, std::string style_id,
                     int heading_level, ParagraphFormat format,
                     std::vector<HyperlinkInfo> hyperlinks,
                     std::shared_ptr<ParagraphBinding> binding)
    : text_(std::move(text)),
      alignment_(alignment),
      first_run_style_(std::move(first_run_style)),
      has_page_break_(has_page_break),
      runs_(std::move(runs)),
      style_id_(std::move(style_id)),
      heading_level_(heading_level),
      format_(std::move(format)),
      hyperlinks_(std::move(hyperlinks)),
      binding_(std::move(binding)) {}

const std::string& Paragraph::text() const noexcept { return text_; }

ParagraphAlignment Paragraph::alignment() const noexcept { return alignment_; }

const RunStyle& Paragraph::first_run_style() const noexcept { return first_run_style_; }

bool Paragraph::has_page_break() const noexcept { return has_page_break_; }

const std::vector<Run>& Paragraph::runs() const noexcept { return runs_; }

const std::string& Paragraph::style_id() const noexcept { return style_id_; }

int Paragraph::heading_level() const noexcept { return heading_level_; }

const ParagraphFormat& Paragraph::format() const noexcept { return format_; }

const std::vector<HyperlinkInfo>& Paragraph::hyperlinks() const noexcept { return hyperlinks_; }

Run Paragraph::add_run(const std::string& text, const RunStyle& style) {
  if (!binding_) {
    throw std::logic_error("paragraph is not bound to a writable document");
  }
  pugi::xml_node paragraph = paragraph_node_for_model_binding(*binding_);
  if (!paragraph) {
    throw std::out_of_range("bound paragraph no longer exists");
  }

  pugi::xml_node run = paragraph.append_child("w:r");
  apply_run_style_for_model(run, style, binding_->style_catalog.get());
  append_run_text_for_model(run, text);
  if (binding_->dirty != nullptr) {
    *binding_->dirty = true;
  }

  Run added_run(text, resolve_run_style_reference(style, binding_->style_catalog.get()));
  if (runs_.empty()) {
    first_run_style_ = style;
  }
  runs_.push_back(added_run);
  text_ += text;
  return added_run;
}

Paragraph Paragraph::insert_paragraph_before(const std::string& text, ParagraphAlignment alignment) {
  return insert_paragraph_before(std::vector<Run>{Run(text, {})}, alignment);
}

Paragraph Paragraph::insert_paragraph_before(const std::vector<Run>& runs,
                                             ParagraphAlignment alignment) {
  if (!binding_) {
    throw std::logic_error("paragraph is not bound to a writable document");
  }

  Paragraph inserted;
  insert_paragraph_before_binding(*binding_, runs, alignment, inserted);
  return inserted;
}

}  // namespace docxcpp
