#pragma once

#include <string>
#include <vector>

namespace docxcpp {

enum class ParagraphAlignment {
  Inherit,
  Left,
  Center,
  Right,
  Justify,
};

struct RunStyle {
  bool bold{false};
  bool italic{false};
  bool underline{false};
  int font_size_pt{0};
  std::string color_hex;
  std::string font_name;
  std::string highlight;
};

class Run {
public:
  explicit Run(std::string text = {}, RunStyle style = {});

  const std::string& text() const noexcept;
  const RunStyle& style() const noexcept;

private:
  std::string text_;
  RunStyle style_;
};

class Paragraph {
public:
  explicit Paragraph(std::string text = {}, ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                     RunStyle first_run_style = {}, bool has_page_break = false,
                     std::vector<Run> runs = {}, std::string style_id = {}, int heading_level = -1);

  const std::string& text() const noexcept;
  ParagraphAlignment alignment() const noexcept;
  const RunStyle& first_run_style() const noexcept;
  bool has_page_break() const noexcept;
  const std::vector<Run>& runs() const noexcept;
  const std::string& style_id() const noexcept;
  int heading_level() const noexcept;

private:
  std::string text_;
  ParagraphAlignment alignment_;
  RunStyle first_run_style_;
  bool has_page_break_;
  std::vector<Run> runs_;
  std::string style_id_;
  int heading_level_;
};

}  // namespace docxcpp
