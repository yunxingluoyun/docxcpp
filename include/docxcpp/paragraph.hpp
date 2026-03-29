#pragma once

#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace docxcpp {

/**
 * @brief 段落对齐方式。
 */
enum class ParagraphAlignment {
  Inherit,
  Left,
  Center,
  Right,
  Justify,
};

/**
 * @brief 单个 run 的字符样式。
 */
struct RunStyle {
  bool bold{false};
  bool italic{false};
  bool underline{false};
  int font_size_pt{0};
  std::string color_hex;
  std::string font_name;
  std::string highlight;
};

/**
 * @brief 制表位对齐方式。
 */
enum class TabAlignment {
  Left,
  Center,
  Right,
  Decimal,
  Bar,
};

/**
 * @brief 制表位前导符样式。
 */
enum class TabLeader {
  Spaces,
  Dots,
  Dashes,
  Lines,
  Heavy,
  MiddleDot,
};

/**
 * @brief 一个制表位定义，位置单位为 pt。
 */
struct TabStop {
  int position_pt{0};
  TabAlignment alignment{TabAlignment::Left};
  TabLeader leader{TabLeader::Spaces};
};

/**
 * @brief 行距模式。
 */
enum class LineSpacingMode {
  Exact,
  AtLeast,
  Multiple,
};

/**
 * @brief 段落格式集合，长度单位均为 pt。
 */
struct ParagraphFormat {
  std::optional<int> left_indent_pt;
  std::optional<int> right_indent_pt;
  std::optional<int> first_line_indent_pt;
  std::optional<int> space_before_pt;
  std::optional<int> space_after_pt;
  std::optional<int> line_spacing_pt;
  std::optional<double> line_spacing_multiple;
  std::optional<LineSpacingMode> line_spacing_mode;
  std::vector<TabStop> tab_stops;
  std::optional<bool> keep_together;
  std::optional<bool> keep_with_next;
  std::optional<bool> page_break_before;
};

/**
 * @brief 超链接读取结果。
 */
struct HyperlinkInfo {
  std::string text;
  std::string url;
};

struct ParagraphBinding;

class Run {
public:
  /**
   * @brief 构造一个文本 run。
   * @param text run 文本内容。
   * @param style run 的字符样式。
   */
  explicit Run(std::string text = {}, RunStyle style = {});

  /**
   * @brief 返回 run 文本内容。
   * @return run 文本的只读引用。
   */
  const std::string& text() const noexcept;
  /**
   * @brief 返回 run 的字符样式。
   * @return run 样式的只读引用。
   */
  const RunStyle& style() const noexcept;

private:
  std::string text_;
  RunStyle style_;
};

class Paragraph {
public:
  /**
   * @brief 构造一个段落对象。
   * @param text 段落纯文本。
   * @param alignment 段落对齐方式。
   * @param first_run_style 首个 run 的样式快照。
   * @param has_page_break 段落是否包含分页符。
   * @param runs 段落中的 runs。
   * @param style_id 段落样式 ID。
   * @param heading_level 标题级别，非标题通常为 -1。
   * @param format 段落格式。
   * @param hyperlinks 段落中的超链接信息。
   * @param binding 指向底层 XML 的绑定，用于可变写入。
   */
  explicit Paragraph(std::string text = {}, ParagraphAlignment alignment = ParagraphAlignment::Inherit,
                     RunStyle first_run_style = {}, bool has_page_break = false,
                     std::vector<Run> runs = {}, std::string style_id = {}, int heading_level = -1,
                     ParagraphFormat format = {}, std::vector<HyperlinkInfo> hyperlinks = {},
                     std::shared_ptr<ParagraphBinding> binding = {});

  /**
   * @brief 返回段落纯文本内容。
   * @return 段落文本的只读引用。
   */
  const std::string& text() const noexcept;
  /**
   * @brief 返回段落对齐方式。
   * @return 段落对齐枚举值。
   */
  ParagraphAlignment alignment() const noexcept;
  /**
   * @brief 返回首个 run 的样式。
   * @return 首个 run 样式的只读引用。
   */
  const RunStyle& first_run_style() const noexcept;
  /**
   * @brief 返回该段落是否包含分页符。
   * @return 如果包含分页符则为 true。
   */
  bool has_page_break() const noexcept;
  /**
   * @brief 返回段落中的所有 runs。
   * @return run 列表的只读引用。
   */
  const std::vector<Run>& runs() const noexcept;
  /**
   * @brief 返回段落样式 ID，例如 Heading1、Title。
   * @return 样式 ID 的只读引用。
   */
  const std::string& style_id() const noexcept;
  /**
   * @brief 返回标题级别。
   * @return 标题级别；非标题通常为 -1。
   */
  int heading_level() const noexcept;
  /**
   * @brief 返回段落格式信息。
   * @return 段落格式的只读引用。
   */
  const ParagraphFormat& format() const noexcept;
  /**
   * @brief 返回段落中的超链接信息。
   * @return 超链接信息列表的只读引用。
   */
  const std::vector<HyperlinkInfo>& hyperlinks() const noexcept;
  /**
   * @brief 向现有段落末尾追加一个 run，并同步写回文档。
   * @param text 要追加的文本。
   * @param style 要追加的字符样式。
   * @return 新增的 run 对象快照。
   */
  Run add_run(const std::string& text, const RunStyle& style = {});

private:
  std::string text_;
  ParagraphAlignment alignment_;
  RunStyle first_run_style_;
  bool has_page_break_;
  std::vector<Run> runs_;
  std::string style_id_;
  int heading_level_;
  ParagraphFormat format_;
  std::vector<HyperlinkInfo> hyperlinks_;
  std::shared_ptr<ParagraphBinding> binding_;
};

}  // namespace docxcpp
