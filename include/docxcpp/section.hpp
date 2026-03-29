#pragma once

#include <vector>

namespace docxcpp {

/**
 * @brief 页面尺寸，单位为 pt。
 */
struct PageSize {
  int width_pt{0};
  int height_pt{0};
};

/**
 * @brief 页面方向。
 */
enum class PageOrientation {
  Portrait,
  Landscape,
};

/**
 * @brief 页眉页脚类型。
 */
enum class HeaderFooterType {
  Default,
  FirstPage,
  EvenPage,
};

/**
 * @brief 页面边距与页眉页脚距离，单位为 pt。
 */
struct PageMargins {
  int top_pt{0};
  int right_pt{0};
  int bottom_pt{0};
  int left_pt{0};
  int header_pt{0};
  int footer_pt{0};
  int gutter_pt{0};
};

class Section {
public:
  /**
   * @brief 构造一个 section 配置对象。
   * @param page_size 页面尺寸。
   * @param orientation 页面方向。
   * @param margins 页边距与页眉页脚距离。
   */
  explicit Section(PageSize page_size = {}, PageOrientation orientation = PageOrientation::Portrait,
                   PageMargins margins = {});

  /**
   * @brief 返回页面尺寸。
   * @return 页面尺寸的只读引用。
   */
  const PageSize& page_size_pt() const noexcept;
  /**
   * @brief 返回页面方向。
   * @return 页面方向枚举值。
   */
  PageOrientation page_orientation() const noexcept;
  /**
   * @brief 返回页面边距配置。
   * @return 页边距配置的只读引用。
   */
  const PageMargins& page_margins_pt() const noexcept;

private:
  PageSize page_size_;
  PageOrientation orientation_;
  PageMargins margins_;
};

}  // namespace docxcpp
