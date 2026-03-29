#pragma once

#include <cmath>
#include <cstdint>
#include <vector>

#include "docxcpp/export.hpp"

namespace docxcpp {

/**
 * @brief 长度值，内部以 twips 存储，可由多种常见单位构造。
 */
class DOCXCPP_API Length {
public:
  Length() = default;

  static Length from_twips(std::int64_t twips) noexcept;
  static Length from_pt(double pt) noexcept;
  static Length from_inches(double inches) noexcept;
  static Length from_cm(double cm) noexcept;
  static Length from_mm(double mm) noexcept;

  std::int64_t twips() const noexcept;
  double pt() const noexcept;
  double inches() const noexcept;
  double cm() const noexcept;
  double mm() const noexcept;

private:
  explicit Length(std::int64_t twips) noexcept;

  std::int64_t twips_{0};
};

/**
 * @brief python-docx 风格的长度构造器，单位为毫米。
 */
inline Length Mm(double mm) noexcept { return Length::from_mm(mm); }

/**
 * @brief python-docx 风格的长度构造器，单位为厘米。
 */
inline Length Cm(double cm) noexcept { return Length::from_cm(cm); }

/**
 * @brief python-docx 风格的长度构造器，单位为英寸。
 */
inline Length Inches(double inches) noexcept { return Length::from_inches(inches); }

/**
 * @brief python-docx 风格的长度构造器，单位为点。
 */
inline Length Pt(double pt) noexcept { return Length::from_pt(pt); }

/**
 * @brief python-docx 风格的长度构造器，单位为 twips。
 */
inline Length Twips(double twips) noexcept {
  return Length::from_twips(static_cast<std::int64_t>(std::llround(twips)));
}

/**
 * @brief 页面尺寸。
 */
struct PageSize {
  Length width;
  Length height;
};

/**
 * @brief 标准纸张尺寸预设。
 */
enum class StandardPageSize {
  A0,
  A1,
  A2,
  A3,
  A4,
  A5,
  A6,
  A7,
  A8,
  A9,
  A10,
  B0,
  B1,
  B2,
  B3,
  B4,
  B5,
  B6,
  B7,
  B8,
  B9,
  B10,
  Letter,
  Legal,
  Tabloid,
  Ledger,
  Executive,
  Statement,
  Folio,
  Quarto,
};

/**
 * @brief 根据标准纸张预设生成页面尺寸。
 * @param standard_size 标准纸张预设。
 * @return 对应页面尺寸。
 */
DOCXCPP_API PageSize standard_page_size(StandardPageSize standard_size);

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
 * @brief 页面边距与页眉页脚距离。
 */
struct PageMargins {
  Length top;
  Length right;
  Length bottom;
  Length left;
  Length header;
  Length footer;
  Length gutter;
};

class DOCXCPP_API Section {
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
   * @brief 使用标准纸张预设构造一个 section 配置对象。
   * @param standard_size 标准纸张预设。
   * @param orientation 页面方向。
   * @param margins 页边距与页眉页脚距离。
   */
  explicit Section(StandardPageSize standard_size,
                   PageOrientation orientation = PageOrientation::Portrait,
                   PageMargins margins = {});

  /**
   * @brief 返回页面尺寸。
   * @return 页面尺寸的只读引用。
   */
  const PageSize& page_size() const noexcept;
  /**
   * @brief 返回页面方向。
   * @return 页面方向枚举值。
   */
  PageOrientation page_orientation() const noexcept;
  /**
   * @brief 返回页面边距配置。
   * @return 页边距配置的只读引用。
   */
  const PageMargins& page_margins() const noexcept;

private:
  PageSize page_size_;
  PageOrientation orientation_;
  PageMargins margins_;
};

}  // namespace docxcpp
