#include "docxcpp/section.hpp"

#include <utility>

namespace docxcpp {

Section::Section(PageSize page_size, PageOrientation orientation, PageMargins margins)
    : page_size_(std::move(page_size)),
      orientation_(orientation),
      margins_(std::move(margins)) {}

const PageSize& Section::page_size_pt() const noexcept { return page_size_; }

PageOrientation Section::page_orientation() const noexcept { return orientation_; }

const PageMargins& Section::page_margins_pt() const noexcept { return margins_; }

}  // namespace docxcpp
