#pragma once

#include <vector>

namespace docxcpp {

struct PageSize {
  int width_pt{0};
  int height_pt{0};
};

enum class PageOrientation {
  Portrait,
  Landscape,
};

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
  explicit Section(PageSize page_size = {}, PageOrientation orientation = PageOrientation::Portrait,
                   PageMargins margins = {});

  const PageSize& page_size_pt() const noexcept;
  PageOrientation page_orientation() const noexcept;
  const PageMargins& page_margins_pt() const noexcept;

private:
  PageSize page_size_;
  PageOrientation orientation_;
  PageMargins margins_;
};

}  // namespace docxcpp
