#include "docxcpp/section.hpp"

#include <cmath>
#include <utility>

namespace docxcpp {

namespace {

constexpr double kTwipsPerPoint = 20.0;
constexpr double kPointsPerInch = 72.0;
constexpr double kCmPerInch = 2.54;
constexpr double kMmPerInch = 25.4;

}  // namespace

Length::Length(std::int64_t twips) noexcept : twips_(twips) {}

Length Length::from_twips(std::int64_t twips) noexcept { return Length(twips); }

Length Length::from_pt(double pt) noexcept {
  return Length(static_cast<std::int64_t>(std::llround(pt * kTwipsPerPoint)));
}

Length Length::from_inches(double inches) noexcept {
  return from_pt(inches * kPointsPerInch);
}

Length Length::from_cm(double cm) noexcept { return from_inches(cm / kCmPerInch); }

Length Length::from_mm(double mm) noexcept { return from_inches(mm / kMmPerInch); }

std::int64_t Length::twips() const noexcept { return twips_; }

double Length::pt() const noexcept { return static_cast<double>(twips_) / kTwipsPerPoint; }

double Length::inches() const noexcept { return pt() / kPointsPerInch; }

double Length::cm() const noexcept { return inches() * kCmPerInch; }

double Length::mm() const noexcept { return inches() * kMmPerInch; }

PageSize standard_page_size(StandardPageSize standard_size) {
  switch (standard_size) {
    case StandardPageSize::A0:
      return PageSize{Length::from_mm(841), Length::from_mm(1189)};
    case StandardPageSize::A1:
      return PageSize{Length::from_mm(594), Length::from_mm(841)};
    case StandardPageSize::A2:
      return PageSize{Length::from_mm(420), Length::from_mm(594)};
    case StandardPageSize::A3:
      return PageSize{Length::from_mm(297), Length::from_mm(420)};
    case StandardPageSize::A4:
      return PageSize{Length::from_mm(210), Length::from_mm(297)};
    case StandardPageSize::A5:
      return PageSize{Length::from_mm(148), Length::from_mm(210)};
    case StandardPageSize::A6:
      return PageSize{Length::from_mm(105), Length::from_mm(148)};
    case StandardPageSize::A7:
      return PageSize{Length::from_mm(74), Length::from_mm(105)};
    case StandardPageSize::A8:
      return PageSize{Length::from_mm(52), Length::from_mm(74)};
    case StandardPageSize::A9:
      return PageSize{Length::from_mm(37), Length::from_mm(52)};
    case StandardPageSize::A10:
      return PageSize{Length::from_mm(26), Length::from_mm(37)};
    case StandardPageSize::B0:
      return PageSize{Length::from_mm(1000), Length::from_mm(1414)};
    case StandardPageSize::B1:
      return PageSize{Length::from_mm(707), Length::from_mm(1000)};
    case StandardPageSize::B2:
      return PageSize{Length::from_mm(500), Length::from_mm(707)};
    case StandardPageSize::B3:
      return PageSize{Length::from_mm(353), Length::from_mm(500)};
    case StandardPageSize::B4:
      return PageSize{Length::from_mm(250), Length::from_mm(353)};
    case StandardPageSize::B5:
      return PageSize{Length::from_mm(176), Length::from_mm(250)};
    case StandardPageSize::B6:
      return PageSize{Length::from_mm(125), Length::from_mm(176)};
    case StandardPageSize::B7:
      return PageSize{Length::from_mm(88), Length::from_mm(125)};
    case StandardPageSize::B8:
      return PageSize{Length::from_mm(62), Length::from_mm(88)};
    case StandardPageSize::B9:
      return PageSize{Length::from_mm(44), Length::from_mm(62)};
    case StandardPageSize::B10:
      return PageSize{Length::from_mm(31), Length::from_mm(44)};
    case StandardPageSize::Letter:
      return PageSize{Length::from_inches(8.5), Length::from_inches(11.0)};
    case StandardPageSize::Legal:
      return PageSize{Length::from_inches(8.5), Length::from_inches(14.0)};
    case StandardPageSize::Tabloid:
      return PageSize{Length::from_inches(11.0), Length::from_inches(17.0)};
    case StandardPageSize::Ledger:
      return PageSize{Length::from_inches(17.0), Length::from_inches(11.0)};
    case StandardPageSize::Executive:
      return PageSize{Length::from_inches(7.25), Length::from_inches(10.5)};
    case StandardPageSize::Statement:
      return PageSize{Length::from_inches(5.5), Length::from_inches(8.5)};
    case StandardPageSize::Folio:
      return PageSize{Length::from_inches(8.5), Length::from_inches(13.0)};
    case StandardPageSize::Quarto:
      return PageSize{Length::from_mm(215), Length::from_mm(275)};
  }
  return {};
}

Section::Section(PageSize page_size, PageOrientation orientation, PageMargins margins)
    : page_size_(std::move(page_size)),
      orientation_(orientation),
      margins_(std::move(margins)) {}

Section::Section(StandardPageSize standard_size, PageOrientation orientation, PageMargins margins)
    : Section(standard_page_size(standard_size), orientation, std::move(margins)) {}

const PageSize& Section::page_size() const noexcept { return page_size_; }

PageOrientation Section::page_orientation() const noexcept { return orientation_; }

const PageMargins& Section::page_margins() const noexcept { return margins_; }

}  // namespace docxcpp
