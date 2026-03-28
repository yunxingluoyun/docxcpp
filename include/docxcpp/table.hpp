#pragma once

#include <string>
#include <vector>

namespace docxcpp {

class TableCell {
public:
  explicit TableCell(std::string text = {}, std::size_t grid_span = 1,
                     std::string vertical_merge = {});

  const std::string& text() const noexcept;
  std::size_t grid_span() const noexcept;
  const std::string& vertical_merge() const noexcept;

private:
  std::string text_;
  std::size_t grid_span_;
  std::string vertical_merge_;
};

class TableRow {
public:
  explicit TableRow(std::vector<TableCell> cells = {});

  const std::vector<TableCell>& cells() const noexcept;

private:
  std::vector<TableCell> cells_;
};

class Table {
public:
  explicit Table(std::vector<TableRow> rows = {});

  const std::vector<TableRow>& rows() const noexcept;

private:
  std::vector<TableRow> rows_;
};

}  // namespace docxcpp
