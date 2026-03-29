#include "docxcpp/table.hpp"

#include <algorithm>
#include <stdexcept>
#include <utility>
#include <vector>

namespace docxcpp {

TableCell::TableCell(std::string text, std::size_t grid_span, std::string vertical_merge,
                     std::vector<Table> nested_tables)
    : text_(std::move(text)),
      grid_span_(grid_span),
      vertical_merge_(std::move(vertical_merge)),
      nested_tables_(std::move(nested_tables)) {}

const std::string& TableCell::text() const noexcept { return text_; }

std::size_t TableCell::grid_span() const noexcept { return grid_span_; }

const std::string& TableCell::vertical_merge() const noexcept { return vertical_merge_; }

const std::vector<Table>& TableCell::nested_tables() const noexcept { return nested_tables_; }

TableRow::TableRow(std::vector<TableCell> cells) : cells_(std::move(cells)) {}

const std::vector<TableCell>& TableRow::cells() const noexcept { return cells_; }

std::size_t TableRow::column_count() const noexcept {
  std::size_t columns = 0;
  for (const TableCell& cell : cells_) {
    columns += std::max<std::size_t>(1, cell.grid_span());
  }
  return columns;
}

Table::Table(std::vector<TableRow> rows, std::string style_id, std::string style_name)
    : rows_(std::move(rows)),
      style_id_(std::move(style_id)),
      style_name_(std::move(style_name)) {}

const std::vector<TableRow>& Table::rows() const noexcept { return rows_; }

std::size_t Table::row_count() const noexcept { return rows_.size(); }

std::size_t Table::column_count() const noexcept {
  std::size_t columns = 0;
  for (const TableRow& row : rows_) {
    columns = std::max(columns, row.column_count());
  }
  return columns;
}

const std::string& Table::style_id() const noexcept { return style_id_; }

const std::string& Table::style_name() const noexcept { return style_name_; }

const TableCell& Table::cell(std::size_t row_index, std::size_t col_index) const {
  if (row_index >= rows_.size()) {
    throw std::out_of_range("row_index out of range");
  }

  std::vector<std::vector<const TableCell*>> logical_rows;
  logical_rows.reserve(rows_.size());
  for (std::size_t current_row_index = 0; current_row_index < rows_.size(); ++current_row_index) {
    const TableRow& row = rows_[current_row_index];
    std::vector<const TableCell*> logical_row;
    logical_row.reserve(row.column_count());

    for (const TableCell& cell : row.cells()) {
      const std::size_t span = std::max<std::size_t>(1, cell.grid_span());
      if (cell.vertical_merge() == "continue" && !logical_rows.empty()) {
        for (std::size_t offset = 0; offset < span; ++offset) {
          const std::size_t logical_col = logical_row.size();
          if (logical_col < logical_rows.back().size() && logical_rows.back()[logical_col] != nullptr) {
            logical_row.push_back(logical_rows.back()[logical_col]);
          } else {
            logical_row.push_back(&cell);
          }
        }
        continue;
      }

      for (std::size_t offset = 0; offset < span; ++offset) {
        logical_row.push_back(&cell);
      }
    }

    logical_rows.push_back(std::move(logical_row));
  }

  if (col_index >= logical_rows[row_index].size()) {
    throw std::out_of_range("col_index out of range");
  }
  return *logical_rows[row_index][col_index];
}

}  // namespace docxcpp
