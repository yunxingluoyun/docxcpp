#include "docxcpp/table.hpp"

#include <utility>

namespace docxcpp {

TableCell::TableCell(std::string text, std::size_t grid_span, std::string vertical_merge)
    : text_(std::move(text)),
      grid_span_(grid_span),
      vertical_merge_(std::move(vertical_merge)) {}

const std::string& TableCell::text() const noexcept { return text_; }

std::size_t TableCell::grid_span() const noexcept { return grid_span_; }

const std::string& TableCell::vertical_merge() const noexcept { return vertical_merge_; }

TableRow::TableRow(std::vector<TableCell> cells) : cells_(std::move(cells)) {}

const std::vector<TableCell>& TableRow::cells() const noexcept { return cells_; }

Table::Table(std::vector<TableRow> rows) : rows_(std::move(rows)) {}

const std::vector<TableRow>& Table::rows() const noexcept { return rows_; }

}  // namespace docxcpp
