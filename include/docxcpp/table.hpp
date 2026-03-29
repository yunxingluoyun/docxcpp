#pragma once

#include <string>
#include <vector>

namespace docxcpp {

class Table;

class TableCell {
public:
  /**
   * @brief 构造一个表格单元格读模型。
   * @param text 单元格文本。
   * @param grid_span 横向跨列数。
   * @param vertical_merge 纵向合并标记。
   * @param nested_tables 单元格内的嵌套表格。
   */
  explicit TableCell(std::string text = {}, std::size_t grid_span = 1,
                     std::string vertical_merge = {}, std::vector<Table> nested_tables = {});

  /**
   * @brief 返回单元格文本。
   * @return 单元格文本的只读引用。
   */
  const std::string& text() const noexcept;
  /**
   * @brief 返回横向跨列数。
   * @return 横向合并后的逻辑列跨度。
   */
  std::size_t grid_span() const noexcept;
  /**
   * @brief 返回纵向合并标记。
   * @return 常见值为 restart 或 continue；未合并时为空字符串。
   */
  const std::string& vertical_merge() const noexcept;
  /**
   * @brief 返回单元格内的嵌套表格。
   * @return 嵌套表格列表的只读引用。
   */
  const std::vector<Table>& nested_tables() const noexcept;

private:
  std::string text_;
  std::size_t grid_span_;
  std::string vertical_merge_;
  std::vector<Table> nested_tables_;
};

class TableRow {
public:
  /**
   * @brief 构造一个表格行读模型。
   * @param cells 该行包含的物理单元格列表。
   */
  explicit TableRow(std::vector<TableCell> cells = {});

  /**
   * @brief 返回该行的物理单元格列表。
   * @return 物理单元格列表的只读引用。
   */
  const std::vector<TableCell>& cells() const noexcept;
  /**
   * @brief 返回该行按 grid span 展开的逻辑列数。
   * @return 逻辑列数。
   */
  std::size_t column_count() const noexcept;

private:
  std::vector<TableCell> cells_;
};

class Table {
public:
  /**
   * @brief 构造一个表格读模型。
   * @param rows 表格中的物理行列表。
   */
  explicit Table(std::vector<TableRow> rows = {});

  /**
   * @brief 返回表格中的物理行列表。
   * @return 行列表的只读引用。
   */
  const std::vector<TableRow>& rows() const noexcept;
  /**
   * @brief 返回物理行数。
   * @return 物理行数量。
   */
  std::size_t row_count() const noexcept;
  /**
   * @brief 返回最大逻辑列数。
   * @return 表格中的最大逻辑列数。
   */
  std::size_t column_count() const noexcept;
  /**
   * @brief 按逻辑网格坐标访问单元格。
   * @param row_index 逻辑行索引，从 0 开始。
   * @param col_index 逻辑列索引，从 0 开始。
   * @return 对应单元格的只读引用；会考虑横向和纵向合并。
   */
  const TableCell& cell(std::size_t row_index, std::size_t col_index) const;

private:
  std::vector<TableRow> rows_;
};

}  // namespace docxcpp
