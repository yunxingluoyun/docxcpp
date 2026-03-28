# docxcpp 与 python-docx 功能对比

这份文档用于对比当前 `cpp/` 目录下的 C++ 版本 `docxcpp` 与本仓库上游 `python-docx` 的功能覆盖情况。

对比基准：

- `docxcpp`：当前 `cpp/include/docxcpp/*.hpp`、`cpp/src/*.cpp`、`cpp/tests/*.cpp` 中已经实现并测试的能力
- `python-docx`：本仓库 `docs/user/`、`src/docx/`、`tests/` 中体现出的公开能力与主要功能面

状态说明：

- `已支持`：C++ 版本已经实现，且已有当前测试覆盖
- `部分支持`：C++ 版本有对应能力，但范围明显小于 `python-docx`，或只提供较底层/较简化的接口
- `未支持`：当前 C++ 版本还没有实现对应能力

## 总体结论

当前 `docxcpp` 已经具备一个“可用的、面向受控场景”的子集能力，已经支持：

- 创建、打开、保存 `.docx`
- 添加段落、标题、分页符
- 基础 run 样式设置
- 创建表格、写入单元格、单元格追加段落、单元格合并
- 在正文和表格单元格中插入图片
- 设置页面尺寸和页边距
- 读取段落、run、标题、表格、图片以及基础页面设置

但从整体能力看，`python-docx` 仍然明显更完整，更接近“完整的 Word 文档对象模型”；而 `docxcpp` 当前更像是一个“实用型文档生成 + 部分读取库”。

## 功能矩阵

| 功能项 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 打开现有文档 | 已支持 | 已支持 | 两边都能打开 `.docx`。 |
| 从默认模板创建空文档 | 已支持 | 已支持 | C++ 版本使用内置模板目录。 |
| 保存文档 | 已支持 | 已支持 | 两边都能输出 `.docx`。 |
| 追加段落 | 已支持 | 已支持 | 对应 `docxcpp::Document::add_paragraph()`。 |
| 在已有段落前插入段落 | 已支持 | 未支持 | `python-docx` 有 `Paragraph.insert_paragraph_before()`。 |
| 追加标题/Title | 已支持 | 已支持 | C++ 当前支持 `0-9` 级标题。 |
| 按样式名/样式对象应用段落样式 | 已支持 | 部分支持 | C++ 目前只间接支持标题样式和对齐，不支持任意命名段落样式。 |
| 插入分页符 | 已支持 | 已支持 | C++ 有 `Document::add_page_break()`。 |
| 读取段落列表 | 已支持 | 已支持 | C++ 有 `Document::paragraphs()`。 |
| 读取段落文本 | 已支持 | 已支持 | C++ 有 `Paragraph::text()`。 |
| 段落对齐 | 已支持 | 部分支持 | C++ 当前支持 left/center/right/justify/inherit。 |
| 段落格式对象 | 已支持 | 未支持 | C++ 还没有类似 `ParagraphFormat` 的对象。 |
| 段落缩进 | 已支持 | 未支持 | 还没有左缩进/右缩进/首行缩进 API。 |
| 段前段后间距 | 已支持 | 未支持 | 还没有 spacing API。 |
| 行距 | 已支持 | 未支持 | 还没有 line spacing API。 |
| 分页行为控制 | 已支持 | 未支持 | 如 `keep_together`、`page_break_before` 等尚未实现。 |
| 读取段落样式 | 已支持 | 部分支持 | C++ 当前支持 `Paragraph::style_id()` 和 `Paragraph::heading_level()`。 |
| 读取 run 列表 | 已支持 | 已支持 | C++ 当前支持 `Paragraph::runs()`。 |
| 向已有段落追加 run | 已支持 | 未支持 | `python-docx` 有 `Paragraph.add_run()`，C++ 还没有。 |
| run 文本 | 已支持 | 已支持 | C++ 当前主要是读取侧支持；写入侧仍以“单次添加段落生成单个 run”为主。 |
| run 粗体/斜体/下划线 | 已支持 | 部分支持 | C++ 支持这三项，但没有更完整的字体模型。 |
| run 字号 | 已支持 | 已支持 | C++ 当前用整数 pt。 |
| run 字体名 | 已支持 | 已支持 | C++ 已支持基础写入/读取。 |
| run 字体颜色 | 已支持 | 已支持 | C++ 当前支持 RGB 十六进制颜色。 |
| run 高亮 | 已支持 | 已支持 | C++ 当前支持 Word 高亮 token。 |
| run 中的 tab 文本 | 已支持 | 部分支持 | C++ 读取时能映射 `w:tab` 为 `\\t`，但没有 tab stop 写入 API。 |
| 超链接 | 已支持 | 未支持 | C++ 还没有 hyperlink 模型和写入接口。 |
| 注释 comments | 已支持 | 未支持 | C++ 还没有 comments part、锚点、元数据等。 |
| 创建表格 | 已支持 | 已支持 | C++ 有 `Document::add_table()`。 |
| 表格样式设置 | 已支持 | 部分支持 | C++ 目前固定写 `TableGrid`，没有公开表格样式 API。 |
| 读取表格 | 已支持 | 已支持 | C++ 有 `Document::tables()`。 |
| 表格行/列集合 | 已支持 | 部分支持 | C++ 目前有行和单元格读取，没有单独的列对象。 |
| 表格创建后继续加行/列 | 已支持 | 未支持 | C++ 还没有增量结构编辑 API。 |
| 通过对象模型访问单元格 | 已支持 | 部分支持 | C++ 读取侧用 `rows()/cells()`，写入侧更多是基于索引。 |
| 单元格文本 | 已支持 | 已支持 | C++ 有 `TableCell::text()` 与 `set_table_cell()`。 |
| 单元格中多个段落 | 已支持 | 已支持 | C++ 有 `add_table_cell_paragraph()`，也能读回拼接文本。 |
| 嵌套表格 | 已支持 | 未支持 | C++ 还没有 nested table 的读写支持。 |
| 合并单元格 | 已支持 | 部分支持 | C++ 可写合并，也能读 `grid_span` / `vertical_merge`，但缺少完整 layout-grid 抽象。 |
| 省略单元格 / 布局网格辅助 | 已支持 | 未支持 | C++ 还没有类似 `grid_cols_before` / `grid_cols_after` 的能力。 |
| 正文插入内联图片 | 已支持 | 已支持 | 当前 C++ 支持 PNG/JPEG。 |
| 表格单元格内插入图片 | 已支持 | 已支持 | C++ 已有单独 helper。 |
| 图片尺寸设置 | 已支持 | 已支持 | C++ 支持宽高 pt，并能在只给一边时保持纵横比。 |
| 从 file-like 对象插图 | 已支持 | 未支持 | C++ 目前只接受文件路径。 |
| 浮动图片/浮动形状 | 未支持 | 未支持 | 这部分两边都不是常规 API 支持范围。 |
| 读取图片元数据 | 部分支持 | 已支持 | C++ 当前能读尺寸、关系 id、包内目标路径、所在位置。 |
| 统计图片数量 | 部分支持 | 已支持 | C++ 有 `Document::image_count()`。 |
| section 集合 | 已支持 | 未支持 | C++ 目前还没有真正的 section 对象模型。 |
| 新增 section | 已支持 | 未支持 | C++ 暂无多 section 写入 API。 |
| 页面尺寸 | 已支持 | 已支持 | C++ 已支持读写页面宽高。 |
| 页面方向 | 已支持 | 未支持 | C++ 还没有 orientation API。 |
| 页边距 | 已支持 | 部分支持 | C++ 当前只支持 top/right/bottom/left，未覆盖 gutter/header/footer 距离。 |
| 页眉页脚内容 | 已支持 | 未支持 | C++ 还没有 header/footer part 与引用处理。 |
| 样式集合 | 已支持 | 未支持 | C++ 还没有 document styles 模型。 |
| 新增/删除自定义样式 | 已支持 | 未支持 | C++ 当前没有对应能力。 |
| 按名称应用表格样式 | 已支持 | 未支持 | C++ 暂时只固定输出 `TableGrid`。 |
| 按名称应用字符样式 | 已支持 | 未支持 | C++ 还没有 style hierarchy。 |
| 文档 settings 对象 | 已支持 | 未支持 | C++ 还没有 settings API。 |

## 当前 C++ 版本已经做得不错的地方

相对于很多“第一版移植”，当前 `docxcpp` 已经不只是简单导出文本，而是具备了比较实用的结构能力：

- 能生成有效 `.docx`，覆盖段落、标题、分页、表格、内联图片
- 已支持较实用的直接格式：
  - 粗体
  - 斜体
  - 下划线
  - 字号
  - RGB 字体颜色
  - 字体名
  - 高亮
  - 段落对齐
- 可以在正文和表格单元格里插入图片
- 已经具备一批读取侧 API，而不是只会写：
  - 段落
  - run
  - 标题样式信息
  - 表格
  - 合并单元格元数据
  - 图片信息和位置
  - 页面尺寸和页边距
- 当前输出已经通过：
  - C++ 单元测试验证
  - `python-docx` 读取验证

## 相比 python-docx 的主要差距

当前差距最大的几个方向如下。

### 1. 富文本编辑对象模型不够完整

`python-docx` 可以直接操作 `Paragraph` / `Run` 对象，向现有段落追加 run，设置更多字体属性，也支持超链接等更复杂内容。

当前 `docxcpp` 还主要是以 `Document` 级高层方法为主，读取对象偏轻量，写入对象还不够细。

### 2. 段落格式能力不足

`python-docx` 的段落格式支持非常完整，包括：

- 缩进
- tab stops
- 段前段后间距
- 行距
- 分页行为控制

而当前 C++ 版本主要只有：

- 对齐
- 页分页符检测

### 3. 表格模型还不够丰富

`python-docx` 的表格对象模型更完整，支持：

- 更完整的 row/column/cell 抽象
- 动态加行
- 嵌套表格
- 更贴近 Word 布局网格的辅助能力

当前 `docxcpp` 已支持“实用写表格 + 读取基础结构 + 合并信息”，但还不算完整表格 DOM。

### 4. section 与页面布局模型还比较弱

`python-docx` 支持：

- 多个 section
- section break
- 页面方向
- 页眉页脚距离
- 页眉页脚内容访问

当前 `docxcpp` 还主要停留在单 section 的基础页面大小和四个页边距。

### 5. 样式系统与高级文档功能还没有铺开

`python-docx` 还覆盖了：

- 文档样式集合
- 自定义样式
- comments
- settings
- 更广泛的文档 part

这些在当前 C++ 版本里还没有建立起来。

## 读取侧能力对比

如果只看“读取侧 API”，当前状态是：

### python-docx

- 更完整
- 更面向对象
- 更贴近 Word 原始文档模型
- 覆盖 styles、sections、comments、headers/footers 等更多部件

### docxcpp

在当前目标子集内，读取侧已经比较实用，当前已暴露：

- `Document::paragraphs()`
- `Paragraph::text()`
- `Paragraph::alignment()`
- `Paragraph::first_run_style()`
- `Paragraph::runs()`
- `Paragraph::style_id()`
- `Paragraph::heading_level()`
- `Paragraph::has_page_break()`
- `Document::tables()`
- `TableRow::cells()`
- `TableCell::text()`
- `TableCell::grid_span()`
- `TableCell::vertical_merge()`
- `Document::page_size_pt()`
- `Document::page_margins_pt()`
- `Document::picture_sizes_pt()`
- `Document::pictures()`
- `Document::image_count()`

## 兼容性说明

- 当前 `python-docx` 仍然是“功能参考实现”，`docxcpp` 还不是它的 API 兼容版本
- 当前 C++ API 明显更简单，也更偏索引式和实用主义
- 当前 C++ 生成的文档已经通过 `python-docx` 读取验证，说明在目前已实现功能范围内，输出是可用的
- 但在表格合并、段落格式、样式继承等更高层语义上，C++ 版本和 `python-docx` 仍有明显差距

## 建议的后续优先级

如果目标是逐步向 `python-docx` 靠近，建议优先做下面几项。

### 1. 补齐可写的 paragraph/run API

建议优先加上：

- `Paragraph::add_run()`
- 同一段落内多 run 混合格式
- 更完整的字体属性

这是最能明显提升“可用性”的一组能力。

### 2. 扩展段落格式

优先补：

- 缩进
- 段前段后间距
- 行距
- page-break-before 等分页属性

### 3. 扩展 section API

优先补：

- 页面方向
- 多 section 支持
- header/footer 引用和内容

### 4. 加强表格模型

优先补：

- 动态加行
- 嵌套表格
- 面向布局网格的读取辅助

### 5. 建立样式系统

后续可以逐步加入：

- document styles 集合
- 按名称应用段落/表格/字符样式
- 自定义样式创建

## 一句话总结

`python-docx` 目前仍然远比 `docxcpp` 完整，更接近完整 Word 文档对象模型。

但 `docxcpp` 已经达到一个相当实用的阶段：对于受控模板、报表导出、结构化文档生成这类场景，已经能处理文本、格式、表格、图片和基础页面设置。如果后续继续补齐 run 写入、段落格式、section、header/footer、styles 等能力，它会逐步逼近 `python-docx` 的主干能力范围。
