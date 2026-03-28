# docxcpp 与 python-docx 对比

这是一份基于当前代码状态重新整理的对比文档，用来说明 `cpp/` 下的 `docxcpp` 与本仓库上游 `python-docx` 之间，哪些能力已经接近，哪些还明显有差距。

对比依据：

- C++ 侧：`cpp/include/docxcpp/*.hpp`、`cpp/src/*.cpp`、`cpp/tests/*.cpp`
- Python 侧：`src/docx/`、`docs/user/` 与现有测试所体现的公开能力

本文中的状态说明：

- `已支持`：当前 C++ 版本已经实现，并且已有测试覆盖
- `部分支持`：C++ 有对应能力，但明显比 `python-docx` 窄，或者只提供更底层、更简化的接口
- `缺失`：当前 C++ 版本还没有公开能力

## 当前结论

`docxcpp` 已经不再只是一个“能导出几段文本”的雏形，它现在已经能覆盖一批很实用的 `.docx` 生成与读取场景：

- 打开、创建、保存 `.docx`
- 写段落、标题、分页符、多 run 混合内容
- 通过 `Paragraph::add_run()` 向已绑定段落继续追加 run
- 写一组比较实用的 run 样式
- 创建表格、写单元格、给单元格追加段落、合并单元格
- 从文件路径或内存字节缓冲插入图片
- 写外部超链接
- 写基础的段落级批注
- 设置页面尺寸、方向、页边距、页眉页脚距离、gutter
- 设置段落对齐、缩进、段前段后间距、固定行距、基础分页标志
- 读回段落、run、标题/样式信息、段落格式、表格、批注、图片元数据

但它仍然不是 `python-docx` 的等价替代。`python-docx` 依旧拥有更完整的 Word 文档对象模型、更强的编辑能力，以及更深的 style / section / header-footer / table 支持。

## 功能总览

### 一、文档生命周期

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 打开现有文档 | 已支持 | 已支持 | 两边都能读取 `.docx`。 |
| 创建空白文档 | 已支持 | 已支持 | C++ 使用仓库内置模板包。 |
| 保存文档 | 已支持 | 已支持 | 两边都能输出有效 `.docx`。 |

### 二、段落与 run

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 新增段落 | 已支持 | 已支持 | `Document::add_paragraph()`。 |
| 在已有段落前插入段落 | 已支持 | 缺失 | C++ 还没有可变的段落对象 API。 |
| 新增标题 / Title | 已支持 | 已支持 | C++ 当前支持 `0-9` 级。 |
| 插入分页符 | 已支持 | 已支持 | `Document::add_page_break()`。 |
| 一段内写多个 run | 已支持 | 已支持 | C++ 通过 `std::vector<Run>` 写入。 |
| 向已有段落对象继续追加 run | 已支持 | 部分支持 | 现在已有 `Paragraph::add_run()`，但可写对象模型仍然明显小于 `python-docx`。 |
| 读取段落列表 | 已支持 | 已支持 | `Document::paragraphs()`。 |
| 读取 run 列表 | 已支持 | 已支持 | `Paragraph::runs()`。 |
| 读取段落样式信息 | 已支持 | 部分支持 | C++ 目前只有 `style_id()` 和 `heading_level()`。 |

### 三、文本格式

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 粗体 / 斜体 / 下划线 | 已支持 | 部分支持 | C++ 支持这三项，但没有更完整的字体模型。 |
| 字号 | 已支持 | 已支持 | C++ 使用整数 pt。 |
| 字体名 | 已支持 | 已支持 | 支持基础 `ascii` / `hAnsi`。 |
| 字体颜色 | 已支持 | 已支持 | C++ 当前只支持 RGB 十六进制。 |
| 高亮 | 已支持 | 已支持 | 使用 Word highlight token。 |
| run 内换行 | 已支持 | 已支持 | C++ 已支持写入和读回。 |
| run 内 tab 文本 | 已支持 | 部分支持 | C++ 可读回 `w:tab -> \t`，但没有对应写入与 tab stop API。 |
| 按名称应用字符样式 | 已支持 | 缺失 | C++ 还没有 style hierarchy。 |

### 四、段落格式

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 对齐 | 已支持 | 部分支持 | C++ 当前支持 left / center / right / justify / inherit。 |
| 左缩进 / 右缩进 / 首行缩进 | 已支持 | 已支持 | C++ 已支持读写。 |
| 段前 / 段后间距 | 已支持 | 已支持 | C++ 已支持读写。 |
| 行距 | 已支持 | 部分支持 | C++ 当前支持固定行距。 |
| 分页行为标志 | 已支持 | 部分支持 | 已支持 `keep_together`、`keep_with_next`、`page_break_before`。 |
| tab stops | 已支持 | 缺失 | 还没有公开 API。 |
| widow/orphan control 等更细段落标志 | 已支持 | 缺失 | 当前还未实现。 |
| 完整段落格式对象模型 | 已支持 | 部分支持 | `ParagraphFormat` 已有，但范围刻意更小。 |

### 五、超链接与批注

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 写外部超链接 | 已支持 | 已支持 | 正文和表格单元格段落都能写。 |
| 读取超链接对象 | 已支持 | 部分支持 | C++ 现在通过 `Paragraph::hyperlinks()` 暴露段落级 `HyperlinkInfo` 快照，但还不是完整可编辑对象模型。 |
| 读取超链接目标地址 | 已支持 | 已支持 | 现在可通过 `Paragraph::hyperlinks()` 读回 URL。 |
| 写批注 | 已支持 | 部分支持 | C++ 当前支持基础的段落级批注插入。 |
| 读批注 | 已支持 | 部分支持 | `comments()` / `comment_count()` 读取 comments part，但不是完整锚点模型。 |
| 任意 run 范围批注锚定 | 已支持 | 缺失 | C++ 目前不是 run-range 级别。 |
| 回复 / 已解决状态 / 更完整批注编辑 | 已支持 | 缺失 | 还没有实现。 |

### 六、表格

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 创建表格 | 已支持 | 已支持 | `Document::add_table()`。 |
| 读取表格 | 已支持 | 已支持 | `Document::tables()`。 |
| 写单元格文本 | 已支持 | 已支持 | 支持纯文本和多 run。 |
| 单元格内继续追加段落 | 已支持 | 已支持 | `add_table_cell_paragraph()`。 |
| 合并单元格 | 已支持 | 部分支持 | C++ 支持写合并，并读回 `grid_span` / `vertical_merge`。 |
| 表格样式选择 | 已支持 | 部分支持 | C++ 目前固定写 `TableGrid`。 |
| 创建后继续加行 / 列 | 已支持 | 缺失 | 没有增量结构编辑 API。 |
| 嵌套表格 | 已支持 | 缺失 | 当前不支持。 |
| 面向布局网格的辅助能力 | 已支持 | 缺失 | 没有 `python-docx` 那类更高层 helper。 |

### 七、图片

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 正文内联图片 | 已支持 | 已支持 | 当前支持 PNG / JPEG。 |
| 表格单元格内联图片 | 已支持 | 已支持 | C++ 有专门 helper。 |
| 图片尺寸设置 | 已支持 | 已支持 | 以 pt 指定，缺一边时保留纵横比。 |
| file-like 图片输入 | 已支持 | 部分支持 | C++ 通过字节缓冲接口 `add_picture_data*` 提供等价能力。 |
| 浮动图片 / 浮动形状 | 未支持 | 未支持 | 两边都不是常规高层 API。 |
| 读取图片元数据 | 部分支持 | 已支持 | C++ 当前能读尺寸、关系 id、包内路径、所在位置。 |
| 图片数量统计 | 部分支持 | 已支持 | `Document::image_count()`。 |

### 八、section 与页面布局

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 页面尺寸 | 已支持 | 已支持 | 已支持读写。 |
| 页面方向 | 已支持 | 已支持 | Portrait / Landscape。 |
| 页边距 | 已支持 | 已支持 | top / right / bottom / left。 |
| 页眉页脚距离 | 已支持 | 已支持 | 已并入 `PageMargins`。 |
| gutter | 已支持 | 已支持 | 已并入 `PageMargins`。 |
| sections 集合 | 已支持 | 缺失 | C++ 当前仍是隐式单 section。 |
| 新增 section / section break | 已支持 | 缺失 | 当前没有多 section API。 |
| 页眉页脚内容 | 已支持 | 缺失 | 还没有 header/footer part 和引用模型。 |

### 九、样式与文档级系统

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| styles 集合 | 已支持 | 缺失 | C++ 没有 style catalog。 |
| 新增 / 删除自定义样式 | 已支持 | 缺失 | 当前未实现。 |
| 按名称应用段落样式 | 已支持 | 部分支持 | 目前只有标题样式捷径。 |
| 按名称应用表格样式 | 已支持 | 缺失 | 当前只能固定输出 `TableGrid`。 |
| settings 对象 | 已支持 | 缺失 | C++ 没有 settings API。 |
| 更多文档 part 访问能力 | 已支持 | 缺失 | 当前覆盖范围是有意收敛的。 |

## 读取侧的真实现状

当前 `docxcpp` 的读取能力已经不是空白，它已经能读回这些内容：

- `Document::paragraphs()`
- `Paragraph::text()`
- `Paragraph::alignment()`
- `Paragraph::first_run_style()`
- `Paragraph::runs()`
- `Paragraph::style_id()`
- `Paragraph::heading_level()`
- `Paragraph::format()`
- `Paragraph::has_page_break()`
- `Paragraph::hyperlinks()`
- `Document::tables()`
- `TableRow::cells()`
- `TableCell::text()`
- `TableCell::grid_span()`
- `TableCell::vertical_merge()`
- `Document::page_size_pt()`
- `Document::page_orientation()`
- `Document::page_margins_pt()`
- `Document::comments()`
- `Document::comment_count()`
- `Document::picture_sizes_pt()`
- `Document::pictures()`
- `Document::image_count()`

但和 `python-docx` 相比，缺口同样很清楚：

- 没有 sections 集合
- 没有 header / footer 读取模型
- 没有 styles 目录
- 没有可编辑的 paragraph / run 对象图
- 没有更广泛的 Word part 覆盖

## 当前最大的差距

和 `python-docx` 相比，现阶段最关键的差距主要在下面几个方向：

1. 编辑模型
`python-docx` 是围绕可变的 paragraph / run / table 对象建立的。`docxcpp` 仍然主要依赖 `Document` 级 helper 和轻量读取结果。

2. 样式系统
`python-docx` 具备真实的样式系统。`docxcpp` 目前更多还是直接格式化，加上有限的标题样式捷径。

3. section 与页眉页脚
`python-docx` 把 section 当作一等对象来处理。`docxcpp` 现在只支持单 section 的页面设置。

4. 超链接与批注深度
当前 C++ 版已经有基础能力，但还没有 `python-docx` 那种更完整的对象模型、目标地址读取和更细粒度锚点。

5. 表格完整度
当前 C++ 表格能力已经足够生成不少文档，但离 `python-docx` 的完整表格 DOM 仍然有明显距离。

## 建议的下一步

如果目标是继续向 `python-docx` 靠近，优先级比较高的方向是：

1. 继续深化 paragraph / run 可写能力
- 在 `Paragraph::add_run()` 之上继续扩展
- 支持向已有段落追加更丰富的内容类型
- 增加更完整的 run / font 属性

2. 深化超链接与批注
- 在 `HyperlinkInfo` 之上增加更完整的超链接对象模型
- 支持整段之外的批注锚点范围
- 增加更丰富的批注编辑元数据

3. 扩展段落格式
- tab stops
- 更丰富的行距模式
- widow/orphan control 以及更多段落标志

4. 扩展 section 与 header/footer
- 多 section
- section break
- header/footer 引用与内容

5. 加强表格 API
- 动态加行 / 列
- 嵌套表格
- 更好的合并单元格 / 省略单元格辅助

6. 建立样式系统
- styles 集合读取
- 按名称应用段落 / 表格 / 字符样式
- 自定义样式

## 一句话总结

`python-docx` 依然是更完整、更偏编辑型的库。

`docxcpp` 现在已经适合“受控模板 + 结构化生成 + 基础读取”这类场景：文本、混合格式、表格、图片、页面设置、基础超链接和基础批注都已经具备。接下来真正要补的是“深度”，也就是更完整的对象模型、更完整的读取模型、styles、sections，以及更广的 Word 特性覆盖。
