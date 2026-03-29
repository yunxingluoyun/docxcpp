# docxcpp vs python-docx 功能对比

## 概述

**docxcpp** 是 C++ 的 .docx 处理库，基于 pugixml 和 miniz 实现，支持读写。

**python-docx** 是 Python 生态中最成熟的 .docx 库。

本文按功能模块对比两者能力。

---

## 状态说明

| 状态 | 含义 |
| --- | --- |
| 已支持 | 已实现且有测试覆盖 |
| 部分支持 | 已有实现但范围比 python-docx 窄 |
| 缺失 | 尚未实现 |

---

## 文档生命周期

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 打开文档 | ✅ | ✅ |
| 创建文档 | ✅ | ✅ |
| 保存文档 | ✅ | ✅ |

---

## 段落与 Run

### 写入

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 新增段落 | ✅ | ✅ `add_paragraph()` |
| 新增标题 | ✅ | ✅ `add_heading()` |
| 插入分页符 | ✅ | ✅ `add_page_break()` |
| 带样式段落 | ✅ | ✅ `add_styled_paragraph()` |
| 一段中多个 Run | ✅ | ✅ |
| 向已有段落追加 Run | ✅ | ✅ `Paragraph::add_run()` |
| 在已有段落前插入段落 | ✅ | ❌ |

### 读取

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 读取段落列表 | ✅ | ✅ `paragraphs()` |
| 读取 Runs | ✅ | ✅ `runs()` |
| 读取样式 ID | ✅ | ✅ `style_id()` |
| 读取标题级别 | ✅ | ✅ `heading_level()` |

---

## Run 字符格式

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 粗体 | ✅ | ✅ |
| 斜体 | ✅ | ✅ |
| 下划线 | ✅ | ✅ |
| 字号 | ✅ | ✅ |
| 字体名 | ✅ | ✅ |
| 颜色 | ✅ | ✅ |
| 高亮 | ✅ | ✅ |
| 行内换行 | ✅ | ✅ |
| Tab stops | ✅ | ✅ `set_paragraph_tab_stops()` |
| 字符样式 | ✅ | ❌ |

---

## 段落格式

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 对齐 | ✅ | ✅ |
| 左右缩进、首行缩进 | ✅ | ✅ |
| 段前段后间距 | ✅ | ✅ |
| 固定行距 | ✅ | ✅ `set_paragraph_line_spacing_pt()` |
| 多倍行距 | ✅ | ✅ `LineSpacingMode::Multiple` |
| Tab stops | ✅ | ✅ |
| keep_together | ✅ | ✅ |
| keep_with_next | ✅ | ✅ |
| page_break_before | ✅ | ✅ |
| widow/orphan control | ✅ | ❌ |

---

## 表格

### 写入

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 创建表格 | ✅ | ✅ `add_table()` |
| 加行 | ✅ | ✅ `add_table_row()` |
| 加列 | ✅ | ✅ `add_table_column()` |
| 嵌套表格 | ✅ | ✅ `add_nested_table()` |
| 写单元格文本 | ✅ | ✅ `set_table_cell()` |
| 单元格追加段落 | ✅ | ✅ `add_table_cell_paragraph()` |
| 合并单元格 | ✅ | ✅ `merge_table_cells()` |

### 读取

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 读取表格列表 | ✅ | ✅ `tables()` |
| 读取单元格文本 | ✅ | ✅ |
| 读取嵌套表格 | ✅ | ✅ |
| 读取合并状态 | ✅ | ✅ |

---

## 图片

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 正文插入图片 | ✅ | ✅ `add_picture()` |
| 单元格插入图片 | ✅ | ✅ `add_picture_to_table_cell()` |
| 字节缓冲输入 | ✅ | ✅ `add_picture_data()` |
| 指定尺寸 | ✅ | ✅ |
| 读取图片元数据 | 部分支持 | ✅ `pictures()` / `picture_sizes_pt()` |

---

## 超链接

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 正文添加超链接 | ✅ | ✅ `add_hyperlink()` |
| 单元格添加超链接 | ✅ | ✅ `add_hyperlink_to_table_cell()` |
| 读取超链接文本 | ✅ | ✅ |
| 读取超链接 URL | ✅ | ✅ |

---

## 批注

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 添加批注 | ✅ | ✅ `add_comment()` |
| 读取批注元数据 | ✅ | ✅ `comments()` |
| 任意 Run 范围批注锚 | ✅ | ❌ |

---

## 页面与 Section

### 页面设置

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 页面尺寸 | ✅ | ✅ |
| 页面方向 | ✅ | ✅ |
| 页边距 | ✅ | ✅ |
| Header/footer 距离 | ✅ | ✅ |
| Gutter | ✅ | ✅ |

### Section 写入

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 新增 section break | ✅ | ✅ `add_section_break()` |
| 多 section 写入 | ✅ | 部分支持（可添加 break，细节有限） |

### Section 读取

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 读取 sections | ✅ | ✅ `sections()` |
| 多 section 读取 | ✅ | ✅ |

### Header/Footer 写入

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| Header 文本 | ✅ | ✅ `set_header_text()` |
| Footer 文本 | ✅ | ✅ `set_footer_text()` |
| Header 段落 | ✅ | ✅ `add_header_paragraph()` |
| Footer 段落 | ✅ | ✅ `add_footer_paragraph()` |
| 奇偶页不同 | ✅ | ✅ `set_even_and_odd_headers()` |
| 首页不同 | ✅ | ✅ `set_section_different_first_page()` |

### Header/Footer 读取

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 读取 Header 文本 | ✅ | ✅ `headers()` |
| 读取 Footer 文本 | ✅ | ✅ `footers()` |
| 读取 Header 段落 | ✅ | ✅ `header_paragraphs()` |
| 读取 Footer 段落 | ✅ | ✅ `footer_paragraphs()` |

---

## 样式系统

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| Styles 集合 | ✅ | ❌ |
| 按名称应用样式 | ✅ | 部分支持（主要是 heading 快捷方式） |
| 字符样式对象 | ✅ | ❌ |
| 表格样式 | ✅ | ❌ |
| 自定义样式 | ✅ | ❌ |

---

## 总结

| 维度 | python-docx | docxcpp |
| --- | --- | --- |
| 覆盖广度 | 更完整 | 覆盖核心场景 |
| 样式系统 | 完整 | 基本缺失 |
| Header/Footer | 完整对象模型 | 仅文本段落支持 |
| Section 写入 | 完整 | 基础支持 |
| 表格嵌套/动态操作 | 完整 | 支持创建，动态操作有限 |
| 超链接/批注对象模型 | 完整 | 轻量快照 |

docxcpp 已具备实际文档生成和读取能力，适合结构化文档生成与模板填充场景。python-docx 在样式、Header/Footer 对象模型、表格动态操作上更为完整。
