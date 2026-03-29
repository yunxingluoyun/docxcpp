# docxcpp vs python-docx

## 目的

这份文档描述 `docxcpp` 当前与 `python-docx` 的能力对齐情况，重点回答三个问题：

1. `docxcpp` 现在已经能做什么。
2. 它和 `python-docx` 相比还差在哪些层面。
3. 后续继续完善时，优先级最高的方向是什么。

本文不追求穷尽 WordprocessingML 的所有细节，而是以 `docxcpp` 当前公开 API、读写行为以及现有测试覆盖为准。

## 判断口径

| 状态 | 含义 |
| --- | --- |
| 已支持 | `docxcpp` 已实现，且当前有读写或回归测试覆盖 |
| 部分支持 | 已有能力，但范围、接口形态或对象模型明显弱于 `python-docx` |
| 未支持 | 当前没有公开能力，或仅有底层 XML 痕迹但没有可用 API |

## 总览

| 维度 | python-docx | docxcpp |
| --- | --- | --- |
| 文档创建 / 打开 / 保存 | 完整 | 已支持 |
| 段落 / Run 基础编辑 | 完整 | 已支持 |
| 段落格式 | 完整 | 已支持 |
| 字符格式 | 完整 | 已支持 |
| 字符样式 | 完整样式对象模型 | 部分支持，当前支持按样式 ID / 名称引用 |
| 表格基础编辑 | 完整 | 已支持 |
| 表格样式 | 完整样式引用与对象访问 | 部分支持，当前支持按样式 ID / 名称引用 |
| 图片 | 完整 | 已支持 |
| 超链接 | 完整 | 已支持 |
| 批注 | 完整 | 部分支持 |
| 页面设置 / Section | 完整 | 已支持 |
| Header / Footer | 完整对象模型 | 部分支持，读写能力已较完整，但对象模型较轻 |
| Styles 集合 / 自定义样式系统 | 完整 | 未支持 |

## 已经对齐得比较好的部分

### 文档生命周期

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 新建文档 | ✅ | ✅ |
| 打开已有文档 | ✅ | ✅ |
| 保存文档 | ✅ | ✅ |
| 读写同一个文档后再读取 | ✅ | ✅ |

### 段落与 Run

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 新增普通段落 | ✅ | ✅ `add_paragraph()` |
| 新增标题 | ✅ | ✅ `add_heading()` |
| 新增带字符格式的段落 | ✅ | ✅ `add_styled_paragraph()` |
| 一段包含多个 Run | ✅ | ✅ |
| 向已有段落追加 Run | ✅ | ✅ `Paragraph::add_run()` |
| 在已有段落前插入段落 | ✅ | ✅ `Paragraph::insert_paragraph_before()` |
| 读取段落列表 | ✅ | ✅ `paragraphs()` |
| 读取 Run 列表 | ✅ | ✅ `runs()` |
| 读取段落样式 ID | ✅ | ✅ `style_id()` |
| 读取标题级别 | ✅ | ✅ `heading_level()` |

### 字符格式

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 粗体 / 斜体 / 下划线 | ✅ | ✅ |
| 删除线 / 双删除线 | ✅ | ✅ |
| 全大写 / 小型大写 | ✅ | ✅ |
| 上标 / 下标 | ✅ | ✅ |
| 字号 | ✅ | ✅ |
| 字体名 | ✅ | ✅ |
| `ascii` / `hAnsi` / `eastAsia` / `cs` 字体槽位 | ✅ | ✅ |
| 颜色 | ✅ | ✅ |
| 高亮 | ✅ | ✅ |
| Run 内换行 | ✅ | ✅ |

### 段落格式

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 对齐 | ✅ | ✅ |
| 左缩进 / 右缩进 / 首行缩进 | ✅ | ✅ |
| 段前 / 段后间距 | ✅ | ✅ |
| 固定行距 | ✅ | ✅ |
| 多倍行距 / AtLeast | ✅ | ✅ |
| Tab stops | ✅ | ✅ |
| keep_together | ✅ | ✅ |
| keep_with_next | ✅ | ✅ |
| page_break_before | ✅ | ✅ |
| widow / orphan control | ✅ | ✅ |

### 表格基础能力

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 创建表格 | ✅ | ✅ `add_table()` |
| 添加行 | ✅ | ✅ `add_table_row()` |
| 添加列 | ✅ | ✅ `add_table_column()` |
| 写单元格文本 | ✅ | ✅ `set_table_cell()` |
| 单元格写多 Run | ✅ | ✅ |
| 单元格追加段落 | ✅ | ✅ `add_table_cell_paragraph()` |
| 合并单元格 | ✅ | ✅ `merge_table_cells()` |
| 嵌套表格 | ✅ | ✅ `add_nested_table()` |
| 读取表格列表 | ✅ | ✅ `tables()` |
| 读取嵌套表格 | ✅ | ✅ |
| 读取横向 / 纵向合并结果 | ✅ | ✅ |

### 图片、超链接、分页与批注基础能力

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 正文插图 | ✅ | ✅ `add_picture()` / `add_picture_data()` |
| 表格单元格插图 | ✅ | ✅ |
| 指定图片尺寸 | ✅ | ✅ |
| 读取图片信息 | 部分支持 | ✅ `pictures()` / `picture_sizes_pt()` |
| 正文超链接 | ✅ | ✅ `add_hyperlink()` |
| 表格单元格超链接 | ✅ | ✅ `add_hyperlink_to_table_cell()` |
| 读取超链接文本 / URL | ✅ | ✅ |
| 插入分页符 | ✅ | ✅ `add_page_break()` |
| 添加批注 | ✅ | ✅ `add_comment()` |
| 读取批注 | ✅ | ✅ `comments()` |

### 页面设置、Section、页眉页脚

| 功能 | python-docx | docxcpp |
| --- | --- | --- |
| 页面尺寸 | ✅ | ✅ |
| 页面方向 | ✅ | ✅ |
| 页边距 | ✅ | ✅ |
| Header / Footer 距离 | ✅ | ✅ |
| Gutter | ✅ | ✅ |
| 多单位页面设置 | ✅ | ✅ |
| 标准纸张快捷值 | 不以内置枚举提供 | ✅ `StandardPageSize` |
| 新增 section break | ✅ | ✅ `add_section_break()` |
| 读取 sections | ✅ | ✅ `sections()` |
| 写 Header / Footer 文本 | ✅ | ✅ |
| 写 Header / Footer 段落 | ✅ | ✅ |
| 读取 Header / Footer 文本 | ✅ | ✅ |
| 读取 Header / Footer 段落 | ✅ | ✅ |
| 首页不同 | ✅ | ✅ |
| 奇偶页不同 | ✅ | ✅ |

## 目前属于“部分支持”的部分

这些能力已经不是空白，但和 `python-docx` 相比还停留在较薄的一层。

### 1. 字符样式

`python-docx` 可以通过样式对象体系处理字符样式。  
`docxcpp` 当前已经支持：

- 在 `RunStyle` 上设置 `character_style_id`
- 在 `RunStyle` 上设置 `character_style_name`
- 写出 `w:rStyle`
- 读取已有文档中的 `w:rStyle`
- 从 `styles.xml` 解析样式 ID 与样式名称的对应关系

但当前仍然缺少：

- `styles()` 之类的样式集合访问入口
- 独立的字符样式对象
- 创建、修改、自定义字符样式的 API
- 基于继承链的完整样式解析

结论：字符样式已具备“引用既有样式”的能力，但还没有形成完整样式系统。

### 2. 表格样式

`docxcpp` 当前已经支持：

- `add_table(..., style_id_or_name)`
- `add_nested_table(..., style_id_or_name)`
- `set_table_style(table_index, style_id_or_name)`
- 读取 `w:tblStyle`
- 读取 `Table::style_id()` 和 `Table::style_name()`

但当前仍然缺少：

- 表格样式对象模型
- 样式集合遍历与查询
- 创建 / 修改自定义表格样式
- 对表格样式内部条件格式规则的直接建模

结论：表格样式已经能“按名称或 ID 套用并读回”，但还不是完整的样式层。

### 3. 批注锚点

`docxcpp` 现在可以添加和读取批注，但写入入口仍偏简化。  
与 `python-docx` 相比，当前缺的主要是：

- 将批注精确锚定到任意 Run 范围
- 更完整的批注对象编辑能力

结论：批注已经够做基础读写，但锚点精度还不够。

### 4. Header / Footer 与 Section 对象模型

`docxcpp` 的读写能力已经覆盖了很多常见场景，但它仍然更偏“文档操作 API”，而不是 `python-docx` 那种更完整的对象树。

当前特点：

- 能按 section 写文本、段落、样式化段落
- 能读回 header / footer 文本和段落
- 能控制 default / first / even 三类页眉页脚

当前不足：

- Header / Footer 缺少更完整的可变对象模型
- Section 内更多细粒度元素还没有统一对象化

结论：能力够用，但抽象层还偏轻。

## 当前明显落后于 python-docx 的部分

### 样式系统

这是 `docxcpp` 当前和 `python-docx` 差距最大的模块。

`python-docx` 侧重的是：

- `styles` 集合
- 样式对象访问
- 按样式类型检索
- 更自然地处理段落样式、字符样式、表格样式

`docxcpp` 当前还没有这些统一入口。  
现在的样式能力主要表现为：

- 标题样式快捷方式
- 段落 `style_id()` 读取
- 字符样式引用
- 表格样式引用

换句话说，`docxcpp` 已经开始“认识样式”，但还没有“管理样式”。

### 自定义样式与样式编辑

当前 `docxcpp` 还不支持：

- 创建新样式
- 修改 `styles.xml` 中已有样式定义
- 暴露样式继承关系
- 暴露样式默认值解析

这意味着在复杂模板或企业文档体系里，`docxcpp` 目前更适合作为“消费现有样式”的库，而不是“维护整套样式系统”的库。

## docxcpp 当前更有特色的点

虽然总体广度仍不如 `python-docx`，但 `docxcpp` 现在也有几处比较鲜明的特点：

- 同时提供读和写，不只是生成文档。
- 针对 `python-docx` 做了双向互通验证，不只验证自己能写，还验证能读 `python-docx` 生成的文档。
- 页面设置已经统一到长度模型，支持 `mm`、`cm`、`inches`、`pt`、`twips`。
- 提供了标准纸张快捷值，做页面设置时更直接。
- 对表格样式和字符样式，已经能按“样式名称”引用，而不强迫调用方只处理内部 ID。

## 适用场景判断

### 更适合用 docxcpp 的场景

- C++ 项目内直接生成或读取 `.docx`
- 模板字段填充、报表生成、结构化导出
- 对段落、表格、图片、分页、页眉页脚有较完整但不过度复杂的需求
- 已有模板样式基本稳定，只需要引用既有样式

### 更适合用 python-docx 的场景

- 需要完整样式系统
- 需要更成熟、更宽的对象模型
- 需要围绕样式、Section、Header/Footer 做更细的程序化控制
- 需要更成熟的 Python 自动化生态

## 后续优先级建议

如果继续按 `python-docx` 对齐，我建议优先顺序如下：

1. 完整样式系统入口：`styles()`、样式查询、段落/字符/表格样式对象。
2. 段落样式和表格样式的显式设置与读取统一化。
3. 批注锚点升级到任意 Run 范围。
4. Header / Footer / Section 更完整的对象模型。
5. 将更多长度相关 API 从旧的 `*_pt` 风格继续统一到 `Length`。

## 一句话结论

`docxcpp` 现在已经覆盖了 `python-docx` 的一大块核心实用能力，尤其是段落、字符格式、表格、页面设置、页眉页脚、图片、超链接和基础批注；但它和 `python-docx` 的主要差距，已经从“能不能做”收缩到了“样式系统是否完整、对象模型是否足够深”。
