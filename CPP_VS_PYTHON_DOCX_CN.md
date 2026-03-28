# docxcpp 与 python-docx 对比

这份文档描述的是当前仓库里 `cpp/` 下的 `docxcpp`，与上游 `python-docx` 相比已经做到什么、还缺什么。

## 一句话结论

`docxcpp` 已经从“只能写几段文本”的阶段，进展到“可以生成并读取一批真实 Word 文档场景”的阶段。

它现在已经能覆盖：

- 段落、标题、run
- 段落格式
- 表格与基础合并
- 图片
- 超链接
- 基础批注
- 页面设置
- 一部分读取侧 API

但它仍然不是 `python-docx` 的等价替代。`python-docx` 的对象模型、样式系统、section/header-footer、表格能力，依然更完整。

## 状态说明

- `已支持`：当前已经实现，并且在现有测试里有覆盖
- `部分支持`：已有能力，但范围明显比 `python-docx` 小
- `缺失`：还没有公开 API 或还没实现

## 当前已实现的重点

`docxcpp` 目前已经支持：

- 打开、创建、保存 `.docx`
- 新增段落、标题、分页符
- 多 run 段落和多 run 标题
- `Paragraph::add_run()`
- run 基础格式：粗体、斜体、下划线、字号、字体、颜色、高亮
- 段落格式：对齐、缩进、段前段后、固定行距、基础分页标志
- 表格：创建、写单元格、单元格追加段落、基础合并
- 图片：文件路径输入、字节缓冲输入、正文图片、单元格图片
- 超链接：正文和表格单元格
- 批注：基础段落级插入与读取
- 页面设置：纸张、方向、页边距、header/footer 距离、gutter
- 读取段落、runs、style id、heading level、段落格式、表格、图片、批注
- `sections()` 读取，多 section 读取已支持

## 功能对比

### 文档生命周期

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 打开文档 | 已支持 | 已支持 | 两边都支持 `.docx` 打开 |
| 创建文档 | 已支持 | 已支持 | C++ 使用内置模板 |
| 保存文档 | 已支持 | 已支持 | 两边都能输出有效文档 |

### 段落与 run

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 新增段落 | 已支持 | 已支持 | `Document::add_paragraph()` |
| 新增标题 | 已支持 | 已支持 | 支持标题级别 |
| 插入分页符 | 已支持 | 已支持 | `Document::add_page_break()` |
| 一段中多个 run | 已支持 | 已支持 | 使用 `std::vector<Run>` |
| 向已有段落追加 run | 已支持 | 部分支持 | 当前是 `Paragraph::add_run()` |
| 在已有段落前插入段落 | 已支持 | 缺失 | 还没有对应 C++ API |
| 读取段落 | 已支持 | 已支持 | `Document::paragraphs()` |
| 读取 runs | 已支持 | 已支持 | `Paragraph::runs()` |

### run 与文本格式

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 粗体/斜体/下划线 | 已支持 | 已支持 | 基础格式已具备 |
| 字号 | 已支持 | 已支持 | pt |
| 字体名 | 已支持 | 已支持 | 基础字体名写入 |
| 字体颜色 | 已支持 | 已支持 | RGB hex |
| 高亮 | 已支持 | 已支持 | Word token |
| 行内换行 | 已支持 | 已支持 | 写入与读取都支持 |
| tab stops | 已支持 | 缺失 | 还没有 API |
| 字符样式 | 已支持 | 缺失 | 还没有 style system |

### 段落格式

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 对齐 | 已支持 | 已支持 | left/center/right/justify |
| 左右缩进、首行缩进 | 已支持 | 已支持 | 已支持读写 |
| 段前段后间距 | 已支持 | 已支持 | 已支持读写 |
| 行距 | 已支持 | 部分支持 | 当前是固定行距 |
| keep_together / keep_with_next / page_break_before | 已支持 | 已支持 | 已具备 |
| widow/orphan control | 已支持 | 缺失 | 未实现 |

### 表格

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 创建表格 | 已支持 | 已支持 | `Document::add_table()` |
| 读取表格 | 已支持 | 已支持 | `Document::tables()` |
| 写单元格文本 | 已支持 | 已支持 | 文本和多 run 都支持 |
| 单元格内追加段落 | 已支持 | 已支持 | 已支持 |
| 合并单元格 | 已支持 | 部分支持 | 已支持基础合并与读取 |
| 动态加行/列 | 已支持 | 缺失 | 还没有 API |
| 嵌套表格 | 已支持 | 缺失 | 还没有 API |

### 图片

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 正文图片 | 已支持 | 已支持 | PNG/JPEG |
| 单元格图片 | 已支持 | 已支持 | 已支持 |
| 图片尺寸 | 已支持 | 已支持 | pt |
| file-like 输入 | 已支持 | 部分支持 | C++ 用字节缓冲等价实现 |
| 读取图片元数据 | 部分支持 | 已支持 | C++ 侧读取更直接 |

### 超链接与批注

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 写超链接 | 已支持 | 已支持 | 正文和表格单元格都支持 |
| 读超链接目标地址 | 已支持 | 已支持 | `Paragraph::hyperlinks()` |
| 超链接对象模型 | 已支持 | 部分支持 | 当前是轻量快照，不是完整对象 |
| 写批注 | 已支持 | 部分支持 | 当前是基础段落级批注 |
| 读批注 | 已支持 | 部分支持 | 可读 comments part 元数据 |
| 任意 run 范围批注 | 已支持 | 缺失 | 还未实现 |

### 页面与 section

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| 页面尺寸 | 已支持 | 已支持 | 已支持读写 |
| 页面方向 | 已支持 | 已支持 | portrait/landscape |
| 页边距 | 已支持 | 已支持 | 已支持读写 |
| header/footer 距离 | 已支持 | 已支持 | 已支持 |
| gutter | 已支持 | 已支持 | 已支持 |
| `sections()` 集合读取 | 已支持 | 部分支持 | 已支持多 section 读取 |
| 新增 section / section break | 已支持 | 缺失 | 还没有写入 API |
| header/footer 内容 | 已支持 | 缺失 | 还没有对象模型 |

### 样式系统

| 功能 | python-docx | docxcpp | 说明 |
| --- | --- | --- | --- |
| styles 集合 | 已支持 | 缺失 | 还没有 |
| 按名称应用段落样式 | 已支持 | 部分支持 | 当前主要是标题捷径 |
| 字符样式 / 表格样式 | 已支持 | 缺失 | 未实现 |
| 自定义样式 | 已支持 | 缺失 | 未实现 |

## 读取侧现状

当前读取侧已经不是空壳，已经可以读：

- 段落文本
- 段落对齐
- runs
- style id
- heading level
- 段落格式
- 页面设置
- `sections()`
- 表格与单元格
- 图片元数据
- comments 元数据
- hyperlinks 文本和 URL

但离 `python-docx` 还有这些距离：

- 没有完整 styles 体系
- 没有 header/footer 读取模型
- 没有更完整的 paragraph/run 可编辑对象图
- 没有更广的 Word parts 覆盖

## 当前最值得继续补的方向

如果目标是继续向 `python-docx` 靠近，建议优先做：

1. 扩展 paragraph / run 写入模型
2. 深化超链接与批注对象模型
3. 扩展段落格式，特别是 tab stops 和更完整行距
4. 增加 section break、多 section 写入、header/footer
5. 扩展表格 API
6. 建立 styles 系统

## 总结

`python-docx` 仍然是更完整、覆盖更深的实现。

`docxcpp` 当前已经具备了“能实际生成和读取一批真实 Word 文档”的能力，尤其适合结构化文档生成、模板填充和自动化产出。下一阶段最关键的不是“再加几个 helper”，而是继续补齐对象模型、样式系统、section/header-footer 和更强的编辑能力。
