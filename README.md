# docxcpp

C++ `.docx` 读写库，基于 [pugixml](https://pugixml.org/) 和 [miniz](https://github.com/richgel999/miniz) 实现。

## 能力概览

| 分类 | 支持内容 |
| --- | --- |
| 文档 | 打开 / 创建 / 保存 `.docx` |
| 段落 | 新增段落、标题、分页符、样式段落、多 Run 段落 |
| Run | 粗体、斜体、下划线、字号、字体、颜色、高亮 |
| 段落格式 | 对齐、缩进、段前段后、行距（固定/多倍/至少）、Tab Stops、分页控制 |
| 表格 | 创建、加行/列、嵌套表格、单元格合并 |
| 图片 | 正文/单元格插入，支持文件路径和字节缓冲 |
| 超链接 | 正文和表格单元格中的外部超链接 |
| 批注 | 段落级批注的写入与读取 |
| 页面设置 | 页面尺寸、方向、页边距、Header/Footer 距离、Gutter |
| Section | Section Break、多 Section 读取、Header/Footer 内容（文本/段落） |
| 读取 | 段落、Runs、表格、图片元数据、批注元数据、超链接、Section 配置 |

## 项目结构

```
include/docxcpp/    头文件
src/                实现
examples/           示例
tests/              测试（含 python-docx 回归验证）
3rdparty/           pugixml、miniz
```

## 快速开始

```bash
cmake -S . -B build
cmake --build build
ctest --test-dir build --output-on-failure
```

## 示例

**最小示例** — [examples/minimal.cpp](examples/minimal.cpp)

```cpp
docxcpp::Document document;
document.add_paragraph("Hello, World!");
document.add_heading("A Heading", 1);
document.save("output.docx");
```

**完整示例** — [examples/full_feature_5_pages.cpp](examples/full_feature_5_pages.cpp)

覆盖标题、混合 Run、段落格式、表格（含嵌套与合并）、图片、超链接、批注、Section Break、Header/Footer。

## 测试

所有测试均通过 `ctest` 运行，其中 `python_docx_regression` 使用 python-docx 读取 docxcpp 生成的文档并验证结构正确性。

```bash
ctest --test-dir build --output-on-failure
```

## 与 python-docx 对比

详见 [docxcpp_vs_python_docx.md](docxcpp_vs_python_docx.md)。
