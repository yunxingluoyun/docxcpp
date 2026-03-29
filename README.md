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
| 页面设置 | 页面尺寸、方向、页边距、Header/Footer 距离、Gutter，支持 `pt/mm/cm/inches/twips` |
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

## 构建依赖

构建 `docxcpp` 本身需要：

- C++17 编译器
- CMake 3.20+
- Python 3

其中：

- `pugixml` 和 `miniz` 已内置在仓库 `3rdparty/` 目录中，不需要额外安装
- Python 3 主要用于构建时生成内嵌默认模板，以及运行部分测试
- 如果只是在你自己的项目里链接已安装好的 `docxcpp`，则不需要额外再关心 `pugixml` 和 `miniz`

## 库构建

### 默认构建

默认情况下会构建静态库：

```bash
cmake -S . -B build
cmake --build build
```

### 构建静态库

```bash
cmake -S . -B build-static -DBUILD_SHARED_LIBS=OFF
cmake --build build-static
```

### 构建动态库

`docxcpp` 支持跨平台构建共享库：

```bash
cmake -S . -B build-shared -DBUILD_SHARED_LIBS=ON
cmake --build build-shared
```

### 安装库

静态库和动态库都可以通过 `cmake --install` 安装：

```bash
cmake --install build-static --prefix /your/install/prefix
cmake --install build-shared --prefix /your/install/prefix
```

安装后会导出：

- 头文件
- `docxcpp` 库文件
- `docxcppConfig.cmake`
- `docxcppTargets.cmake`

默认模板已内嵌到库中，所以安装后的 `.dll` / `.so` / `.dylib` 不依赖源码目录。

## 外部项目使用

安装后可通过 CMake 包导入：

```cmake
find_package(docxcpp CONFIG REQUIRED)
target_link_libraries(your_target PRIVATE docxcpp::docxcpp)
```

如果库安装在非系统默认路径，可以在配置你的项目时传入：

```bash
cmake -S . -B build -DCMAKE_PREFIX_PATH=/your/install/prefix
```

## 测试

完整测试：

```bash
ctest --test-dir build --output-on-failure
```

其中 `ctest` 包含：

- 常规读写回归
- `python-docx` 双向互通回归
- 安装烟雾测试：把当前构建安装到临时前缀，再用独立外部 CMake 工程通过 `find_package(docxcpp)` 链接并运行

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

**页面单位示例**

```cpp
document.set_page_size(docxcpp::StandardPageSize::A4);

document.set_page_margins(docxcpp::PageMargins{
    docxcpp::Inches(1.0),
    docxcpp::Cm(2.0),
    docxcpp::Inches(1.0),
    docxcpp::Cm(2.0),
    docxcpp::Pt(36),
    docxcpp::Pt(24),
    docxcpp::Twips(360),
});
```

也可以直接指定常见标准纸型：

```cpp
document.set_page_size(docxcpp::StandardPageSize::A0);
document.set_page_size(docxcpp::StandardPageSize::B5);
document.set_page_size(docxcpp::StandardPageSize::Letter);
document.add_section_break(docxcpp::Section(docxcpp::StandardPageSize::Legal));
```

**字体样式示例**

```cpp
docxcpp::RunStyle style;
style.font_name = "Arial";
style.character_style_name = "Emphasis";
style.east_asia_font_name = "Microsoft YaHei";
style.complex_script_font_name = "Amiri";
style.all_caps = true;
style.strike = true;
style.superscript = true;
```

**样式引用示例**

```cpp
document.add_table(3, 4, "Light Shading Accent 1");
document.set_table_style(0, "Medium Grid 1 Accent 1");

docxcpp::RunStyle run_style;
run_style.character_style_id = "Strong";
document.add_styled_paragraph("Styled by character style", run_style);
```

## 与 python-docx 对比

详见 [docxcpp_vs_python_docx.md](docxcpp_vs_python_docx.md)。
