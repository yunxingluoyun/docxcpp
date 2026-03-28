#  docxcpp

本仓库内容：

- 正在持续演进的 C++ 版本 `docxcpp`

`docxcpp` 的目标不是简单做一个最小导出器，而是逐步向 `python-docx` 的文档对象模型靠近，同时保持一个可用、可测试、可扩展的 C++ `.docx` 读写库。

## `docxcpp` 

当前 C++ 实现依赖：

- `pugixml`
- `miniz`

源码位置：

- 头文件：[cpp/include/docxcpp](./include/docxcpp)
- 实现：[cpp/src](./src)
- 示例：[cpp/examples](./examples)
- 测试：[cpp/tests](./tests)

## `docxcpp` 当前能力

已经支持的核心能力：

- 打开、创建、保存 `.docx`
- 新增段落、标题、分页符
- 写多 run 混合段落和混合标题
- 通过 `Paragraph::add_run()` 向已有段落继续追加 run
- run 基础格式：粗体、斜体、下划线、字号、字体、颜色、高亮
- 段落格式：对齐、缩进、段前段后、固定行距、基础分页标志
- 创建表格、写单元格文本、向单元格追加段落、合并单元格
- 插入图片：文件路径和内存字节缓冲两种入口
- 正文和表格单元格中的外部超链接
- 基础段落级批注
- 页面设置：纸张大小、方向、页边距、header/footer 距离、gutter
- 读取段落、runs、style id、heading level、段落格式
- 读取表格、图片元数据、批注元数据
- 读取 `sections()` 集合，支持多 section 读取

## 当前定位

`docxcpp` 现在已经适合这些场景：

- 程序化生成结构化 Word 文档
- 受控模板上的内容填充
- 基础读取与结构检查
- 用 C++ 批量产出带表格、图片、超链接、批注的 `.docx`

它还没有完全达到 `python-docx` 的覆盖深度，尤其是：

- 完整样式系统
- 更强的 paragraph/run 可编辑对象模型
- 多 section 写入和 section break
- header/footer 内容模型
- 更完整的表格对象能力

## 示例

最小示例：

- [minimal.cpp](./examples/minimal.cpp)

完整中文示例：

- [full_feature_5_pages.cpp](./examples/full_feature_5_pages.cpp)

这个完整示例会生成一个 5 页中文 `.docx`，覆盖当前已实现的大部分功能。

## 构建

```bash
cmake -S . -B build
cmake --build build
ctest --test-dir build --output-on-failure
```

## 运行 C++ 示例

运行最小示例：

```bash
./build/docxcpp_example
```

运行完整中文示例：

```bash
./build/docxcpp_full_feature_example build/test-artifacts/full-feature-5-pages-zh.docx
```

## Python 验证完整示例

仓库中提供了一个基于 `python-docx` 的验证脚本：

- [validate_full_feature_example.py](./tests/validate_full_feature_example.py)

运行方式：

```bash
python3 cpp/tests/validate_full_feature_example.py build/test-artifacts/full-feature-5-pages-zh.docx
```

它会校验：

- 页面设置
- 段落内容与段落格式
- 表格内容
- 图片数量
- 超链接关系
- comments part
- 分页结构

## 对比文档

英文版：

- [CPP_VS_PYTHON_DOCX.md](./CPP_VS_PYTHON_DOCX.md)

中文版：

- [CPP_VS_PYTHON_DOCX_CN.md](./CPP_VS_PYTHON_DOCX_CN.md)
