# 数据分析稿 Word 自动排版

用 Node.js + `docx` 库把结构化数据渲染成可直接交付的 Word 分析稿：标题层级、表格、图表、页眉页脚、页码、分节全部程序化控制。

## 组成

```
generate_report.js    主生成脚本（docx 文档构建）
package.json          依赖：docx、image-size
parsed.json           输入数据（解析后的分析指标）
```

## 设计要点

- **样式集中定义**：调色板、字号、表格线型、页面边距都在脚本顶部以常量集中管理，便于整套报告统一改版
- **表格与图表混排**：表格用 `Table`/`TableRow`/`TableCell` 构建，图片通过 `ImageRun` 按实际尺寸插入，避免拉伸变形
- **页眉页脚与页码**：用 `Header`/`Footer` + `PageNumber` 域代码实现自动页码，无需手工维护

## 快速开始

```bash
npm install
node generate_report.js
```

输出 Word 文档路径在脚本内 `OUTPUT` 常量中指定。

## 脱敏说明

本仓库已移除：客户分析稿（.docx / .pdf）、图表产物、本机绝对路径。

## 许可

MIT License，见 [LICENSE](LICENSE)。
