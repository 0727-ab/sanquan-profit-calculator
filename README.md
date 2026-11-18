# 三全利润与人工测算器

给市场分析员工用的单品利润 / 人工成本测算 + 损益数据复盘工具。一个 HTML 文件，双击用 Chrome 或 Edge 打开，不用安装，断网可用；可导出带公式的 Excel。

> 演示数据全部为假设。

## 直接使用

打开 `dist/三全利润与人工测算器.html`。使用方法见 [docs/使用说明.txt](docs/使用说明.txt)。

## 目录

| 路径 | 内容 |
|---|---|
| `src/engine.js` | 计算引擎：利润、补贴、阶梯激励、人工计算器、公式解析、自动分析 |
| `src/xlsx.js` | Excel 导出：纯 JS 生成带公式的 .xlsx（无第三方库） |
| `src/xlsxread.js` | Excel 读取：在浏览器本地解压并读取 .xlsx（不上传） |
| `src/review.js` | 损益复盘引擎：数据类型模板、字段检查、勾稽核对、当月 / 累计对比、子项下钻、结论规则、生成测算方案 |
| `src/reviewexport.js` | 复盘结果导出：Excel 数据表、XMind 括号图 |
| `src/defaults.js` | 演示数据和人工预设方案 A / B / C |
| `src/app.html` | 界面、图表（手写 SVG）、交互 |
| `build.py` | 把 `src/` 合成为 `dist/` 里的单个 HTML |
| `tests/test.js` | 验收案例（计算引擎） |
| `tests/test_review.js` | 复盘引擎测试（虚构数据） |
| `tests/test_xlsx.js` | 导出 Excel，用于和网页结果比对 |

## 修改后重新生成

```bash
python3 build.py          # 生成 dist/三全利润与人工测算器.html
node tests/test.js        # 跑验收案例，应显示 ALL PASS
```

验证 Excel 公式（需要 LibreOffice）：

```bash
node tests/test_xlsx.js nocache nc.xlsx          # 导出不带缓存值的 Excel
soffice --headless --calc --convert-to xlsx --outdir lo nc.xlsx   # 让 LibreOffice 重算
# 再用 openpyxl 读取 lo/nc.xlsx，与 expected.json 比对
```

## 设计要点

- **计算和界面分离**：网页、Excel 导出、测试共用同一个计算引擎。
- **Excel 保留公式**：.xlsx 本质是装着 XML 文本的压缩包；导出器按行号把每条规则拼成 Excel 公式（如 `B22*B5`），再打包下载。
- **员工自定义公式**：解析成语法树，同一棵树既能在网页里计算，也能翻译成 Excel 公式；只支持 + − × ÷、IF、MIN、MAX、ROUND，保证一一对应。

## 数据复盘：导入规则

- **数据类型模板**（`review.js` 的 `SOURCES`）：每种报表一条，写明需要哪几张表、表头特征。认不出来的文件直接拒绝。新增报表类型时在这里加一条。
- **必需字段**缺一个就不导入；**可选字段**缺了，对应分析隐藏，绝不当 0 计算。
- 导入后自动做勾稽核对（收入 − 各项成本费用 = 净利润）和同期月份检查。
- **真实经营数据不要提交到本仓库**（仓库公开）。测试只用虚构数据。

## 版本
- V2：新增“数据复盘”——导入系统损益明细（本期 / 同期），自动生成当月 / 累计的产品维度、费用维度分析和结论草稿；差异大的费用自动下钻到子项；导出 Excel / XMind；一键用实际数据生成测算方案。

- V1.1：人工“不计人工”选项；含 / 不含人工利润对比；瀑布图显示全部人工；各商品单独的人工分摊比例。
- V1：利润测算、人工成本计算器、SKU 汇总、自动分析、带公式的 Excel 导出。
