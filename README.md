# Excel 工具箱

浏览器里处理 Excel：拆分、合并、透视，全程不用装 Excel、不用写 VBA。

处理完即释放，不保存、不外传、不做二次利用。

## 功能

| Tab | 做什么 | 关键点 |
|---|---|---|
| 按列拆分 | 按 1–2 列的唯一值把一张表拆成多个 xlsx 并打包 zip | 文件名自动去掉 `\ / : * ? " < > \|`，空值归入「空值」，超过 300 个分组直接拦下 |
| 多表合并 | 多个文件 / 多个 sheet 按列名合并 | 并集或交集对齐，缺列留空，可加「来源」列 |
| 数据透视表 | 行维度 × 列维度 × 指标聚合 | 自动带合计行 / 合计列，可切柱状图，导出 xlsx |

三个 Tab 都支持同一套「数据清洗」选项：默认去空格、全角转半角、把 `-` / `无` / `N/A` 统一成空值；数值化、日期化、去重按需打开。

## 运行方式

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

浏览器打开 http://localhost:8501 。Windows + Python 3.11 实测可用。

## 目录结构

```text
├─ app.py                  # 入口：页面配置 + 三个 Tab 编排
├─ core/                   # 纯 Python 处理层，不 import streamlit
│  ├─ loader.py            # 读文件 → DataFrame；表头探测；多 sheet
│  ├─ cleaner.py           # 全角归一、空值统一、数值化、日期解析、去重
│  ├─ splitter.py          # 按列拆分 + 文件名清洗
│  ├─ merger.py            # 多表合并
│  ├─ pivot.py             # 数据透视表
│  └─ exporter.py          # DataFrame → xlsx / zip 字节流
├─ ui/
│  ├─ components.py        # 上传区、清洗面板、预览、下载、错误提示
│  └─ pages/               # split.py / merge.py / pivot.py
├─ data/sample_sales.xlsx  # 1000 行演示数据（含脏数据）
├─ tests/test_core.py      # core 层单元测试
└─ .streamlit/config.toml  # 上传上限 50MB、主题
```

## 自测

```bash
python tests/test_core.py
```

用 `data/sample_sales.xlsx` 可以走一遍完整流程：

- 拆分：按「省份」拆 → 32 个文件（31 个省份 + 1 个空值）+ `_拆分说明.txt`，行数合计等于原表；
- 合并：拆出来的两个文件再合回去，行数应等于原表；
- 透视：行「省份」× 列「季度」→ 合计行等于销售额总和（先开「数值化」）。

## 部署

本地或局域网（同一 WiFi 下手机可直接访问）：

```bash
python -m streamlit run app.py --server.address 0.0.0.0
```

手机浏览器打开 `http://电脑IP:8501`，注意 Windows 防火墙放行该端口。

Streamlit Community Cloud：

1. 把仓库推到 GitHub；
2. 在 share.streamlit.io 新建应用，主文件填 `app.py`；
3. 依赖由 `requirements.txt` 自动安装，无需额外配置。

## 常见问题

- **.xls 读不了**：扩展名是 `.xls` 但内容其实是 xlsx / CSV 的（手机、系统导出常见）会自动识别；真正的老版二进制 `.xls` 请在 Excel / WPS 里「另存为 .xlsx」再上传。
- **表头识别错了**：页面上的「表头行号」可以手动改；如果提示「疑似表头在第 N 行」，优先按提示核对。
- **金额列没法求和**：先打开「数据清洗 → 数值化」，`¥1,234`、`12%` 这类文本会转成数字（百分号按 0.12 折算）。
- **日期列没变成日期**：日期化只处理列名带「日期 / 时间 / date / time」的列。
- **身份证号前导零丢了**：打开「全部按文本读取」再上传。
- **分组太多**：超过 300 个文件会被拒绝，换一列分组或先筛选数据。

## 进度

| 天 | 内容 | 状态 |
|---|---|---|
| D1 | 骨架 + 上传解析 | ☑ |
| D2 | 拆分核心 | ☑ |
| D3 | 拆分容错 + zip + 说明文件 | ☑ |
| D4 | 多表合并 | ☑ |
| D5 | 数据透视表 | ☑ |
| D6 | 清洗 + 柱状图 + 表头启发式 | ☑ |
| D7 | 示例数据 + README / 部署 + 手机实测 | 部分 ☑（部署待做） |
| D8 | 缓冲：修 bug、提交 | ☐ |
