# 项目规则（Excel 工具箱）

## 架构分层

- `app.py` 只做页面配置与 Tab 编排，不写业务逻辑。
- `ui/pages/*` 只负责收集参数与展示，不写 pandas 处理。
- `core/*` 只做数据处理，**禁止 import streamlit**。
- 所有导出走 `io.BytesIO`，不写磁盘、不落临时文件。
- 不引入 `requirements.txt` 之外的依赖（altair 为 Streamlit 自带传递依赖，已获准用于饼图）。

## 错误与文案

- 统一用 `ExcelToolError(message, reason, suggestion)`，中文三段式，不向用户抛 traceback。
- 提示里不出现英文技术术语（如 DataFrame、traceback），必要术语放在「可能原因」里。

## 交互约定

- 每批最多改 3 个业务文件，做完暂停等验收。
- 破坏性操作（覆盖、删除已有数据）必须二次确认。
- 数据只在内存中处理，不接第三方 API，不写日志中的业务数据。

## 功能口径

- 上传上限 50MB / 20 万行；拆分、套名最多 300 个输出文件。
- 表头：自动探测 + 手动「表头行号」兜底。
- 批量套名占位符固定为 `XXX`（全局替换），生成文件名用姓名。
- 图表支持柱状图 / 折线图 / 饼图。

## 测试

- 改 `core/` 必须补 `tests/test_core.py` 用例。
- 提交前执行 `python tests/test_core.py`，要求全绿。
