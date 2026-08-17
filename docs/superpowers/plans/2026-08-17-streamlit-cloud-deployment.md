# 护栏数据异常检查工具云端部署实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 subagent-driven-development（推荐）或 executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 将 v1.7 护栏 Excel 检查与修改规则封装为由 GitHub 托管、可在 Streamlit Community Cloud 免费部署并通过网址访问的多文件上传/下载应用。

**架构：** 从 v1.7 提取与界面无关的 Excel 业务核心，新增无状态云端批处理服务；每次请求在独立临时目录中保存上传文件，输出检查结果 Excel 和已修改数据 ZIP，并在返回字节数据后清理目录。Streamlit 入口只负责访问口令、文件上传、进度展示、结果预览和下载，不保存跨用户全局文件路径。

**技术栈：** Python 3.11、openpyxl、Streamlit、pytest、GitHub Actions。

---

## 文件结构

- `guardrail_core.py`：从 v1.7 提取的数据模型、表头识别、异常判断、Excel 汇总和自动修改逻辑。
- `cloud_service.py`：上传文件校验、安全文件名、请求级临时目录、批处理、结果序列化和 ZIP 打包。
- `app.py`：Streamlit 页面入口及可选环境密钥口令验证。
- `tests/test_cloud_service.py`：自动生成 Excel，验证上传校验、处理结果和 ZIP 结构。
- `tests/test_core_rules.py`：对核心高度和螺栓规则做特征测试，确保提取后规则不漂移。
- `requirements.txt`、`requirements-dev.txt`：运行和测试依赖。
- `.streamlit/config.toml`：上传上限、无头运行和遥测配置。
- `.github/workflows/tests.yml`：GitHub 推送时运行测试。
- `.gitignore`：排除真实 Excel、口令、缓存和输出文件。
- `README.md`：使用、GitHub、Streamlit 部署和数据安全说明。

### 任务 1：提取核心并建立特征测试

**文件：**
- 创建：`guardrail_core.py`
- 创建：`tests/test_core_rules.py`

- [ ] **步骤 1：编写失败的特征测试**

测试通过内存工作簿构造护栏高度和螺栓缺失数据，调用 `InspectionProcessor.process_files()`，断言输出“零值标记错误”和“异常数值”。

- [ ] **步骤 2：运行测试验证失败**

运行：`python -m pytest tests/test_core_rules.py -v`
预期：因 `guardrail_core` 不存在而失败。

- [ ] **步骤 3：提取最少业务核心**

复制 v1.7 中从常量、数据模型到 `export_modified_workbooks()` 的业务代码，删除 Tkinter、HTTP 服务、浏览器启动和全局桌面状态依赖，保持业务函数签名和规则不变。

- [ ] **步骤 4：运行测试验证通过**

运行：`python -m pytest tests/test_core_rules.py -v`
预期：2 个特征测试通过。

### 任务 2：实现隔离的云端批处理服务

**文件：**
- 创建：`cloud_service.py`
- 创建：`tests/test_cloud_service.py`

- [ ] **步骤 1：编写失败的上传校验测试**

测试 `.txt` 被拒绝、超过文件数量被拒绝、重复文件名被安全去重，并断言异常信息为中文。

- [ ] **步骤 2：运行测试验证失败**

运行：`python -m pytest tests/test_cloud_service.py -k validate -v`
预期：因服务接口不存在而失败。

- [ ] **步骤 3：实现上传数据模型与校验**

实现 `UploadedFileData(name: str, content: bytes)`、`validate_uploads()` 和 `safe_filename()`；仅接受 `.xlsx/.xlsm`，默认单次最多 20 个文件、单文件 50 MiB、总计 200 MiB，并阻断路径穿越和空文件。

- [ ] **步骤 4：运行上传校验测试验证通过**

运行：`python -m pytest tests/test_cloud_service.py -k validate -v`
预期：上传校验测试全部通过。

- [ ] **步骤 5：编写失败的端到端批处理测试**

自动生成含护栏高度和螺栓缺失表的工作簿，调用 `process_uploads()`，断言结果表有两个工作表、结果预览包含异常、ZIP 同时包含“标黄版”和“不标黄版”文件，并断言临时目录不泄漏到返回值。

- [ ] **步骤 6：运行端到端测试验证失败**

运行：`python -m pytest tests/test_cloud_service.py -k process -v`
预期：因 `process_uploads()` 未实现而失败。

- [ ] **步骤 7：实现请求级临时处理**

使用 `TemporaryDirectory` 写入校验后的上传文件；调用 `InspectionProcessor`、`export_groups_to_excel()` 和 `export_modified_workbooks()`；将结果 Excel、修改数据 ZIP、预览行、统计和日志读取为内存对象后返回。

- [ ] **步骤 8：运行端到端测试验证通过**

运行：`python -m pytest tests/test_cloud_service.py -v`
预期：全部通过。

### 任务 3：实现 Streamlit 页面和可选口令

**文件：**
- 创建：`app.py`
- 创建：`.streamlit/config.toml`
- 修改：`requirements.txt`
- 修改：`requirements-dev.txt`

- [ ] **步骤 1：编写失败的应用冒烟测试**

使用 `streamlit.testing.v1.AppTest` 加载 `app.py`，断言页面标题、文件上传控件和“开始处理”按钮存在，且没有启动异常。

- [ ] **步骤 2：运行冒烟测试验证失败**

运行：`python -m pytest tests/test_streamlit_app.py -v`
预期：因 `app.py` 不存在而失败。

- [ ] **步骤 3：实现最少 Streamlit 页面**

页面提供多文件上传、文件清单、处理按钮、统计指标、异常结果表、日志折叠区、结果 Excel 下载和修改数据 ZIP 下载；`APP_PASSWORD` 未配置时公开访问，配置后使用常量时间比较验证口令。

- [ ] **步骤 4：运行冒烟测试验证通过**

运行：`python -m pytest tests/test_streamlit_app.py -v`
预期：页面无异常且关键控件存在。

### 任务 4：补齐仓库配置和部署文档

**文件：**
- 创建：`.github/workflows/tests.yml`
- 创建：`.gitignore`
- 创建：`README.md`

- [ ] **步骤 1：配置 GitHub Actions**

在 Windows/Linux 无关的 Ubuntu Python 3.11 环境安装 `requirements-dev.txt` 并运行 `python -m pytest -q`。

- [ ] **步骤 2：编写部署文档**

记录本地运行命令、GitHub 推送、Streamlit Community Cloud 创建应用、入口文件 `app.py`、可选 `APP_PASSWORD` Secret、文件限制、临时存储和敏感数据注意事项。

- [ ] **步骤 3：执行配置自检**

运行 Python TOML/YAML 基本读取及 `python -m compileall .`，确认文件语法有效。

### 任务 5：完整验证、审查、GitHub 与上线

**文件：**
- 检查：全部变更文件

- [ ] **步骤 1：运行完整测试与语法检查**

运行：`python -m pytest -q`、`python -m compileall -q app.py cloud_service.py guardrail_core.py tests`。
预期：零失败、命令退出码均为 0。

- [ ] **步骤 2：启动并探测 Streamlit**

运行：`python -m streamlit run app.py --server.headless true --server.port 8501`；轮询 `http://127.0.0.1:8501/_stcore/health`，预期返回 `ok`，再用浏览器检查标题、上传控件和按钮。

- [ ] **步骤 3：执行独立代码审查**

检查上传验证、路径穿越、ZIP 内容、临时目录清理、敏感信息、并发隔离和业务规则保持；修复所有阻断问题后重新运行完整验证。

- [ ] **步骤 4：创建 Git 仓库和提交**

在项目目录初始化 `main` 分支，提交经验证的源程序、测试和文档，不提交任何真实 Excel 或口令。

- [ ] **步骤 5：连接 GitHub**

若本机已有 GitHub 凭据，则创建公开仓库 `guardrail-data-check` 并推送；若未登录，则保留本地已验证提交，明确要求用户完成一次 GitHub 授权后继续。

- [ ] **步骤 6：部署 Streamlit Community Cloud**

使用 GitHub 登录 Streamlit Community Cloud，选择仓库、`main` 分支和 `app.py`；如遇登录或授权墙，停止并请用户完成授权。部署后访问公网 URL，并验证页面能加载和处理测试 Excel。
