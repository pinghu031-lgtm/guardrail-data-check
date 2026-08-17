# 护栏数据异常检查工具（云端共享版）

本项目将“护栏数据异常检查工具 v1.7”的 Excel 检查与自动修改规则封装为 Streamlit 应用。代码由 GitHub 托管，可部署到免费的 Streamlit Community Cloud，使用者通过浏览器上传文件并下载结果。

## 功能

- 批量上传 `.xlsx`、`.xlsm` 文件；
- 检查护栏高度异常标记；
- 检查螺栓缺失备注及负数异常；
- 将连续异常合并为桩号区间；
- 下载包含“护栏高度”和“螺栓缺失”工作表的异常汇总；
- 下载“标黄版”和“不标黄版”已修改数据 ZIP；
- 可通过 Streamlit Secret 配置页面访问口令。

## 数据处理方式

每次点击“开始处理”都会创建独立临时目录。上传文件、异常汇总和修改结果只在该次处理期间使用，返回下载数据后临时目录即被清理。程序不会把用户上传的 Excel 写入 GitHub 仓库。

默认限制：

- 单次最多 20 个文件；
- 单文件最大 50 MiB；
- 单次总大小最大 200 MiB；
- 仅接受 `.xlsx` 和 `.xlsm`。

> 本项目使用第三方云端运行环境。具有保密要求或明确禁止外传的数据，应部署到受控的单位服务器，不应上传到公共云服务。

## 项目结构

```text
.
├── app.py                    # Streamlit 页面入口
├── cloud_service.py          # 上传校验、临时隔离、打包下载
├── guardrail_core.py         # v1.7 Excel 业务处理核心
├── requirements.txt          # 云端运行依赖
├── requirements-dev.txt      # 测试依赖
├── tests/                    # 自动化测试
└── .github/workflows/        # GitHub Actions 测试
```

## 本地运行

要求 Python 3.11。

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

浏览器打开终端显示的地址，通常为：

```text
http://localhost:8501
```

## 运行测试

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
python -m compileall -q app.py cloud_service.py guardrail_core.py tests
```

## 推送到 GitHub

在本项目目录执行：

```bash
git init -b main
git add .
git commit -m "feat: 新增护栏数据异常检查云端版"
git remote add origin https://github.com/<用户名>/guardrail-data-check.git
git push -u origin main
```

仓库中不得提交真实业务 Excel、访问口令或其他敏感信息；`.gitignore` 已默认排除 Excel、ZIP 和本地 Secret。

## 部署到 Streamlit Community Cloud

1. 登录 [Streamlit Community Cloud](https://share.streamlit.io/)；
2. 使用 GitHub 账号授权 Streamlit 读取目标仓库；
3. 点击 **Create app**；
4. 选择 GitHub 仓库 `guardrail-data-check`；
5. 分支填写 `main`；
6. Main file path 填写 `app.py`；
7. 点击 **Deploy**；
8. 部署完成后访问平台生成的 `https://*.streamlit.app` 地址。

Streamlit 与 GitHub 连接后，向 `main` 分支推送更新会触发应用重新部署。

官方文档：

- [Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud)
- [连接 GitHub 账号](https://docs.streamlit.io/deploy/streamlit-community-cloud/get-started/connect-your-github-account)
- [部署应用](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app)

## 配置访问口令

应用未配置口令时公开访问。需要口令时，在 Streamlit 应用设置的 **Secrets** 中添加：

```toml
APP_PASSWORD = "由管理者设置的口令"
```

保存后应用会重新运行。真实口令不得写入 `app.py`，也不得提交 `.streamlit/secrets.toml`。

本地调试可创建 `.streamlit/secrets.toml` 并写入同样的配置，该文件已经被 `.gitignore` 排除。

## 更新业务规则

业务规则集中在 `guardrail_core.py`。修改规则时应同步增加或更新 `tests/` 中的回归测试，并在推送前执行完整测试。GitHub Actions 也会在每次推送及 Pull Request 时自动运行测试。
