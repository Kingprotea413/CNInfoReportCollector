# 巨潮资讯年报资产负债表采集器

这是一个用 Python 编写的桌面采集工具，用来调用巨潮资讯 `p_stock2300` 接口，抓取单个公司的多年资产负债表，并筛选出年报后导出到本地 Excel。

默认公司是 `长江电力`，每次只处理一个公司。程序从终端启动，但会弹出一个桌面窗口，输入公司名称或证券代码后点击按钮即可开始采集。

## 功能

- 调用巨潮资讯 `p_stock2300`
- 自动生成 `Accept-EncKey` 请求头
- 自动检索公司名称/证券代码
- 只保留年报数据
  - 当前筛选规则：`ENDDATE` 为 `12-31`
  - 且 `F003V` 为 `合并本期`
- 按模板报表格式导出到 `outputs/`
  - 第一列是报表项目
  - 后续列是各年度报告期
  - 行名与用户提供的参考 Excel 一致
- 支持选择导出单位
  - GUI 可选 `元`、`千元`、`万元`、`亿元`
  - 命令行可通过 `--unit` 指定
- 提供 GUI 进度条和状态提示
- 提供命令行无界面模式，方便测试

## 环境

- Windows
- Python 3.12
- 依赖见 `requirements.txt`

## 安装依赖

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" -m pip install -r requirements.txt
```

## 启动 GUI

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py
```

## 命令行模式

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --headless --company 长江电力
```

指定单位示例：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --headless --company 长江电力 --unit 万元
```

## 输出说明

导出的 Excel 文件名格式：

```text
证券代码_证券简称_annual_balance_sheet.xlsx
```

例如：

```text
600900_长江电力_annual_balance_sheet.xlsx
```

当前导出会尽量按参考模板行名输出。直接字段、合计项和常用比率都已做了格式化整理。

## 测试

运行单元测试：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" -m unittest discover -s .\tests -v
```

运行真实联网 smoke test：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --headless --company 长江电力
```

测试 GUI 是否能成功创建窗口：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --self-test-gui
```

## GitHub Actions 自动打包

仓库已包含 GitHub Actions 工作流 `/.github/workflows/build-exe.yml`：

- 推送到 `main` 时自动在 GitHub 的 Windows runner 上打包 `CNInfoReportCollector.exe`
- 提交 Pull Request 时自动验证能否成功打包
- 可在 GitHub 的 `Actions` 页面手动触发
- 生成的 exe 会作为 workflow artifact 上传，名称为 `CNInfoReportCollector-windows-exe`
