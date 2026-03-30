# 巨潮资讯年报资产负债表采集器

这是一个用 Python 编写的桌面采集工具，用来调用巨潮资讯 `p_stock2300` 接口，抓取单个公司的多年资产负债表，并筛选出年报后导出到本地 Excel。

默认公司是 `长江电力`。程序支持 GUI 和命令行两种使用方式。

## 功能

- 调用巨潮资讯 `p_stock2300`
- 自动生成 `Accept-EncKey` 请求头
- 自动搜索公司名称或证券代码
- 仅保留年报数据
  - `ENDDATE` 以 `12-31` 结尾
  - `F003V` 为 `合并本期`
- 按模板格式导出 Excel
- 支持导出单位切换
  - GUI 可选 `元`、`千元`、`万元`、`亿元`
  - 命令行可通过 `--unit` 指定
- GUI 内显示进度和状态
- 提供无界面命令行模式，便于测试和自动化

## 环境

- Windows 或 macOS
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

指定导出单位示例：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --headless --company 长江电力 --unit 万元
```

## 输出说明

- GUI 模式下，Excel 会保存到窗口中选择的目录
- 命令行模式下，默认输出目录是 `outputs/`，也可以通过 `--output-dir` 指定
- 导出文件名格式：

```text
证券代码_证券简称_annual_balance_sheet.xlsx
```

示例：

```text
600900_长江电力_annual_balance_sheet.xlsx
```

## 测试

运行单元测试：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" -m unittest discover -s .\tests -v
```

运行真实联网 smoke test：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --headless --company 长江电力
```

验证 GUI 能否创建：

```powershell
& "$env:LocalAppData\Programs\Python\Python312\python.exe" .\app.py --self-test-gui
```

## GitHub Actions

仓库包含两个自动化工作流：

- `/.github/workflows/build-exe.yml`
  - push 到 `main`、提交 Pull Request，或手动触发时自动构建
  - 输出 Windows `exe + zip`
  - 输出 macOS `arm64 zip` 和 `x64 zip`
- `/.github/workflows/release.yml`
  - push 形如 `v0.1.0` 的 tag 时自动创建 GitHub Release
  - Release 会附带 Windows `exe`、Windows `zip`、macOS `arm64 zip`、macOS `x64 zip`

发布示例：

```powershell
git tag v0.1.0
git push origin v0.1.0
```

## macOS 说明

- Apple Silicon 芯片的 Mac 使用 `macos-arm64.zip`
- Intel 芯片的 Mac 使用 `macos-x64.zip`
- 当前 macOS 构建未签名、未 notarize
- 第一次打开时，可能需要右键选择“打开”，或在系统设置里手动放行
- 如果运行失败，错误日志会写到：
  - macOS: `~/Library/Application Support/CNInfoReportCollector/last_error.log`
  - Windows: `%LOCALAPPDATA%\CNInfoReportCollector\last_error.log`
