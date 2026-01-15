# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个证书照片分类工具 (CertPhotoSorter)，用于根据 Excel 报名表中的信息，将学员照片按证书类型自动分类。

**核心功能**：
- 从 Excel 文件读取身份证号码、证书、姓名三列数据
- 递归扫描照片目录，从文件名中提取身份证号进行匹配
- 按证书类型创建文件夹并复制对应照片
- 生成详细的运行报告和 CSV 明细

## 构建命令

```powershell
# 构建可执行文件（生成 x64 和 x86 两个版本）
pwsh build_exe.ps1
```

**构建产物**：`dist/CertPhotoSorter_x64.exe` 和 `dist/CertPhotoSorter_x86.exe`

**构建依赖**：
- .NET Framework 4.x（使用 Windows 自带的 `csc.exe` 编译器）
- 无需外部包管理器（nuget 等）

## 运行方式

### GUI 模式
双击 EXE 文件即可启动图形界面。

### CLI 模式
```bash
CertPhotoSorter_x64.exe --excel "报名表.xlsx" --photos "照片目录" [--out "输出目录"] [--sheet "工作表名"] [--dry-run] [--log "日志路径"]
```

**参数说明**：
- `--excel`：Excel 文件路径（必需）
- `--photos`：照片根目录（必需）
- `--out`：输出目录（可选，默认在 Excel 同目录下生成"按证书分类"文件夹）
- `--sheet`：指定工作表名（可选，默认自动识别）
- `--dry-run`：仅生成报告，不复制文件
- `--log`：日志文件路径（可选）

## 代码架构

### 入口点
- `Program.cs`：根据命令行参数判断启动 GUI 或 CLI 模式

### 双界面层
- `MainForm.cs`：Windows Forms GUI（支持文件夹选择、工作表加载、进度条）
- `CliRunner.cs`：命令行参数解析和执行入口

### 核心业务逻辑
- `Processor.cs`：主处理流程
  1. 读取 Excel 数据构建 `idToCerts` 映射
  2. 递归扫描照片文件
  3. 从文件名提取身份证号进行匹配
  4. 按证书分类复制照片
  5. 生成报告和 CSV 文件

### 数据访问层
- `ExcelReader.cs`：Excel 读取，采用 **三重 Provider 回退机制**：
  1. `Microsoft.ACE.OLEDB.16.0`
  2. `Microsoft.ACE.OLEDB.12.0`
  3. `Microsoft.Jet.OLEDB.4.0`

  **工作表自动识别**：优先选择包含"身份证号码"和"证书"列且名称包含"最终"或"报名"的工作表，按行数降序排列。

### 工具模块
- `IdUtils.cs`：身份证号提取逻辑
  - 优先匹配 18 位身份证号
  - 其次匹配 17 位 + 计算校验位
  - 最后匹配 15 位老版身份证号
  - 校验位算法：GB 11643-1999 标准（权重 [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]）

- `CsvUtils.cs`：CSV 字段转义（RFC 4180）

### 数据模型
- `Models.cs`：
  - `RunSettings`：运行配置
  - `RunResult`：运行结果统计
  - `OpRow`：每张照片的操作记录

### 本地化
- `Texts.cs`：所有中文字符串常量（硬编码 Unicode 转义以兼容旧编译器）

## Excel 列要求

Excel 必须包含以下三列（列名不可更改）：
- `身份证号码`：用于与照片文件名匹配
- `证书`：作为输出文件夹名称
- `姓名`：用于报告记录

## 照片匹配逻辑

1. 从文件名（不含扩展名）中提取身份证号
2. 在 Excel 中查找对应记录
3. **一对多处理**：同一身份证号如对应多个证书，照片会复制到所有证书文件夹
4. **未匹配照片**：放入"未匹配"文件夹
   - 文件名中无身份证号 → 标记为"未匹配-无身份证号"
   - 身份证号不在 Excel 中 → 标记为"未匹配-身份证号不在Excel"

## 输出文件结构

```
输出目录/
├── [证书名称A]/
│   └── 匹配的照片文件
├── [证书名称B]/
│   └── 匹配的照片文件
├── 未匹配/
│   └── 无法匹配的照片
├── 运行报告_*.txt
├── 明细_*.csv
└── 证书汇总_*.csv
```

## 文件名冲突处理

使用 `文件名_序号.扩展名` 格式自动去重，序号从 2 开始递增。

## 证书文件夹名称处理

证书名称中的非法文件名字符会被替换为下划线，并去除首尾空格和点。
