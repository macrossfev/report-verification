# 水质检测报告验证分析系统

本项目用于自动化验证水质检测报告的数据一致性和完整性，支持原始记录检查、报告文件检查、交叉验证及公示表比对等多种模式。

## 环境要求

- Python 3.8+
- 依赖包见 `requirements.txt`

## 快速开始

```bash
# 赋予执行权限
chmod +x start.sh

# 查看帮助
./start.sh
```

## 四种运行模式

### 1. `-oridata` 仅检查原始记录

仅扫描原始记录文件（.xls），检查数据完整性、格式规范等，不涉及报告文件。

```bash
./start.sh -oridata
./start.sh -oridata -r /path/to/data
```

### 2. `-report` 仅检查报告文件

仅扫描报告文件（.xlsx），检查报告自身的格式、编号、数据等问题，不做原始记录交叉验证。

```bash
./start.sh -report
./start.sh -report -r /path/to/reports
```

### 3. `-datareport` 基于原始记录检查报告（完整交叉验证）

同时读取原始记录和报告文件，将两者数据进行交叉比对，检查数据是否一致。

```bash
./start.sh -datareport
./start.sh -datareport -r /path/to/reports
```

### 4. `-public` 公示表与电子报告交叉比对

将公示表（Publicsheet）中的数据与电子报告进行比对，检查对外公示数据的一致性。

```bash
./start.sh -public
./start.sh -public -r /path/to/Publicsheet
```

## 其他选项

| 选项 | 说明 |
|------|------|
| `-r <目录>` | 指定扫描目录 |
| `-o <文件>` | 自定义输出文件路径 |
| `-h` | 查看完整帮助信息 |

## 项目结构

```
report-verification/
  analyze_reports.py    # 主分析脚本
  start.sh              # 启动脚本（自动配置环境）
  requirements.txt      # Python 依赖
  report/               # 报告文件目录
  water/                # 原始记录目录
  Publicsheet/          # 公示表目录
```
