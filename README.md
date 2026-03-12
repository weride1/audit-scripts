# Audit Scripts - 内部审计自动化工具集

> A collection of Python-based utilities for internal audit automation, data processing, and Excel reporting.
> 基于 Python 的内部审计自动化工具集，用于数据处理与报表生成。

## 📂 项目简介 (Overview)

本仓库是一个集中管理内部审计相关脚本的工具库。旨在通过自动化脚本提高审计工作效率，减少重复性手工操作。

**主要用途：**
-  **自动化处理**：批量处理财务、业务数据。
- 📊 **报表生成**：自动从源数据生成标准化的 Excel/PDF 报告。
- 🔍 **数据校验**：自动核对数据一致性，发现异常点。
- ️ **工具函数**：提供常用的数据清洗、格式转换等辅助功能。

> 💡 **注意**：本仓库仅存储**代码逻辑**。所有敏感数据文件（如 `.xlsx`, `.csv`, `.pdf`）均已被 `.gitignore` 规则屏蔽，严禁上传至仓库。

## 🚀 通用使用指南 (General Usage)

### 1. 环境准备
确保您的电脑已安装 **Python 3.8+**。

### 2. 安装依赖
大多数脚本依赖以下核心库。建议在运行任何新脚本前，先安装基础依赖：

```bash
pip install pandas openpyxl xlsxwriter
