# LiFSI Aspen 工艺设计资料

## 项目介绍

本项目用于整理 LiFSI 连续化合成工艺设计相关资料，包含 Aspen Plus 建模步骤、组件输入表、参数表、分析表模板以及 WPS/Office 自动化辅助脚本，主要服务于课程设计和工艺方案整理。

## 技术栈

- Aspen Plus 工艺模拟
- PowerShell 自动化脚本
- VBA/WPS 宏
- Markdown 文档资料
- CSV 表格模板

## 部署要求

- Aspen Plus（用于流程建模）
- WPS 或 Microsoft Office（用于公式/文档处理）
- Windows PowerShell
- 基础化工工艺设计数据

## 运行流程

1. 根据 Aspen 组件输入表建立物性和组分数据。
2. 按建模实施步骤在 Aspen Plus 中搭建流程。
3. 使用参数表补充 Block、物流和反应条件。
4. 根据 HAZOP 模板整理风险分析。
5. 如需生成公式文档，可运行 PowerShell/VBA 辅助脚本。

## 项目结构

- Aspen_*.md：建模数据、组件和 Block 参数说明
- HAZOP_分析表_模板.csv：风险分析模板
- create_wps_formula_doc.ps1：WPS 文档辅助脚本
- WpsInsertFormula.bas：WPS/Office 宏脚本

## 上传说明

本仓库只保留项目运行和二次开发所需的代码、配置、数据库脚本及少量必要静态资源。

以下内容不会上传：论文、答辩材料、临时文档、依赖目录、构建产物、压缩包、数据集、模型权重、视频、日志、本地工具包以及密钥配置。
