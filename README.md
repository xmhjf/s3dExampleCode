# s3dExampleCode

Smart 3D Programming ExampleCode From S3D 2016

## 项目简介

这是一个 Smart 3D (S3D) 2016 版本的编程示例代码仓库，提供了多个领域的开发示例，帮助开发人员更好地理解和使用 Smart 3D 的开发接口。本仓库包含了结构设计、结构制造、管架支撑以及设备等多个领域的示例代码，展示了如何使用 Smart 3D API 进行二次开发。

## 目录结构

### 1. 结构设计 (StructDetail)
- 自定义装配 (CustomAssemblies)
  - 边缘特征 (EdgeFeatures) - 用于处理结构边缘特征的自定义实现
  - 连接 (Connections) - 结构连接的自定义定义
  - 穿透 (Penetrations) - 结构穿透处理
- 规则 (Rules)
  - 端部切割边缘映射规则 (EndCutEdgeMappingRule) - 处理结构端部切割的映射关系

### 2. 结构制造 (StructManufacturing)
- 规则 (Rules)
  - 制造模板集 (MfgTemplateSet)
    - 几何规则 (GeometryRule)
    - 管状模板几何规则 (TubularTemplateGeometryRule)
  - 收缩规则 (Shrinkage)
    - 参数规则 (ParameterRule)
      - 装配部件规则 (AssemblyPartSumRule, AssemblyPartIncrementRule)
      - 内置构件规则 (BuiltUpProfileCustomRule, BuiltUpMemberCustomRule)
      - 面板装配规则 (PanelAssemblyRule)
      - 预制分段规则 (PreSubBlockRule)
  - 自定义报告规则 (CustomReportRules)
    - 构件报告 (ProfileReports) - 包含重量报告和依赖检查
    - 构件报告 (MemberReports) - 制造构件自定义报告
    - 嵌套报告 (NestingReports) - 零件嵌套布局报告

### 3. 管架与支撑 (HangersAndSupports)
- 插件 (Plugins)
  - 第三方插件示例 (Example3rdPartyPlugin) - 演示如何开发自定义第三方插件
- 规则 (Rules)
  - 自定义支撑定义 (CustomSupportDefinition)
    - 管道支撑 (Pipe)
      - CSD 示例 (CSD_Sample)
      - FINL 装配 (FINL_Assy)
      - Oglaend 装配 (OglaendAssemblies)
- 符号 (Symbols)
  - 电缆桥架 (CableTray)
    - 桥架装运 (TrayShip)

### 4. 设备 (Equipment)
- 规则 (Rules)
  - 设备命名规则 (EquipmentNamingRules)
- 用户自定义表单 (UserDefinedFormDefinitions)
  - 泵装配表单定义 (PumpAsmFormDefinition)

### 5. 船舶结构 (ShipStructure)
- 几何构造 (GeometricConstructions)
  - 高级板系统 (AdvancedPlateSystems)

## 主要功能

1. **第三方插件开发**
   - 提供完整的第三方插件开发示例
   - 演示如何与 Smart 3D 核心程序交互
   - 包含必要的程序集引用说明
   - 支持 C# 和 VB.NET 两种开发语言

2. **自定义规则开发**
   - 结构设计规则：端部切割、边缘特征处理
   - 制造规则：模板定义、收缩规则、参数化规则
   - 命名规则：设备、网格等对象的命名规则
   - 报告生成规则：构件报告、嵌套报告等

3. **自定义装配体开发**
   - 边缘特征处理：自定义边缘特征定义
   - 连接定义：结构连接的自定义实现
   - 支撑系统：管道支撑、电缆桥架等
   - 设备装配：泵类设备的自定义装配

4. **资源本地化**
   - 支持多语言资源文件 (.resx)
   - 包含英文 (en-US) 资源定义
   - 可扩展的资源管理机制

## 开发环境要求

- Smart Plant 3D 2016
- .NET Framework 4.0 或更高版本 (支持 C# 和 VB.NET)
- Visual Studio 2010 或更高版本
- SQL Server（用于目录数据库）

## 必要的依赖项

核心程序集位置：`$SmartPlant3D/core/Container/Bin/Assemblies/Release/`

主要引用：
- CommonMiddle.dll - 核心中间层组件
- CommonClient.dll - 客户端通用组件
- HSImportAssemblies.dll - 管架支撑导入组件
- GridsMiddle.dll - 网格系统中间层组件

## 使用说明

1. 环境配置
   - 确保已正确安装 Smart Plant 3D 2016
   - 配置必要的开发环境
   - 设置正确的程序集引用路径

2. 项目配置
   - 根据具体示例的要求配置必要的程序集引用
   - 确保资源文件 (.resx) 正确加载
   - 检查命名空间是否正确设置

3. 编译部署
   - 编译相关项目
   - 生成必要的 CAB 文件
   - 部署到正确的符号共享目录

4. 测试验证
   - 按照各示例中的 ReadMe 文件进行测试
   - 验证功能是否正常工作
   - 检查本地化资源是否正确加载

## 注意事项

- 使用前请确保已获得相关许可
- 遵循 Intergraph Corporation 的版权声明
- 建议在开发环境中进行测试，避免在生产环境直接使用示例代码
- 注意保持正确的程序集引用路径
- 确保符号共享目录配置正确
- 建议先阅读相关示例的 ReadMe 文件

## 版权声明

Copyright (C) 2009, Intergraph Corporation. All rights reserved.

## 贡献指南

- 遵循现有的代码结构和命名规范
- 确保添加适当的注释和文档
- 保持资源文件的多语言支持
- 测试用例需覆盖新增功能

## 相关文档

- Smart Plant 3D API 文档
- 各模块 ReadMe 文件
- 示例代码中的注释说明

