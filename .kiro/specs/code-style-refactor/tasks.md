# Implementation Plan: Code Style Refactor

## Overview

本实现计划将项目代码按照 CONTRIBUTE.md 规范进行重构。采用渐进式方法，按模块逐步改造，确保每个阶段都能独立验证。

## Tasks

- [x] 1. 项目基础设施配置
  - [x] 1.1 更新 pyproject.toml 添加新依赖
    - 添加 loguru、pydantic-settings、ruff、pytest、hypothesis 依赖
    - 配置 Ruff linter 和 formatter
    - 配置 pytest
    - _Requirements: 8.1, 9.1_
  - [x] 1.2 创建 .pre-commit-config.yaml
    - 配置 ruff check 和 ruff format hooks
    - 配置 uv lock --check hook
    - _Requirements: 8.3_
  - [x] 1.3 创建测试目录结构
    - 创建 tests/ 目录
    - 创建 tests/conftest.py 共享 fixtures
    - _Requirements: 9.2_

- [x] 2. 配置管理模块
  - [x] 2.1 创建 src/core/config.py
    - 使用 pydantic-settings 定义 Settings 类
    - 定义所有配置项（temp_dir, default_max_tokens 等）
    - _Requirements: 6.1, 6.3, 6.4_
  - [ ]* 2.2 编写配置模块单元测试
    - 测试配置加载和验证
    - _Requirements: 6.4_

- [x] 3. 数据模型模块
  - [x] 3.1 创建 src/core/models.py
    - 定义 SplitMode、TokenStrategy 枚举
    - 定义 ChunkConfig、ChunkResult、ConversionResult dataclass
    - 定义 MergedCellInfo、TableNote、ChunkStats dataclass
    - _Requirements: 3.2, 5.4, 7.2_
  - [x] 3.2 创建 Pydantic 验证模型
    - 定义 ProcessRequest 模型用于外部输入验证
    - _Requirements: 7.1_
  - [ ]* 3.3 编写数据模型属性测试
    - **Property 8: Dataclass for Data Objects**
    - **Validates: Requirements 3.2, 7.2**
  - [ ]* 3.4 编写枚举常量属性测试
    - **Property 9: Enum for Constants**
    - **Validates: Requirements 5.4**

- [x] 4. Checkpoint - 基础模块验证
  - 确保所有测试通过，如有问题请询问用户

- [x] 5. Excel 转换器重构
  - [x] 5.1 重构 excel2html_openpyxl_enhanced.py 为 converter.py
    - 将函数封装为 ExcelToHtmlConverter 类
    - 替换 print 为 loguru logger
    - 替换 os.path 为 pathlib
    - 添加完整类型提示
    - 确保方法不超过 50 行
    - _Requirements: 1.1-1.6, 2.1-2.6, 3.1, 3.3, 4.1-4.5_
  - [ ]* 5.2 编写转换器单元测试
    - 测试转换功能
    - 测试合并单元格处理
    - _Requirements: 1.3, 1.4_

- [x] 6. HTML 切分器重构
  - [x] 6.1 重构 html2chunk.py 为 chunker.py
    - 将函数封装为 HtmlChunker 类
    - 替换 print 为 loguru logger
    - 使用 ChunkConfig 和 ChunkResult 数据模型
    - 添加完整类型提示
    - 确保方法不超过 50 行
    - _Requirements: 1.1-1.6, 2.1-2.6, 3.1, 3.3, 4.1-4.5_
  - [ ]* 6.2 编写切分器单元测试
    - 测试按 token 切分
    - 测试按行数切分
    - _Requirements: 3.3_

- [x] 7. Checkpoint - 核心模块验证
  - 确保所有测试通过，如有问题请询问用户

- [x] 8. 业务处理器重构
  - [x] 8.1 重构 handlers.py
    - 创建 ProcessingState dataclass 替代全局变量
    - 创建 ExcelProcessHandler 类
    - 替换 print 为 loguru logger
    - 替换 os.path 为 pathlib
    - 添加完整类型提示
    - _Requirements: 1.1-1.6, 2.1-2.6, 3.1, 4.1-4.5, 10.1-10.3_
  - [ ]* 8.2 编写处理器属性测试
    - **Property 6: No Global Mutable State**
    - **Validates: Requirements 10.1**

- [x] 9. Pipeline 重构
  - [x] 9.1 重构 pipeline.py
    - 将函数封装为 ConversionPipeline 类
    - 替换 print 为 loguru logger
    - 使用新的数据模型
    - 添加完整类型提示
    - _Requirements: 1.1-1.6, 2.1-2.6, 3.1, 4.1-4.5_

- [x] 10. UI 模块更新
  - [x] 10.1 更新 ui.py
    - 使用重构后的 ExcelProcessHandler
    - 使用枚举替代魔术字符串
    - _Requirements: 5.4_
  - [x] 10.2 更新 main.py
    - 配置 loguru sink
    - 加载 Settings 配置
    - _Requirements: 2.6, 6.1_

- [x] 11. 模块导出更新
  - [x] 11.1 更新 __init__.py 文件
    - 更新 src/core/excel2html/__init__.py
    - 更新 src/core/__init__.py
    - 更新 src/app/__init__.py
    - _Requirements: 3.1_

- [x] 12. Checkpoint - 集成验证
  - 确保所有测试通过，如有问题请询问用户

- [x] 13. 代码风格属性测试
  - [ ]* 13.1 编写路径操作属性测试
    - **Property 1: Path Operations Use Pathlib**
    - **Validates: Requirements 1.1, 1.2, 1.5, 1.6**
  - [ ]* 13.2 编写日志使用属性测试
    - **Property 2: Logging Uses Loguru**
    - **Validates: Requirements 2.1, 2.2, 2.3**
  - [ ]* 13.3 编写类型语法属性测试
    - **Property 3: Modern Type Syntax**
    - **Validates: Requirements 4.1, 4.2, 4.3**
  - [ ]* 13.4 编写类型提示完整性属性测试
    - **Property 4: Complete Type Hints**
    - **Validates: Requirements 4.5**
  - [ ]* 13.5 编写字符串格式化属性测试
    - **Property 5: F-String Formatting**
    - **Validates: Requirements 5.1, 5.2**
  - [ ]* 13.6 编写方法长度属性测试
    - **Property 7: Method Length Limit**
    - **Validates: Requirements 3.3**
  - [ ]* 13.7 编写配置集中化属性测试
    - **Property 10: Centralized Config**
    - **Validates: Requirements 6.2**

- [x] 14. 清理和文档
  - [x] 14.1 删除旧文件
    - 删除 excel2html_openpyxl.py（已被 converter.py 替代）
    - 删除 excel2html_unstructed.py（未使用）
    - _Requirements: 3.1_
  - [x] 14.2 运行 Ruff 格式化
    - 执行 ruff check --fix
    - 执行 ruff format
    - _Requirements: 8.1, 8.2_

- [x] 15. Final Checkpoint - 完整验证
  - 运行所有测试确保通过
  - 运行 ruff check 确保无 lint 错误
  - 如有问题请询问用户

## Notes

- 任务标记 `*` 的为可选测试任务，可跳过以加快 MVP 进度
- 每个任务都引用了具体的需求条款以确保可追溯性
- Checkpoint 任务用于阶段性验证
- 属性测试验证代码风格的正确性属性
