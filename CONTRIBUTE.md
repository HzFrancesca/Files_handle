# Python 3.12 项目开发规范

## 1. 核心原则 (Core Principles)

###  路径处理 

- **强制使用 `pathlib`**：严禁使用 `os.path` 或字符串拼接路径。
- 利用 `/` 运算符构建路径，使用 `.read_text()` / `.write_text()` 处理文件 IO。
  - ❌ `os.path.join(dir, "file.txt")`
  - ✅ `Path(dir) / "file.txt"`

###  日志管理 

- **强制使用 `loguru`**：弃用 Python 标准库 `logging`。
- 利用 `loguru` 的开箱即用特性（颜色高亮、自动序列化、异常回溯）。
- 禁止在代码中编写复杂的 Logger 配置 Boilerplate，统一在入口文件 (`main.py` / `__init__.py`) 配置 Sink。
  - ❌ `logging.getLogger(__name__).info(...)`
  - ✅ `from loguru import logger` -> `logger.info(...)`

###  架构与 OOP

- **面向对象优先**：核心业务逻辑必须封装在 `class` 中，禁止“脚本式”全局函数编程。
- **数据模型**：纯数据对象强制使用 `@dataclass`。
- **单一职责 (SRP)**：
  - 单个方法/函数**严禁超过 50 行**。
  - 复杂逻辑必须拆解为私有辅助方法 (`_helper_method`)。

## 2. 类型系统

利用 Python 3.12 新特性进行静态类型检查：

- **原生泛型 (PEP 585)**：使用 `list`, `dict`, `tuple`, `set`。
  -  禁用 `typing.List`, `typing.Dict` 等旧式写法。
- **并集语法 (PEP 604)**：
  - ✅ `int | str` (替代 `Union[int, str]`)
  - ✅ `User | None` (替代 `Optional[User]`)
- **类型别名 (Python 3.12)**：
  - ✅ 使用 `type` 关键字定义别名：`type UserID = int | str`

## 3. 语法规范 

- **格式化**：仅使用 **f-string** (`f"{var}"`)，禁止 `%` 或 `.format()`。
- **控制流**：在状态机或解析逻辑中，优先使用 **`match/case`** (Python 3.10+)。
- **常量**：魔术字符串必须封装为 `Enum` 或 `StrEnum`。

## 4. 进阶工具与最佳实践 

- **包与环境管理 (Package Management via uv)**：
  - **强制使用 `uv`** 管理项目依赖与虚拟环境。
  - 弃用 `pip`, `poetry`, `pipenv`。
  - 使用 `uv pip install`或者 `uv add` 安装项目依赖
  - 利用 `uv sync` 生成 `uv.lock` 文件，确保开发与生产环境严格一致。
- **数据校验 (Pydantic v2)**：
  - 外部输入（API Payload, JSON 配置）**强制使用 `Pydantic v2`** 进行运行时校验。
  - 内部数据流转使用 `Dataclass`。
- **配置管理 (Configuration)**：
  - 严禁在业务代码中散落 `os.getenv`。
  - 使用 `pydantic-settings` 统一管理环境变量与配置。
- **代码质量 (Tooling)**：
  - 统一使用 **`Ruff`** 替代 Black/Flake8/Isort，配置在 `pyproject.toml` 中。
- **并发编程 (AsyncIO)**：
  - 使用 Python 3.11+ 的 **`asyncio.TaskGroup`** 管理并发任务，弃用裸露的 `asyncio.gather` 以确保异常安全性。

## 5. 测试与质量保证

- **测试框架 (Pytest)**:
  - **强制使用 `pytest`**，弃用 `unittest`。
  - 异步代码必须配合 `pytest-asyncio` 测试。
  - 使用 `conftest.py` 管理共享 Fixtures。
- **Git Hooks (Pre-commit)**:
  - 项目必须配置 `.pre-commit-config.yaml`。
  - 提交代码前强制自动运行 `ruff check`, `ruff format` 和 `uv lock --check`。

