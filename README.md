# Excel Table Export

把策划改动的 Excel 配置表安全地变成 Unity 运行数据 —— 核心链路：
**Excel 预校验 → Unity 导出 → 产物核对 → 提交四件套**。运行时读的是导出的 JSON，不是 Excel 本身。

> An agent skill for Unity projects built on the ExportTablesTool pipeline:
> pre-validate Excel tables, export via the Unity menu, verify the outputs,
> and commit the four artifacts together.

## 适用前提

面向**已内置 ExportTablesTool 导表方案**的 Unity 项目，方案由两个互相独立的文件组成：

| 角色 | 文件 | 关键标识 |
|---|---|---|
| 导出工具 | `Assets/Scripts/Tools/ExportTablesTool.cs` | 命名空间 `KTXGame.ExportTables`，菜单 `Tools/Export Tables` |
| 运行时加载器 | `Assets/Scripts/Utils/LoadData.cs` | API `TableExists` / `LoadJSONData<T>` / `ClearCacheForEditor()` |

开工前先在目标项目里检索 `class ExportTablesToolPaths` 或 `namespace KTXGame.ExportTables`；
找不到即说明该项目未使用本方案，本 skill 不适用。

## 功能

- 进 Unity 之前先用随附脚本对 Excel 做静态预校验，把结构错误挡在导出之前
- 校验有意比导表工具**更严**：工具对非法字段名会静默剥离字符、不报错（问题推迟到运行时才出现），
  脚本对同一问题直接判 ERROR
- 改表、新增表、同步导出产物、排查运行时加载失败的完整工作流见 [`SKILL.md`](SKILL.md)

## 依赖

- **Python 3 + `openpyxl`** —— 仅预校验脚本使用
- 目标 Unity 项目已内置上述导表方案

## 快速开始

```bash
python scripts/validate_tables.py "<Excel 源目录或单个 .xlsx>"
```

路径参数可以是 Excel 源目录（递归）或单个 `.xlsx`。退出码 `0` 表示无错误；
有 ERROR 必须先改 Excel 再继续。脚本只读，不会改任何文件。

## 已知限制

- 只适用于 ExportTablesTool + LoadData 这套导表方案，不适用于通用表格数据分析
- Excel 源只认 `.xlsx`，不支持 `.xls` / `.csv`
- 字段名一旦投入使用不要随意改名；脚本报 ERROR 而工具能跑通时，优先修 Excel，不要放宽校验规则

## 项目结构

```
.
├── SKILL.md                  给 AI agent 的完整工作流入口
├── README.md                 本文件（给人看）
├── LICENSE                   MIT
└── scripts/
    └── validate_tables.py    Excel 静态预校验脚本（只读）
```

安装为 agent skill 时，把整个目录放入 skills 目录即可；`SKILL.md` 是入口。

## License

[MIT](LICENSE) © 2026 zhmge
