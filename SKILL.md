---
name: excel-table-export
description: Unity 项目的 Excel 配置表导表全流程（基于 ExportTablesTool 导出 + LoadData 运行时加载方案）：预校验 Excel、经 Unity 菜单 Tools/Export Tables 导出 JSON 与生成 GameDataTables.cs、核对导出产物、排查运行时加载失败。当目标 Unity 项目已内置该导表工具（命名空间 KTXGame.ExportTables）并需改动或新增配置表、同步导出产物、或排查表加载问题时使用。不适用于与该工具无关的通用表格数据分析。
---

# Excel 配置表导表（ExportTablesTool）

把策划改动的 Excel 配置表安全地变成游戏运行数据。核心链路：**Excel 预校验 → Unity 导出 → 产物核对 → 提交四件套**。运行时读的是导出的 JSON，不是 Excel 本身。

## 适用前提与依赖

本 skill 面向**已内置这套导表方案的 Unity 项目**，方案由两个互相独立的文件组成：

| 角色 | 文件 | 关键标识 |
|---|---|---|
| 导出工具 | `Assets/Scripts/Tools/ExportTablesTool.cs` | 命名空间 `KTXGame.ExportTables`，内含路径类 `ExportTablesToolPaths`，菜单 `Tools/Export Tables` |
| 运行时加载器 | `Assets/Scripts/Utils/LoadData.cs` | 无命名空间；API `TableExists(string)` / `LoadJSONData<T>(string)` / `ClearCacheForEditor()` |

**开工前先确认工具在场**：在目标项目里检索 `class ExportTablesToolPaths` 或 `namespace KTXGame.ExportTables`。找不到即说明该项目未使用本方案，本 skill 不适用，不要凭本文档去假定一个不存在的工具或路径。

两个文件均零业务耦合（只依赖 `System.*` / `UnityEditor` / `UnityEngine`），可整体移植。**Step 4 描述的是 `LoadData` 的契约，属该方案的下游配套，不是导出工具本身生成的内容**——若目标项目只有导出工具而没有 `LoadData`，Step 4 与 Step 5 的相关部分不成立。

## 路径与配置：以代码为准

下表是 `ExportTablesToolPaths` 中三个常量的**默认值，不是固定值**。项目可以改这三个 `const`，因此**接手新项目或换项目时先读这个类**：

| 常量 | 默认值 | 角色 |
|---|---|---|
| `ExcelRootRelativePath` | `Assets/Resources/Docs/Tables` | Excel 源目录 |
| `OutputRelativePath` | `Assets/Resources/GameData/ExportedTables` | 导出根目录（JSON 产物） |
| `GeneratedCodeRelativePath` | `Assets/Scripts/GameData` | 生成代码目录 |

实际值与默认值不一致时**以代码为准**，并将本文档后续所有路径理解为上表对应的**角色目录**，而不是那串字面字符串。Excel 源只认 `.xlsx`，不支持 `.xls` / `.csv`。

## 关键约束与红线

- **禁止手改 `GameDataTables.cs`**（下次导出会被覆盖）
- 业务代码只能经 `LoadData.LoadJSONData<T>(tableName)` 读表，不得直接读 JSON 或绕过 manifest
- Excel 字段名一旦投入使用不要随意改名

## 输出路径一致性风险（重要）

导出工具的 `OutputRelativePath` **允许被项目改代码覆盖**，而 `LoadData` 侧有**两处独立的硬编码常量**，改工具侧不会带动它们：

| 位置 | 值 | 用途 |
|---|---|---|
| `LoadData.cs` 内 `ExportedTablesResourcesPath` | `"GameData/ExportedTables"` | **Resources 相对路径**，经 `Resources.Load<TextAsset>` 读取 |
| `LoadData.cs` 内缓存失效 Postprocessor 监听前缀 | `"Assets/Resources/GameData/ExportedTables/"` | 判定资源变更以失效缓存 |

注意两者**格式都不同**：工具侧写 `Assets/Resources/...`，加载器侧写 `Resources` 之后的相对段。后果是——项目一旦改了导出路径，**导出会成功、运行却读不到表，且重导后缓存不刷新**，且没有任何报错指向路径。

**换项目时必须自查这三点：**

1. 工具 `OutputRelativePath` 是否仍为 `Assets/Resources/GameData/ExportedTables`
2. 该目录是否位于某个 `Resources/` 下，且其后的子路径恰为 `GameData/ExportedTables`
3. 若项目改过输出路径，须同步修改 `LoadData` 的常量与 Postprocessor 前缀（属代码改动，需人工确认后再动）

## Sheet 结构约定

每个可导出的 Sheet 至少 4 行，前 3 行是表头：

| 行 | 含义 | 要求 |
|---|---|---|
| 第 1 行 | 字段说明 | 给人看，写入 JSON 的 columns.description |
| 第 2 行 | 字段类型 | 只支持 int / int32 / long / int64 / float / single / double / bool / boolean / string（int32→int、int64→long、single→float、boolean→bool 自动归一化）；写 List\<int\>、enum、vector3 等会直接报错 |
| 第 3 行 | 字段名 | 必须匹配 `^[A-Za-z_][A-Za-z0-9_]*$`（字母或下划线开头，只用字母数字下划线）。作为 JSON 键名和 C# 字段名；不满足时工具**不报错**，只会静默剥离字符，导致两者不一致 |
| 第 4 行起 | 数据行 | 空行跳过；整表无有效数据行则该 Sheet 被跳过 |

每个有效列的三行表头必须齐全，缺任何一项导出报错。

Sheet 命名规则：默认 Sheet 名 = 表名 = 类型名（生成 `XXXRecord` / `XXXTable`）。用 `结构名.表名`（第一个 `.` 拆分）可以让同结构多张表复用一套类型；此时同一结构名下所有 Sheet 的**字段名、类型、顺序必须完全一致**（字段说明可不同），最终表名必须全局唯一。

## 执行流程

### Step 0 — 先确认导出配置

按上文「路径与配置」一节，读出目标项目 `ExportTablesTool.cs` 中 `ExportTablesToolPaths` 三个常量的**实际值**。**不要假定路径**：实际值与默认值不一致时以代码为准，并将后续所有路径理解为其对应的角色目录。

### Step 1 — Excel 预校验（先跑，别急着进 Unity）

运行随附脚本做静态检查，把结构错误挡在进 Unity 之前：

```bash
<python> <此 skill 目录>/scripts/validate_tables.py <Excel 源目录 | 单个 .xlsx>
```

- `<python>` 指任意 Python 3（需已安装 `openpyxl`），不要绑定某个特定解释器或安装位置
- 路径参数可以是 Excel 源目录（递归）或单个 `.xlsx`
- 退出码 0 = 无错误；有 ERROR 必须先改 Excel 再继续，WARNING 需人工判断（多为非阻断项）
- 脚本只读，不会改任何文件

校验脚本有意比导表工具**更严**：工具对非法字段名是静默剥离字符、不报错（问题会推迟到运行时才炸出来），脚本直接判 ERROR。**脚本报 ERROR 而工具能跑通时，优先修 Excel，不要放宽校验规则。**

### Step 2 — Unity 导出

1. 优先使用宿主环境提供的 Unity 编辑器自动化能力触发菜单 `Tools/Export Tables`；先确认 Editor 未在编译 / Domain Reload
2. 没有此类自动化能力时，提示用户手动执行菜单导出，不要用文本方式伪造导出产物
3. 导出范围按需选择：只改了一张表就只选该文件，不必全量导出
4. 但**范围导出只对范围内的文件做清理**——若本次涉及删除表、或清空了某张 Sheet 的全部数据行，导出范围必须覆盖到那个文件（或直接全量），否则旧 JSON 与 manifest 条目会残留。见 Step 3 的说明

### Step 3 — 核对导出产物

导出成功后核对三处（对照 Unity Console 的解析日志）：

- **导出根目录**下对应 JSON 已更新。文件名为 `Excel文件名_最终表名.json`（去掉 `.xlsx` 扩展名，用 `.` 后面的**最终表名**，不是完整 Sheet 名）。命名示例：
  - `Demo.xlsx` 的 Sheet `S1`（无点号，表名即 Sheet 名）→ `Demo_S1.json`
  - `Demo.xlsx` 的 Sheet `Row.Item01` → `Demo_Item01.json`（取点号后的最终表名）
  - Excel 若放在 Excel 源目录的子目录里，导出路径会**镜像同名子目录**，不是一律平铺在根下
- **导出根目录**下 `tables_manifest.json` 索引已更新。manifest 是**合并写入**的：每次导出读取旧 manifest、并入本次条目，并把范围内已消失的表连同 JSON 一起清掉
- **生成代码目录**下 `GameDataTables.cs` 中存在对应的 `XXXTable` / `XXXRecord`（如预期新生成）。代码按 **schemaName** 去重生成，同结构的多张 Sheet 只产出一套类型，但每张 Sheet 各自的 JSON 都会导出

> **manifest 的合并与 stale 清理都受导出范围限制**（`MergeManifest` / `FindStaleEntries` 内部都有 `IsWithinScope` 判断）：范围内的一切"从零重建"，范围外的历史条目被**原样保留**。所以只导某一个文件时，别处已经消失的表不会被清理——stale JSON 会残留，且因为 manifest 里还有该条目，运行时 `LoadJSONData` 依然能把它加载出来。**要清理必须对包含该文件的目录（或全量）导一次。**

### Step 4 — 运行时加载要点（涉及代码改动时）

本节由 `LoadData` 提供，属该方案的下游配套。

```csharp
ItemTable table = LoadData.LoadJSONData<ItemTable>("Item");
```

- 泛型类型用 **schemaName** 生成的 `XXXTable`；表名参数用**最终表名**（普通 Sheet = Sheet 名，`结构名.表名` = `.` 后面的部分）。两者不要混：`Row.Item01` 对应 `LoadJSONData<RowTable>("Item01")`
- `LoadData.TableExists("Item")` 可做无日志的存在性预检，适合在改代码前先确认表名拼写
- 缓存键是 `泛型类型全名 + ":" + 表名`，同一张表用不同泛型加载会各存一份、互不覆盖
- Editor 下重导后缓存自动失效（`AssetPostprocessor` 监听导出根目录的增删改移）；仍可用 `LoadData.ClearCacheForEditor()` 手动清

**返回 null 之前，`LoadJSONData` 会做一轮硬校验，任一不通过都直接返回 null 并打 `Debug.LogError`：**

1. 目标类型必须有 `string tableName` 字段（public 或 non-public 均可），且其值必须与被请求的表名一致（忽略大小写）
2. 目标类型必须有 `data` 字段，且是数组或单泛型集合（`T[]` / `List<T>`）
3. 从 JSON 文本正则抽出 `columns` 里的 `{"name":…,"type":…}`，**逐列**在行类型上查找同名字段，并校验字段类型与 schema 类型**精确相同**
4. 校验是**单向**的：行类型多出的字段不报错；schema 里的字段缺失或类型不符才报错

所以"自己写个差不多的 XXXRecord"必然失败——字段名与类型必须与导出 schema 逐字一致。报错信息会直接指出是哪个字段缺失或类型不符，按提示改即可。

返回 null 时按顺序查：① 是否已导出（导出根目录下有无 JSON）→ ② manifest 里有没有该表名 → ③ 上面的硬校验是否通过 → ④ 路径一致性（见上文「输出路径一致性风险」）

### Step 5 — 提交清单

导出涉及的产物必须同步提交，缺一不可：

1. 改动的 `.xlsx`（及其 `.meta`）
2. 导出根目录下更新的 JSON（含新增/删除，`.meta` 配对）
3. `tables_manifest.json`
4. `GameDataTables.cs`

只产出 commit message 供用户确认，**不自动 commit**。不提交 Office 锁文件（`~$*.xlsx`）和临时文件。

## 排错速查

| 现象 | 常见原因 | 修复方向 |
|---|---|---|
| 导出报表头缺失 | 第 1/2/3 行某列缺内容（只填了两项也报错） | 补齐该列说明、类型、字段名 |
| 导出报表头缺失，且该列在 Excel 里看着有值 | 表头前 3 行存在合并单元格，合并后只有左上角单元格有值 | 取消表头行的合并单元格，导表工具按逐列读取表头 |
| 导出报类型不支持 | 用了 List\<int\>、enum、decimal、vector3 等 | 改为基础标量类型；确需复杂类型先扩展导表工具本身 |
| 导出报重复表名 | 多个 Sheet 最终表名相同（含 `结构名.表名` 拆分后的撞名） | 改 Sheet 名保证最终表名唯一 |
| 导出报同结构字段不一致 | 同 schemaName 的 Sheet 字段名/类型/顺序不同 | 对齐结构，或改用不同结构名 |
| 某 Sheet 没导出来 | 少于 4 行 / 无有效表头 / 无数据行 | 查工具窗口提示和 Console 日志 |
| 改了 Excel 但运行没变化 | 没重新导出 / 导出范围选错 / 资源未刷新 | 重新执行 Step 2-3 |
| 删了表但旧数据仍能加载到 | 只删了 Excel 没重新导出，或范围导出没覆盖到该表（scope 外条目原样保留） | 对包含该文件的目录或全量重跑一次导出，让 manifest 同步清理 |
| Excel 里已清空的 Sheet，运行时却还能读到旧数据 | 该 Sheet 之前导出过、之后被清空，而同范围的导出没覆盖到它 | 同上。核对手段：比对 `tables_manifest.json` 里该表的 `rowCount` 与 Excel 实际数据行数 |
| 导出成功、路径也对，运行时仍读不到表 | 项目改过 `OutputRelativePath`，但 `LoadData` 的硬编码常量没同步 | 见上文「输出路径一致性风险」的自查三点 |
| 运行时加载返回 null | 未导出、表名错、泛型类型不匹配 | 按 Step 4 的排查顺序处理 |
| 运行时报 `row type 'XXXRecord' is missing field 'yyy'` | 该列字段名以数字开头或含 `-`/空格等非法字符。工具**不报错**，只静默剥离字符生成 C# 字段名（`1stItem`→`_1stItem`、`my-field`→`myfield`），而 JSON 键保持原名 | 字段名改为以字母或下划线开头，且只用 `[A-Za-z0-9_]`。这类问题校验脚本会提前判 ERROR |
| 运行时报 JSON `tableName` 与请求表名不匹配 | 泛型类型用错（用了另一个结构的 `XXXTable`） | 换成与 Sheet 的 schemaName 对应的类型 |
| 生成的类名不是预期的（如多出 `_` 前缀或字母被吃掉） | 表名含数字开头、`-`、空格等，`SanitizeTypeIdentifier` 会剥离分隔符并把分隔后的首字母大写（`001_Node`→`_001Node`） | 表名只用字母、数字、下划线，且不以数字开头 |
