# 自动化执行记录

## 2026-10-30 首次执行：数组公式（CSE / spill）系列改造

用户给了 5 条命令，全部完成并通过 Excel COM 冒烟测试（`tools/PublishingSmokeTest`，0 失败）。

1. **代入法归零考虑数组公式** —— `RootsOfEquations.successive` 增加拦截：被减数/减数任一方属于数组公式（旧式 CSE 或新式溢出）即提示并放弃修改。
2. **更新默认值考虑新旧式数组公式** —— `Publishing.run` 改为数组感知：多格区域整块读写，兜底值用同形数组常数；新式溢出只写锚点，旧式 CSE 走 `FormulaArray`。
3. **新增「转溢出公式」命令** —— `ArrayFormula.convertCseToSpill` + Ribbon `btnCseToSpill`，旧式 CSE 一键转新式 spill，失败自动回滚。
4. **另存为脱公式独立程序** —— 新建解决方案 `C:\Application Data\GitHub\xp44mm\FormulaStripper`（F# + WPF + OpenXml），改写 xlsx 去掉 dna6 自定义函数，输出「原名（崔胜利）.xlsx」静默覆盖；Ribbon 按钮改为启动它。
5. **包裹自定义函数扩到全工作簿** —— `Publishing.wrapWorkbook` 扫描所有含自定义函数的公式，多格数组兜底值用数组常数。

### 关键背景（后续复用，勿重新摸索）
- Excel COM 上：旧式 CSE `HasArray=true/HasSpill=false`，新式溢出 `HasArray=false/HasSpill=true`、`SpillingToRange` 仅锚点有、`CurrentArray` 抛异常。
- 向溢出区非锚点写入**不报错但静默破坏溢出**，所有写入前必须判「是否为主格」。
- xlsx 中 XLL 自定义函数写作 `_xll.函数名(...)`。
- `BuiltinFunctions.fs` 被两个解决方案共享（`<Compile Include>` 链接），改函数名清单只改一处。
- VSTO C# 工程本机无法用 `dotnet build` 编译（缺 VSTO targets），改动只能靠代码审查；F# 工程可正常构建。

### 产物
`ExcelNumericalMethods/ArrayFormula.fs`、`BuiltinFunctions.fs`（新）、`Publishing.fs`/`RootsOfEquations.fs`/`Traversal.fs`（改）、`ExcelVSTO/Ribbon1.cs`+`Ribbon1.Designer.cs`（改）、`tools/PublishingSmokeTest`（扩测）、`../FormulaStripper/`（新解决方案）。
