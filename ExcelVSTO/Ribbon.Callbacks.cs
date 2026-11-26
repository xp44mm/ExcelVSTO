using System;
using System.Linq;
using System.Windows;
using ExcelNumericalMethods;
using ExcelWPF;

using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelVSTO
{
    public partial class Ribbon
    {
        /// <summary>
        /// 复制当前工作簿信息：C# 只做链接，核心逻辑在 F# 的 WorkbookInfo.copyCurrent 中。
        /// </summary>
        public void BtnCopyActiveWorkbookInfo_Click(Office.IRibbonControl control)
        {
            ExcelNumericalMethods.WorkbookInfo.copyCurrent(Globals.ThisAddIn.Application);
        }

        public void BtnSelectInUsedRange_Click(Office.IRibbonControl control)
        {
            var sel = Globals.ThisAddIn.Application.Selection as Range;
            var ws = sel.Worksheet;


            try
            {
                var usedRange = ws.UsedRange.get_Address();
                var selectedAddress = sel.get_Address();
                var addr = Prune.prune(usedRange, selectedAddress);
                ws.Range[addr].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void BtnSelectArray_Click(Office.IRibbonControl control)
        {
            //选中一个单元格，如果这个单元格没有数组，选择不变，否则选择整个数组。
            var cell = Globals.ThisAddIn.Application.ActiveCell;
            try
            {
                cell.CurrentArray.Select();
            }
            catch
            {
                cell.Select();
            }
        }

        public void BtnMergeCells_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Range;
            var opt = app.DisplayAlerts;
            try
            {
                app.DisplayAlerts = false;
                ExcelNumericalMethods.NumericalMethods.merge(sel);
            }
            finally
            {
                app.DisplayAlerts = opt;
            }
        }

        /// <summary>
        /// 复制选择：把当前工作簿文件名与当前选择的区域地址复制到剪贴板。
        /// 格式与核心逻辑在 F# 的 SelectionCopy.copySelection 中实现。
        /// </summary>
        public void BtnCopySelection_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            var sheet = app.ActiveSheet as Worksheet;
            var sel = app.Selection as Range;

            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            if (sel == null)
            {
                MessageBox.Show("当前没有选择单元格区域！");
                return;
            }

            if (sheet == null)
            {
                MessageBox.Show("当前没有活动工作表！");
                return;
            }

            try
            {
                ExcelNumericalMethods.SelectionCopy.copySelection(sel, wb.FullName, sheet.Name);
                app.StatusBar = "当前选择已经放入剪贴板";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 复制当前单元格：C# 只做链接，核心逻辑在 F# 的 ActiveCellCopy.copyActiveCell 中。
        /// 没有可复制内容时不改动剪贴板，在状态栏提示。
        /// </summary>
        public void BtnCopyActiveCell_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var cell = app.ActiveCell as Range;
            if (cell == null)
            {
                app.StatusBar = "执行了命令但没有内容可复制";
                return;
            }

            try
            {
                var text = ExcelNumericalMethods.ActiveCellCopy.copyActiveCell(cell);
                app.StatusBar = string.IsNullOrEmpty(text)
                    ? "执行了命令但没有内容可复制"
                    : "当前单元格的值已放入剪贴板";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void BtnUnroll_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.fillColumns(sel);
        }

        public void BtnRollup_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.tidyColumns(sel);
        }

        public void BtnInsertBlank_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.split(sel).Select();
        }

        public void BtnRemoveBlank_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.removeBlank(sel).Select();
        }

        /// <summary>
        /// 竖框线：清除选择的全部框线后，为每个单元格添加左右竖框线，并添加最顶行上框线与最底行下框线。
        /// 核心逻辑在 F# 的 TableBorder 中实现。
        /// </summary>
        public void BtnVerticalBorder_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Range;
            if (sel == null)
            {
                MessageBox.Show("当前没有选择单元格区域！");
                return;
            }

            try
            {
                ExcelNumericalMethods.TableBorder.apply(sel);
                app.StatusBar = "已应用竖框线";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void BtnAlternateRows_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.alternateColor(sel);
        }

        public void BtnIncrease_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.plusCell(1, sel);
        }

        public void BtnDecrease_Click(Office.IRibbonControl control)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            NumericalMethods.plusCell(-1, sel);
        }

        public void BtnSuccessive_Click(Office.IRibbonControl control)
        {
            var goalCell = Globals.ThisAddIn.Application.ActiveCell;
            try
            {
                RootsOfEquations.successive(goalCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void BtnBisect_Click(Office.IRibbonControl control)
        {
            var goalCell = Globals.ThisAddIn.Application.ActiveCell;
            try
            {
                RootsOfEquations.bisect(goalCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        /// <summary>
        /// 工作表引用：把本工作表公式对外部地址（其他工作表 / 其他工作簿）的引用，
        /// 集中写到 UsedRange 下方空一行处的中转区：每个外部区域写一条区域公式，
        /// 由动态数组向右下溢出成原形状；原公式改为引用同形状的中转区域，
        /// 使复杂公式内部不再嵌套外部地址。
        /// 整行整列、3D 引用、超大区域不中转，在表尾逐行写提示；
        /// 多格数组公式（Ctrl+Shift+Enter 数组或动态数组溢出）按数组单元整体收集与改写，
        /// 不会因为「属数组」被跳过而漏掉其中的外部引用。
        /// 本表完全没有可中转的外部地址时，也在表尾写一行说明。
        /// 每次都从 UsedRange 下方重新追加，不识别、不复用上一次的中转区。
        /// 核心逻辑在 F# 的 WorksheetRefs.relay 中实现。
        /// </summary>
        public void btn_referencesOfWorksheet_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var ws = app.ActiveSheet as Worksheet;

            if (ws == null)
            {
                MessageBox.Show("当前没有活动工作表！");
                return;
            }

            var oldCalculation = app.Calculation;
            try
            {
                app.Calculation = XlCalculation.xlCalculationManual;

                var result = ExcelNumericalMethods.WorksheetRefs.relay(ws);

                var parts = new System.Collections.Generic.List<string>();
                if (result.NewRelayCount > 0)
                {
                    parts.Add($"建立中转区 A{result.RelayStartRow} 起 {result.RelayRows} 行 × {result.RelayCols} 列（{result.NewRelayCount} 个单元格，区域公式动态数组溢出）");
                }
                if (result.RewrittenCells > 0)
                {
                    parts.Add($"改写公式 {result.RewrittenCells} 个单元格");
                }
                if (result.HintCount > 0)
                {
                    parts.Add($"未中转 {result.SkippedRefs} 处，提示写在 A{result.HintStartRow} 起 {result.HintCount} 行");
                }
                if (result.NoRefNoticeRow > 0)
                {
                    parts.Add($"本表没有可中转的外部地址，已在 A{result.NoRefNoticeRow} 写明");
                }
                if (result.NewRelayCount == 0 && result.RewrittenCells == 0 && result.HintCount == 0
                    && result.NoRefNoticeRow == 0)
                {
                    parts.Add("没有需要中转的外部地址");
                }
                if (result.HintCount > 0)
                {
                    var sample = string.Join("、", result.SkippedTexts.Take(3));
                    var more = result.SkippedTexts.Length > 3 ? " 等" : "";
                    parts.Add($"未中转：{sample}{more}");
                }
                if (result.ArrayUnits.Length > 0)
                {
                    parts.Add($"多格数组公式 {result.ArrayUnits.Length} 处按数组整体改写（{string.Join("、", result.ArrayUnits.Take(3))}{(result.ArrayUnits.Length > 3 ? " 等" : "")}）");
                }
                if (result.FailedCells.Length > 0)
                {
                    parts.Add($"写入失败 {result.FailedCells.Length} 处（{string.Join("、", result.FailedCells)}）");
                }
                app.StatusBar = "工作表引用：" + string.Join("；", parts) + "。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                app.Calculation = oldCalculation;
                app.Calculate();
            }
        }

        /// <summary>
        /// 删除中转单元格：当前单元格必须是一元表达式（内联到其他公式后不必加括号、不改变语义）。
        /// 扫描当前工作簿全部公式，把引用本单元格的地址替换为本单元格的公式，然后清空本单元格。
        /// 无法内联时（区域引用包含本单元格、公式带顶层运算符、多格数组、合并单元格等）
        /// 弹出原因列表且不做任何改动。核心逻辑在 F# 的 RelayDelete.delete 中实现。
        /// </summary>
        public void BtnDeleteRelayCell_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var cell = app.ActiveCell as Range;
            if (cell == null)
            {
                MessageBox.Show("当前没有活动单元格！");
                return;
            }

            try
            {
                var result = ExcelNumericalMethods.RelayDelete.delete(cell);
                if (result.Deleted)
                {
                    app.StatusBar = result.Message;
                }
                else
                {
                    MessageBox.Show(
                        "未删除 " + result.Address + "：" + Environment.NewLine
                        + string.Join(Environment.NewLine, result.Blockers));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// 标记跨表引用：扫描当前工作簿全部公式，把被本工作表以外的单元格引用的单元格背景色涂成洋红。
        /// 只处理工作簿内部的「工作表!地址」引用（含 3D 引用），外部工作簿引用与名称引用不处理。
        /// 核心逻辑在 F# 的 CrossSheetRefs.mark 中实现。
        /// </summary>
        public void BtnMarkCrossSheetRefs_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            try
            {
                var result = ExcelNumericalMethods.CrossSheetRefs.mark(wb);
                var parts = new System.Collections.Generic.List<string>();
                if (result.MarkedAreas == 0)
                {
                    parts.Add("没有被其他工作表引用的单元格");
                }
                else
                {
                    parts.Add($"已标记引用区域 {result.MarkedAreas} 个");
                    parts.Add($"涉及工作表 {result.MarkedSheets.Length} 张（{string.Join("、", result.MarkedSheets)}）");
                }
                if (result.MissingSheets.Length > 0)
                {
                    parts.Add($"公式引用了不存在的工作表 {result.MissingSheets.Length} 个（{string.Join("、", result.MissingSheets)}）");
                }
                if (result.ProtectedSheets.Length > 0)
                {
                    parts.Add($"受保护无法标色的工作表 {result.ProtectedSheets.Length} 张（{string.Join("、", result.ProtectedSheets)}）");
                }
                app.StatusBar = "标记跨表引用：" + string.Join("；", parts) + "。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 恢复默认显示：C# 只做链接，核心逻辑在 F# 的 DisplayDefaults.apply 中。
        /// </summary>
        public void BtnRestoreDisplayDefaults_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var ws = app.ActiveSheet as Worksheet;
            if (ws == null)
            {
                MessageBox.Show("当前没有活动工作表！");
                return;
            }

            try
            {
                var result = ExcelNumericalMethods.DisplayDefaults.apply(app, ws);
                app.StatusBar = $"已恢复默认显示：已将 {result.Item1} 行的行高、{result.Item2} 列的列宽恢复为默认值。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "恢复默认显示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 工作簿另存为 SQLite：C# 只做链接，核心逻辑在 F# 的 SqliteCommands.saveToSqlite 中。
        /// </summary>
        public void btnSaveToSqlite_Click(Office.IRibbonControl control)
        {
            ExcelNumericalMethods.SqliteCommands.saveToSqlite(Globals.ThisAddIn.Application);
        }

        /// <summary>
        /// 从 SQLite 导入工作表：C# 只做链接，核心逻辑在 F# 的 SqliteCommands.createFromSqlite 中。
        /// </summary>
        public void btnCreateFromSqlite_Click(Office.IRibbonControl control)
        {
            ExcelNumericalMethods.SqliteCommands.createFromSqlite(Globals.ThisAddIn.Application);
        }

        /// <summary>
        /// 更新默认值：扫描当前工作簿全部工作表，不限函数名，处理所有公式为
        /// =IFERROR(函数(...), 常量) 的单元格：将兜底常量替换为该单元格最新真值，并以绿色标记。
        /// 直接在当前工作簿上位修改，不生成副本、不弹输入框与结果确认框（结果用状态栏提示）。
        /// 核心逻辑（扫描、过滤、改写、标色、统计）在 F# 的 Publishing.updateDefaults 中实现。
        /// </summary>
        public void BtnUpdateDefaults_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            try
            {
                var result = Publishing.updateDefaults(wb);
                if (result.ProcessedCount == 0
                         && result.ErrorCells.Length == 0
                         && result.UnconformCells.Length == 0
                         && result.ArrayFormulaCells.Length == 0
                         && result.ProtectedSheets.Length == 0)
                {
                    app.StatusBar = "未找到 IFERROR(函数(...), 常量) 结构的公式。";
                }
                else
                {
                    var parts = new System.Collections.Generic.List<string> { $"已更新 {result.ProcessedCount} 个单元格" };
                    if (result.ErrorCells.Length > 0)
                    {
                        parts.Add($"跳过错误值 {result.ErrorCells.Length} 个");
                    }
                    if (result.UnconformCells.Length > 0)
                    {
                        parts.Add($"跳过格式不符/非常量兜底 {result.UnconformCells.Length} 个");
                    }
                    if (result.ArrayFormulaCells.Length > 0)
                    {
                        parts.Add($"数组公式写入失败 {result.ArrayFormulaCells.Length} 个");
                    }
                    if (result.ProtectedSheets.Length > 0)
                    {
                        parts.Add($"跳过受保护工作表 {result.ProtectedSheets.Length} 个");
                    }
                    app.StatusBar = "更新默认值：" + string.Join("；", parts) + "。";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 包裹自定义函数：扫描当前工作簿全部工作表，把所有含 dna6 自定义函数（_xll. 前缀或
        /// 非内置函数名）的公式 =函数(...) 包裹为 =IFERROR(函数(...), 兜底值)。
        /// 兜底值取当前计算结果；多格数组公式（旧式 CSE 区域 / 新式溢出区）取整个区域的值，
        /// 兜底值用同形数组常数（如 {1,2;3,4}）；单格用标量。
        /// 已是 IFERROR、内置函数公式、错误值、受保护工作表不处理。
        /// 结果在状态栏提示，有跳过项时弹窗列出；核心逻辑在 F# 的 Publishing.wrapWorkbook 中实现。
        /// </summary>
        public void BtnWrapFunction_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            try
            {
                var result = Publishing.wrapWorkbook(wb);
                var text = new System.Text.StringBuilder();
                text.AppendLine($"包裹自定义函数：已包裹 {result.WrappedCount} 个单元格。");
                if (result.SkippedCells.Length > 0)
                {
                    text.AppendLine($"未能包裹 {result.SkippedCells.Length} 个：");
                    foreach (var item in result.SkippedCells)
                    {
                        // F# 元组在 .NET Framework 下是 System.Tuple，不能用 var (a, b) 解构
                        text.AppendLine($"  {item.Item1}：{item.Item2}");
                    }
                }
                if (result.ProtectedSheets.Length > 0)
                {
                    text.AppendLine($"跳过受保护工作表 {result.ProtectedSheets.Length} 个：{string.Join("、", result.ProtectedSheets)}");
                }

                if (result.WrappedCount == 0 && result.SkippedCells.Length == 0 && result.ProtectedSheets.Length == 0)
                {
                    app.StatusBar = "未找到含 dna6 自定义函数的公式。";
                }
                else if (result.SkippedCells.Length == 0 && result.ProtectedSheets.Length == 0)
                {
                    app.StatusBar = $"包裹自定义函数：已包裹 {result.WrappedCount} 个单元格。";
                }
                else
                {
                    var detail = new TextWindow("包裹自定义函数结果", text.ToString().TrimEnd());
                    detail.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 旧式 CSE 数组公式一键转换为新式 spill（溢出）数组公式：
        /// 扫描当前工作簿全部工作表，把 Ctrl+Shift+Enter 录入的数组公式整块清除后，
        /// 只在左上角单元格写入同一公式，由 Excel 自动溢出。
        /// 核心逻辑在 F# 的 ArrayFormula.convertCseToSpill 中实现。
        /// </summary>
        public void BtnCseToSpill_Click(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            try
            {
                var result = ArrayFormula.convertCseToSpill(wb);
                var text = new System.Text.StringBuilder();
                text.AppendLine($"旧式 CSE 数组公式已转换为新式溢出公式 {result.ConvertedAddresses.Length} 处。");
                if (result.ConvertedAddresses.Length > 0)
                {
                    text.AppendLine("已转换：");
                    foreach (var addr in result.ConvertedAddresses)
                    {
                        text.AppendLine("  " + addr);
                    }
                }
                if (result.FailedCells.Length > 0)
                {
                    text.AppendLine($"转换失败（已回滚）{result.FailedCells.Length} 处：");
                    foreach (var item in result.FailedCells)
                    {
                        // F# 元组在 .NET Framework 下是 System.Tuple，不能用 var (a, b) 解构
                        text.AppendLine($"  {item.Item1}：{item.Item2}");
                    }
                }
                if (result.ProtectedSheets.Length > 0)
                {
                    text.AppendLine($"跳过受保护工作表 {result.ProtectedSheets.Length} 个：{string.Join("、", result.ProtectedSheets)}");
                }

                if (result.FailedCells.Length == 0 && result.ProtectedSheets.Length == 0)
                {
                    app.StatusBar = $"旧式 CSE 数组公式已转换为新式溢出公式 {result.ConvertedAddresses.Length} 处。";
                }
                else
                {
                    var detail = new TextWindow("转溢出公式结果", text.ToString().TrimEnd());
                    detail.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
