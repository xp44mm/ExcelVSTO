using System;
using System.Linq;
using System.Windows;
using ExcelNumericalMethods;
using ExcelWPF;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void BtnSelectInUsedRange_Click(object sender, RibbonControlEventArgs e)
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
        private void BtnSelectArray_Click(object sender, RibbonControlEventArgs e)
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

        private void BtnMergeCells_Click(object sender, RibbonControlEventArgs e)
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
        /// 复制选择：把当前选择的单元格地址与公式字符串复制到剪贴板。
        /// 格式与核心逻辑在 F# 的 SelectionCopy.copySelection 中实现。
        /// </summary>
        private void BtnCopySelection_Click(object sender, RibbonControlEventArgs e)
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

        private void BtnUnroll_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.fillColumns(sel);
        }

        private void BtnRollup_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.tidyColumns(sel);
        }

        private void BtnInsertBlank_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.split(sel).Select();
        }

        private void BtnRemoveBlank_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.removeBlank(sel).Select();
        }

        /// <summary>
        /// 竖框线：清除选择的全部框线后，为每个单元格添加左右竖框线，并添加最顶行上框线与最底行下框线。
        /// 核心逻辑在 F# 的 TableBorder 中实现。
        /// </summary>
        private void BtnVerticalBorder_Click(object sender, RibbonControlEventArgs e)
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

        private void BtnAlternateRows_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.alternateColor(sel);
        }

        private void BtnIncrease_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            ExcelNumericalMethods.NumericalMethods.plusCell(1, sel);
        }

        private void BtnDecrease_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = (Range)Globals.ThisAddIn.Application.Selection;
            NumericalMethods.plusCell(-1, sel);
        }

        private void BtnSuccessive_Click(object sender, RibbonControlEventArgs e)
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

        private void BtnBisect_Click(object sender, RibbonControlEventArgs e)
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

        private void deprecatedFormulae_Click(Object sender, RibbonControlEventArgs e)
        {
            //var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            //var result =
            //    ValidationFormula.validate(wb)
            //    .Select(tpl => tpl.Item1 + tpl.Item2 + tpl.Item3)
            //    .ToArray()
            //    ;

            //if (result.Length == 0)
            //{
            //    MessageBox.Show("当前工作簿公式都支持！");
            //}
            //else
            //{
            //    var text = String.Join(Environment.NewLine, result);
            //    var dlg = new TextWindow("不支持的公式", text);
            //    dlg.ShowDialog();
            //}

        }

        private void clearName_button_Click(object sender, RibbonControlEventArgs e)
        {
            //var wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            //var names = wb.Names
            //    .Cast<Name>()
            //    .Where(nm => nm.Visible)
            //    .Select(nm => new Tuple<string, string>(nm.Name, (string)nm.RefersTo))
            //    .ToArray()
            //    ;

            //var cells =
            //    wb.Worksheets
            //    .Cast<Worksheet>()
            //    .SelectMany(wsx =>
            //        Traversal.getCellsOfWorksheet(wsx)
            //        .Where(rg => (bool)rg.HasFormula)
            //        .Select(rg => new Tuple<string, string, string>(wsx.Name, rg.get_Address(), (string)rg.Formula))
            //    )
            //    .ToArray()
            //    ;

            //var result =
            //    NameOps.replaceNames(names, cells)
            //    .Select(tpl => $"Sheets({Quotation.quote(tpl.Item1)}).Range(\"{tpl.Item2}\").Formula = {Quotation.quote(tpl.Item3)}")
            //    .ToArray()
            //    ;

            //if (result.Length == 0)
            //{
            //    MessageBox.Show("当前工作簿没有使用的名称！");
            //}
            //else
            //{
            //    var text = String.Join(Environment.NewLine, result);
            //    var dlg = new TextWindow("清除名称", text);
            //    dlg.ShowDialog();

            //}

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
        private void btn_referencesOfWorksheet_Click(object sender, RibbonControlEventArgs e)
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
        private void BtnDeleteRelayCell_Click(object sender, RibbonControlEventArgs e)
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

        private void btn_dependentsOfWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            //var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            //var ws = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            //var cells = wb.Worksheets
            //    .Cast<Worksheet>()
            //    .Where(wsx => wsx.Name != ws.Name)
            //    .SelectMany(wsx =>
            //        Traversal.getCellsOfWorksheet(wsx)
            //        .Where(rg => (bool)rg.HasFormula)
            //        .Select(rg => new Tuple<string, string, string>(wsx.Name, rg.get_Address(), (string)rg.Formula))
            //    )
            //    .ToArray();

            //var result =
            //    WorksheetOps.dependents(ws.Name, cells)
            //    .Select((tuple) => tuple.Item1 + tuple.Item2 + tuple.Item3)
            //    .ToArray();

            //if (result.Length == 0)
            //{
            //    MessageBox.Show("当前工作表没有引用其他工作表！");
            //}
            //else
            //{
            //    var text = String.Join(Environment.NewLine, result);
            //    var dlg = new TextWindow("工作表依赖", text);
            //    dlg.ShowDialog();

            //}


        }

        /// <summary>
        /// 标记跨表引用：扫描当前工作簿全部公式，把被本工作表以外的单元格引用的单元格背景色涂成洋红。
        /// 只处理工作簿内部的「工作表!地址」引用（含 3D 引用），外部工作簿引用与名称引用不处理。
        /// 核心逻辑在 F# 的 CrossSheetRefs.mark 中实现。
        /// </summary>
        private void BtnMarkCrossSheetRefs_Click(object sender, RibbonControlEventArgs e)
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

        private void btnSaveToSqlite_Click(object sender, RibbonControlEventArgs e)
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "另存为 SQLite 数据库",
                Filter = "SQLite 数据库 (*.db)|*.db|所有文件 (*.*)|*.*",
                DefaultExt = ".db",
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = wb.Path,
                FileName = System.IO.Path.GetFileNameWithoutExtension(wb.Name) + ".db"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    ExcelNumericalMethods.SqliteWorkbook.saveWorkbookAs(dlg.FileName, wb);
                    MessageBox.Show("工作簿已保存到 SQLite 数据库：\n" + dlg.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnCreateFromSqlite_Click(object sender, RibbonControlEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "从 SQLite 数据库创建工作簿",
                Filter = "SQLite 数据库 (*.db)|*.db|所有文件 (*.*)|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var wb = ExcelNumericalMethods.SqliteWorkbook.createWorkbookFrom(Globals.ThisAddIn.Application, dlg.FileName);
                    wb.Activate();
                    // 保存为与数据库文件同路径、同名称、仅扩展名为 .xlsx 的文件
                    var savePath = System.IO.Path.ChangeExtension(dlg.FileName, ".xlsx");
                    wb.SaveAs(savePath);
                    MessageBox.Show("已从数据库创建新工作簿并保存为：\n" + savePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// 脱公式：扫描全部工作表，将公式中含标记函数的单元格固化为当前计算结果，生成副本。
        /// </summary>
        private void BtnStripFormulas_Click(object sender, RibbonControlEventArgs e)
        {
            PublishStripFormulas();
        }

        /// <summary>
        /// 更新默认值：扫描当前工作簿全部工作表，不限函数名，处理所有公式为
        /// =IFERROR(函数(...), 常量) 的单元格：将兜底常量替换为该单元格最新真值，并以绿色标记。
        /// 直接在当前工作簿上位修改，不生成副本、不弹输入框与结果确认框（结果用状态栏提示）。
        /// 核心逻辑（扫描、过滤、改写、标色、统计）在 F# 的 Publishing.updateDefaults 中实现。
        /// </summary>
        private void BtnUpdateDefaults_Click(object sender, RibbonControlEventArgs e)
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
                        parts.Add($"跳过数组公式 {result.ArrayFormulaCells.Length} 个");
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
        /// 包裹自定义函数：把活动单元格的公式 =函数(...) 包裹为 =IFERROR(函数(...), 当前真值)。
        /// 兜底值取当前计算结果（.NET "0.##" 格式）；非公式、已是 IFERROR、错误值等不处理。
        /// 结果在状态栏提示；核心逻辑在 F# 的 Publishing.wrapFunction 中实现。
        /// </summary>
        private void BtnWrapFunction_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var cell = app.ActiveCell as Microsoft.Office.Interop.Excel.Range;
            if (cell == null)
            {
                MessageBox.Show("当前没有活动单元格！");
                return;
            }

            try
            {
                var result = Publishing.wrapFunction(cell);
                app.StatusBar = result.Message;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 脱公式：不限函数名，把所有 =IFERROR(函数(...), 常量) 结构的单元格固化为当前计算结果，
        /// 另存副本，副本名为「当前工作簿名（崔胜利）.xlsx」（与源工作簿同目录）。源工作簿全程不被修改。
        /// 核心逻辑（扫描、过滤、固化、统计）在 F# 的 Publishing.stripFormulas 中实现。
        /// </summary>
        private void PublishStripFormulas()
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("当前没有打开的工作簿！");
                return;
            }

            var oldAlerts = app.DisplayAlerts;
            string tempPath = null;
            try
            {
                app.DisplayAlerts = false;

                // 副本名：当前工作簿名（崔胜利）.xlsx；未保存的工作簿存到文档目录
                var dir = string.IsNullOrEmpty(wb.Path)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : wb.Path;
                var baseName = System.IO.Path.GetFileNameWithoutExtension(wb.Name) + "（崔胜利）";
                var finalPath = System.IO.Path.Combine(dir, baseName + ".xlsx");
                var isXlsx = string.Equals(System.IO.Path.GetExtension(wb.Name), ".xlsx", StringComparison.OrdinalIgnoreCase);

                string copyPath;
                if (isXlsx)
                {
                    wb.SaveCopyAs(finalPath);
                    copyPath = finalPath;
                }
                else
                {
                    // 宏工作簿先按原格式另存副本，处理后再另存为 xlsx（不含宏）
                    tempPath = System.IO.Path.Combine(dir, baseName + System.IO.Path.GetExtension(wb.Name));
                    wb.SaveCopyAs(tempPath);
                    copyPath = tempPath;
                }

                var copy = app.Workbooks.Open(copyPath);
                try
                {
                    var result = Publishing.stripFormulas(copy);
                    if (isXlsx)
                    {
                        copy.Save();
                    }
                    else
                    {
                        copy.SaveAs(finalPath, XlFileFormat.xlOpenXMLWorkbook);
                    }

                    var text = Publishing.formatSummary(false, finalPath, result);
                    if (result.ErrorCells.Length == 0
                        && result.UnconformCells.Length == 0
                        && result.ArrayFormulaCells.Length == 0
                        && result.ProtectedSheets.Length == 0)
                    {
                        MessageBox.Show(text);
                    }
                    else
                    {
                        var detail = new TextWindow("脱公式结果", text);
                        detail.ShowDialog();
                    }
                }
                finally
                {
                    copy.Close(false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (tempPath != null && System.IO.File.Exists(tempPath))
                {
                    System.IO.File.Delete(tempPath);
                }
                app.DisplayAlerts = oldAlerts;
            }
        }
    }
}
