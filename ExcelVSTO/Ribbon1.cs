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

        private void btn_referencesOfWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            //var ws = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            //var cells =
            //    Traversal.getCellsOfWorksheet(ws)
            //    .Where(rg => (bool)rg.HasFormula)
            //    .Select(rg => new Tuple<string, string>(rg.get_Address(), (string)rg.Formula))
            //    .ToArray();

            //var inputs =
            //    WorksheetOps.references(ws.Name, cells)
            //    .Select((tuple) => tuple.Item1 + tuple.Item2)
            //    .ToArray();

            //if (inputs.Length == 0)
            //{
            //    MessageBox.Show("当前工作表没有引用其他工作表！");
            //}
            //else
            //{
            //    var text = String.Join(Environment.NewLine, inputs);
            //    var dlg = new TextWindow("工作表引用", text);
            //    dlg.ShowDialog();

            //}

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
        /// 脱公式主流程：输入标记函数名 → 预检（只读）→ SaveCopyAs 生成副本 → 在副本上批量处理 →
        /// 宏工作簿另存为 xlsx → 统计提示。源工作簿全程不被修改。
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

            var action = "脱公式";
            var dlg = new InputWindow(action, "电机额定功率");
            if (dlg.ShowDialog() != true)
            {
                return; // 关闭窗口即取消
            }
            var marker = (dlg.Input ?? "").Trim();
            if (marker.Length == 0)
            {
                MessageBox.Show("标记函数名不能为空，已终止。");
                return;
            }

            var oldAlerts = app.DisplayAlerts;
            string tempPath = null;
            string finalPath = null;
            var aborted = false;
            try
            {
                app.DisplayAlerts = false;

                // 预检（只读）：未找到目标公式、或所有目标单元格均为错误值（可能本机无此函数）时终止
                var (targetCount, allError) = Publishing.precheck(wb, marker);
                if (targetCount == 0)
                {
                    MessageBox.Show($"未找到包含「{marker}」的公式。");
                    return;
                }
                if (allError)
                {
                    MessageBox.Show($"本机无此函数「{marker}」（或所有目标单元格均为错误值），处理已终止，请先排查。");
                    return;
                }

                // 处理前检查标记函数是否存在（=ERROR.TYPE(第一参数) 求值探测，仅只读不写单元格）
                var exists = Publishing.checkFunctionExists(wb, marker);
                if (exists.HasValue && !exists.Value)
                {
                    MessageBox.Show($"本机无此函数「{marker}」，处理已终止。");
                    return;
                }

                // 生成带时间戳的副本
                var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var dir = string.IsNullOrEmpty(wb.Path)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : wb.Path;
                var baseName = "发布版_固化_" + stamp;
                finalPath = System.IO.Path.Combine(dir, baseName + ".xlsx");
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
                    var result = Publishing.run(false, copy, marker);
                    if (result.FunctionMissing)
                    {
                        // 副本上探测到本机无此函数：不保存副本，提示并终止
                        aborted = true;
                        MessageBox.Show($"本机无此函数「{marker}」，处理已终止。");
                        return;
                    }
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
                        var detail = new TextWindow(action + "结果", text);
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
                if (aborted && System.IO.File.Exists(finalPath))
                {
                    // 中止时清理已生成的副本文件（宏工作簿的中间副本由 tempPath 处理）
                    System.IO.File.Delete(finalPath);
                }
                if (tempPath != null && System.IO.File.Exists(tempPath))
                {
                    System.IO.File.Delete(tempPath);
                }
                app.DisplayAlerts = oldAlerts;
            }
        }
    }
}
