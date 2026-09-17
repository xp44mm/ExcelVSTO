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
                    var suggestedPath = System.IO.Path.ChangeExtension(dlg.FileName, ".xlsx");
                    try
                    {
                        System.Windows.Forms.Clipboard.SetText(suggestedPath);
                        MessageBox.Show("已从数据库创建新工作簿。\n建议保存路径已复制到剪贴板（可直接粘贴）：\n" + suggestedPath);
                    }
                    catch
                    {
                        MessageBox.Show("已从数据库创建新工作簿。\n建议保存路径（复制到剪贴板失败，请手动记录）：\n" + suggestedPath);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
