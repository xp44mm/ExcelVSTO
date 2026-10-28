module ExcelNumericalMethods.RootsOfEquations

open Microsoft.Office.Interop.Excel
open System.Text.RegularExpressions
open System.Windows.Forms

///匹配当前工作表的A1形式单元格地址（如 A1、$A$1、A$1、$A1）。
///列名限1~3个字母（A..XFD），因此形如名称、跨工作表、跨工作簿的引用都不会匹配。
let successiveRgx = Regex(@"^=\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*-\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*$")

let bisectRgx = Regex(@"^=\s*\(\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\+\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\)\s*/\s*2\s*$")

///代入法追赶一次：误差单元格公式为 =A2-A1，自动识别公式格（新值）与字面量格（旧值），
///把公式格的当前值写入字面量格，使差值趋近于零。两格同为公式 / 同为字面量时弹窗提示。
///两个引用格必须都在当前工作表。
let successive (deltaCell: Range) =
    if not (unbox deltaCell.HasFormula) then
        failwith "所选单元格应该是公式，例如 =A2-A1"
    let formula = unbox<string> deltaCell.Formula
    let m = successiveRgx.Match(formula)
    if not m.Success then
        failwithf "公式应该为=A2-A1，且只引用当前工作表的单元格: %s" formula
    let ws = deltaCell.Worksheet
    //智能判断：根据两个引用单元格是公式还是字面量，决定追改哪一个。
    //如果引用的两个单元格都是公式"=xxx"，则弹出消息框提示，要改变的单元格
    //如果引用的两个单元格都是字面量，则弹出消息框提示，目标单元格要用"=字面量"公式锁定。
    //如果引用的两个单元格一个是公式，一个是字面量，则修改字面量单元格的值。
    let cell1 = ws.Range(m.Groups.[1].Value)
    let cell2 = ws.Range(m.Groups.[2].Value)
    let c1IsFormula = unbox cell1.HasFormula
    let c2IsFormula = unbox cell2.HasFormula
    if c1IsFormula && c2IsFormula then
        MessageBox.Show(
            sprintf "两个单元格都是公式，请把要追改的那个单元格改为字面量后再执行。\n被减数(%s): %s\n减数(%s): %s"
                (cell1.Address()) (cell1.Formula |> unbox<string>)
                (cell2.Address()) (cell2.Formula |> unbox<string>)
        ) |> ignore
    elif (not c1IsFormula) && (not c2IsFormula) then
        MessageBox.Show(
            sprintf "两个单元格都是字面量，目标单元格要用\"=字面量\"公式锁定后再执行。\n被减数(%s): %A\n减数(%s): %A"
                (cell1.Address()) (cell1.Value2)
                (cell2.Address()) (cell2.Value2)
        ) |> ignore
    else
        //一个公式一个字面量：把公式单元格（新值）写入字面量单元格（旧值），后者追前者
        let formulaCell = if c1IsFormula then cell1 else cell2
        let literalCell = if c1IsFormula then cell2 else cell1
        literalCell.Value2 <- formulaCell.Value2

/// 执行一次对分法。
///平均单元格应输入公式 =(A1+A2)/2，A1、A2 是上下界（必须是字面量），两者都在当前工作表。
let bisect (averageCell: Range) =
    if not (unbox averageCell.HasFormula) then
        failwith "平均值单元格应该是公式，例如 =(A1+A2)/2"
    let formula = unbox<string> averageCell.Formula
    let m = bisectRgx.Match(formula)
    if not m.Success then
        failwithf "公式应该为=(A1+A2)/2，且只引用当前工作表的A1地址: %s" formula
    let ws = averageCell.Worksheet
    let cell1 = ws.Range(m.Groups.[1].Value)
    let cell2 = ws.Range(m.Groups.[2].Value)

    if unbox cell1.HasFormula then
        failwithf "单元格应该输入数值，地址为：%s" (cell1.Address())
    elif unbox cell2.HasFormula then
        failwithf "单元格应该输入数值，地址为：%s" (cell2.Address())
    else
        //平均单元格下一行的单元格是目标单元格,我们希望目标单元格值为零。
        let goalCell = averageCell.get_Offset(1, 0)

        if unbox goalCell.Value2 < 0.0
        then cell1.Value2 <- averageCell.Value2 //如果目标单元格的值小于零，使前面单元格的值为平均单元格的值
        else cell2.Value2 <- averageCell.Value2 //如果目标单元格的值大于零，使后面单元格的值为平均单元格的值
