module ExcelNumericalMethods.RootsOfEquations

open Microsoft.Office.Interop.Excel
open System.Text.RegularExpressions

///匹配当前工作表的A1形式单元格地址（如 A1、$A$1、A$1、$A1）。
///列名限1~3个字母（A..XFD），因此形如名称、跨工作表、跨工作簿的引用都不会匹配。
let successiveRgx = Regex(@"^=\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*-\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*$")

let bisectRgx = Regex(@"^=\s*\(\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\+\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\)\s*/\s*2\s*$")

///代入法追赶一次：目标单元格的减数等于被减数，后者追前者。
///误差单元格应输入公式 =A2-A1，A2 是新值，A1 是旧值（必须是字面量），两者都在当前工作表。
let successive (deltaCell: Range) =
    if not (unbox deltaCell.HasFormula) then
        failwith "误差单元格应该是公式，例如 =A2-A1"
    let formula = unbox<string> deltaCell.Formula
    let m = successiveRgx.Match(formula)
    if not m.Success then
        failwithf "公式应该为=A2-A1，且只引用当前工作表的A1地址: %s" formula
    let ws = deltaCell.Worksheet
    let targetCell = ws.Range(m.Groups.[1].Value) // new value
    let changeCell = ws.Range(m.Groups.[2].Value) // old value
    if unbox changeCell.HasFormula then
        failwithf "输入值应该是字面量: %s" (changeCell.Address())
    else
        changeCell.Value2 <- targetCell.Value2

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
