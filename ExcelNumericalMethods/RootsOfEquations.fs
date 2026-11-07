module ExcelNumericalMethods.RootsOfEquations

open Microsoft.Office.Interop.Excel
open System.Text.RegularExpressions
open System.Windows.Forms

///匹配当前工作表的A1形式单元格地址（如 A1、$A$1、A$1、$A1）。
///列名限1~3个字母（A..XFD），因此形如名称、跨工作表、跨工作簿的引用都不会匹配。
let successiveRgx = Regex(@"^=\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*-\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*$")

let bisectRgx = Regex(@"^=\s*\(\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\+\s*(\$?[A-Za-z]{1,3}\$?\d+)\s*\)\s*/\s*2\s*$")

///减法操作数（被减数 / 减数）的可改性。
type Mutability =
    ///不以等号开头的字面量：可以写入
    | Mutable
    ///以等号开头的公式：值由 Excel 算出，写进去会被重算覆盖，故不可变
    | FormulaImmutable
    ///数组公式的一部分：旧式 CSE 只能整块改，新式溢出区改一格会静默破坏溢出区，故不可变
    | ArrayImmutable

///代入法追赶一次：误差单元格公式为 =A2-A1，自动识别不可变格（新值）与可变格（旧值），
///把不可变格的当前值写入可变格，使差值趋近于零。两格同可变 / 同不可变时弹窗提示。
///两个引用格必须都在当前工作表。
let successive (deltaCell: Range) =
    if not (unbox deltaCell.HasFormula) then
        failwith "所选单元格应该是公式，例如 =A2-A1"
    let formula = unbox<string> deltaCell.Formula
    let m = successiveRgx.Match(formula)
    if not m.Success then
        failwithf "公式应该为=A2-A1，且只引用当前工作表的单元格: %s" formula
    let ws = deltaCell.Worksheet

    let cell1 = ws.Range(m.Groups.[1].Value)
    let cell2 = ws.Range(m.Groups.[2].Value)

    ///单元格的简明描述：地址、公式/值、属于哪种数组公式
    let describe (cell: Range) =
        let kind =
            if ArrayFormula.hasLegacyArray cell then "旧式 CSE 数组公式"
            elif ArrayFormula.hasSpill cell then
                if ArrayFormula.isSpillAnchor cell then "新式溢出公式·锚点"
                else "新式溢出公式·溢出区"
            elif unbox cell.HasFormula then "普通公式"
            else "字面量"
        sprintf "%s（%s）：%s" (cell.Address()) kind (string cell.Formula)

    ///操作数的可改性：数组公式的一部分不可变，以等号开头的公式不可变，其余字面量可变。
    let mutability (cell: Range) =
        if ArrayFormula.isArrayPart cell then ArrayImmutable
        elif (string cell.Formula).TrimStart().StartsWith("=") then FormulaImmutable
        else Mutable

    let m1 = mutability cell1
    let m2 = mutability cell2
    if (m1 = Mutable) <> (m2 = Mutable) then
        //恰好一个可变：把不可变格（新值）的当前值写入可变格（旧值），后者追前者。
        //不可变格是公式或数组公式的一部分，只读不写，故不会破坏数组公式。
        let fixedCell = if m1 = Mutable then cell2 else cell1
        let mutableCell = if m1 = Mutable then cell1 else cell2
        mutableCell.Value2 <- fixedCell.Value2
    else
        //两个都可变（同为字面量）或都不可变（常规公式 / 数组公式）：无从判断该改哪一个
        let hint =
            if m1 = Mutable && m2 = Mutable then
                "两格都是字面量：请把要保留的那一格改为\"=字面量\"形式的公式（如 =A2）锁定，另一格才会被追改。"
            else
                "两格都是公式（常规公式或数组公式），都不可变：请把其中一格改为字面量使其可变，再执行。"
        MessageBox.Show(
            sprintf "无法追值：两个操作数必须恰好有一个是可写的字面量。\n%s\n被减数 %s\n减数 %s"
                hint (describe cell1) (describe cell2)
        ) |> ignore

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
