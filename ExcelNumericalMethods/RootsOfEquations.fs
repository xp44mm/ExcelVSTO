module ExcelNumericalMethods.RootsOfEquations

open Microsoft.Office.Interop.Excel
open ExcelCompiler
open FSharp.Idioms

let friendAddress(cell: Range) = 
    sprintf "%s!%s" cell.Worksheet.Name (cell.Address())

let parseCell(cell: Range) = 
    if unbox cell.HasArray then 
        failwithf "此单元格不应是数组公式，地址为:%s" (friendAddress cell)
    elif unbox cell.HasFormula then
        let expr = 
            let formula = unbox<string> cell.Formula
            formula.TrimStart('=')
        
        ExcelFormulaString.parseToExpr expr
    else
        failwithf "此单元格应该是公式，地址为:%s" (friendAddress cell)

///去单引号
let deapos(s:string) =  s.[1..s.Length-2].Replace("''","'")

///从公式中获取工作表的名称
let smartDeapos (ws:string) =
    if ws.StartsWith("'") then deapos ws else ws

///实例化工作表
let getWorksheet (aws:Worksheet) (ws:string list) =
    match ws with
    | [] -> aws
    | [ws] ->
        let ws = smartDeapos ws
        (aws.Parent :?> Workbook).Worksheets.[ws] :?> Worksheet
    | _ -> failwithf "%A" ws

let toRange (aws:Worksheet) (ws:string list) (addr:string) =
    let ws = getWorksheet aws ws
    ws.Range(addr)

///代入法追赶一次：目标单元格的减数等于被减数，后者追前者
let successive(deltaCell: Range) =
    match parseCell deltaCell with
    | Sub(Reference(ws0,[addr0]), Reference(ws1,[addr1])) ->
            let ws = deltaCell.Worksheet
            let targetCell = toRange ws ws0 addr0 // new value
            let changeCell = toRange ws ws1 addr1 // old value
            if unbox changeCell.HasFormula then
                failwithf "输入值应该是字面量: %s" (friendAddress changeCell)
            else
                changeCell.Value2 <- targetCell.Value2
    | _ ->
        failwithf "公式应该为=A2-A1, 误差单元格: %s" (friendAddress deltaCell)

/// 执行一次对分法
let bisect(averageCell: Range) =
    match parseCell averageCell with
    | Div(Add(Reference(ws1,[addr1]),Reference(ws2,[addr2])),Number "2") ->
            let aws = averageCell.Worksheet
            let cell1 = toRange aws ws1 addr1
            let cell2 = toRange aws ws2 addr2

            if unbox cell1.HasFormula then
                failwithf "单元格应该输入数值，地址为：%s" (friendAddress cell1)
            elif unbox cell2.HasFormula then
                failwithf "单元格应该输入数值，地址为：%s" (friendAddress cell2)
            else
                //平均单元格下一行的单元格是目标单元格,我们希望目标单元格值为零。
                let goalCell = averageCell.get_Offset(1, 0)

                if unbox goalCell.Value2 < 0.0
                then cell1.Value2 <- averageCell.Value2 //如果目标单元格的值小于零，使前面单元格的值为平均单元格的值
                else cell2.Value2 <- averageCell.Value2 //如果目标单元格的值大于零，使后面单元格的值为平均单元格的值
    | _ ->
        failwithf "公式应该为=(A1+A2)/2, 单元格: %s" (friendAddress averageCell)

