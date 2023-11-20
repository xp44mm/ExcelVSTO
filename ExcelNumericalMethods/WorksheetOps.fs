module ExcelNumericalMethods.WorksheetOps

open ExcelCompiler
open FSharp.Idioms
open FSharp.Idioms.StringOps

///ws工作表的外部引用，不检查名称对工作表的引用
/// ws=worksheet.Name
/// cell = (rg.get_Address(),rg.Formula)
let references (ws:string) (cells:(string*string)[]) =
    // 判断 tok 是其他单元格的引用
    let isOutReference (tok:ExcelToken) =
        match tok with
        | REFERENCE([wsx],_) when ws != Apostrophe.smartDeapos wsx -> true
        | _ -> false

    cells
    |> Array.Parallel.choose(fun(addr,formula)->
        let tokens = 
            formula  
            //|> ExcelTokenUtils.tokenize
            //|> ExcelDFA.analyze
            //|> Seq.concat
            |> ExcelTokenUtils.tokenize 0
            //|> ExcelDFA.analyze
            |> ExcelExprCompiler.analyze
            //|> Seq.concat
            |> Seq.map(fun{value=tok}->tok)


        if tokens |> Seq.exists(fun tok -> isOutReference tok) then
            Some(addr, formula)
        else None
    )

///其他工作表对ws工作表的依赖
let dependents (ws:string) (cells:(string*string*string)[]) =
    let isReferenceWs ws (tok:ExcelToken) =
        match tok with
        | REFERENCE([wsx],_) when ws == Apostrophe.smartDeapos wsx -> true
        | _ -> false

    cells
    |> Array.Parallel.choose(fun(wsx,addr,formula)->
        let tokens = 
            formula  
            //|> ExcelTokenUtils.tokenize
            //|> ExcelDFA.analyze
            //|> Seq.concat
            |> ExcelTokenUtils.tokenize 0
            //|> ExcelDFA.analyze
            |> ExcelExprCompiler.analyze
            //|> Seq.concat
            |> Seq.map(fun{value=tok}->tok)

        if tokens |> Seq.exists(fun tok -> isReferenceWs ws tok) then
            Some(wsx, addr, formula)
        else None
    )
