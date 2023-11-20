module ExcelNumericalMethods.NumericalMethods

open Microsoft.Office.Interop.Excel

/// = rg.Cells.[r, c] :?> Range
let cell r c (rg:Range) = rg.Cells.[r, c] :?> Range

/// 范围左上角的单元格
let firstCell (rg:Range) = rg |> cell 1 1

///判断单元格是否为真空
let isEmpty (cell:Range) =
    if unbox cell.HasFormula then
        false
    else
        System.String.IsNullOrEmpty(unbox cell.Formula)

///范围单元格保持行列的数组
let getCellArrayOfRange(rg:Range) =
    let rows = rg.Rows.Count
    let cols = rg.Columns.Count
    [|
        for r in 1..rows ->
            [|
                for c in 1..cols do
                    let cell = rg.Cells.[r, c] :?> Range
                    yield cell
            |]
    |]

///合并单元格及其中的内容
let merge(rg:Range) =
    let s =
        getCellArrayOfRange rg
        |> Array.map(
            Array.map(fun cell ->
                match cell.Value2 with
                | null -> ""
                | c -> c.ToString())
            >> String.concat ""
        )
        |> String.concat System.Environment.NewLine
    rg.Merge()
    (firstCell rg).Value2 <- s

///单元格加1减1
let plusCell(operand, rg: Range) =
    let cells = Traversal.getCellsOfRange rg
    cells |> Seq.iter(fun cell ->
                 match cell.Value2 with
                 | :? int as n -> cell.Value2 <- float n + operand
                 | :? float as f -> cell.Value2 <- f + operand
                 | _ -> ())

///清空下方等值的列
let tidyColumns(rg: Range) =
    let rec clearColumn =
        function
        | [] -> [||]
        | head :: tail ->
            let diff c = c <> head && not(isNull c)
            [|//保留第一个元素
              yield false
              match List.tryFindIndex diff tail with
              | None -> yield! tail |> List.map(fun _ -> true)
              | Some j ->
                  yield! [0..j - 1] |> List.map(fun _ -> true)
                  yield! clearColumn(tail |> List.skip j)|]

    //按列排列的数组
    let cellArray =
        getCellArrayOfRange rg |> Array.transpose

    //每个单元格是否应该清除的数据
    let clears =
        cellArray
        |> Array.map(Array.map(fun cell -> cell.Value2))
        |> Array.map(Array.toList >> clearColumn)

    //副作用
    Array.zip clears cellArray
    |> Array.map(fun (bs, cs) -> Array.zip bs cs)
    |> Array.iter(Array.iter(fun (b, c) ->
                      if b then c.ClearContents() |> ignore))

///用第一非空单元格填充下方为空的单元格
let fillColumns(rg: Range) =
    let rec fillColumn =
        function
        | [] -> [||]
        | head :: tail ->
            let diff c = c <> head && not(isNull c)
            [|//保留第一个元素
              yield head
              match List.tryFindIndex diff tail with
              | None -> yield! tail |> List.map(fun _ -> head)
              | Some j ->
                  yield! [0..j - 1] |> List.map(fun _ -> head)
                  yield! fillColumn(tail |> List.skip j)|]

    //按列排列的数组
    let cellArray =
        getCellArrayOfRange rg |> Array.transpose

    //要填充的内容
    let contents =
        cellArray
        |> Array.map(Array.map(fun cell -> cell.Value2))
        |> Array.map(Array.toList >> fillColumn)

    Array.zip contents cellArray
    |> Array.map(fun (bs, cs) -> Array.zip bs cs)
    |> Array.iter(Array.iter(fun (b, c) -> c.Value2<-b))

///交替着色
let alternateColor(rg: Range) =
    let cellArray = getCellArrayOfRange rg

    //每行首列单元格的值
    let firstValues =
        cellArray |> Array.map(fun cells -> cells.[0].Value2)

    //检测每一行是否需要着色
    let rec loop color =
        function
        | [] -> [||]
        | lst ->
            [|let takes =
                  lst
                  |> List.takeWhile((=) lst.Head)
                  |> List.map(fun _ -> color)
              yield! takes
              yield! loop (not color) (List.skip takes.Length lst)|]

    //找到每个单元格应该如何涂色
    let colors = loop true (List.ofArray firstValues)

    //交替颜色取第一个单元格的颜色
    let firstColor =
        let firstCell = rg |> firstCell
        firstCell.Interior.Color

    colors
    |> Array.zip cellArray
    |> Array.iter
           (fun (cells, color) ->
           if color then
               cells |> Seq.iter(fun cell -> cell.Interior.Color <- firstColor))

//插入分组行，分组行是一整行
let split(rg: Range) =
    let cellArray = getCellArrayOfRange rg

    //每行首列单元格的值和行号
    let vrows =
        cellArray |> Array.map(fun cells -> cells.[0].Value2, cells.[0].Row)

    //找到分隔行的行号
    let rec getSeprateRows =
        function
        | [] -> []
        | head::tail ->
            [
                let diff c = c <> fst head && not(isNull c)

                match List.tryFindIndex (fst>>diff) tail with
                | None -> ()
                | Some j ->
                    let tail = tail |> List.skip j
                    yield snd tail.Head
                    yield! getSeprateRows tail
            ]
    let rows = getSeprateRows (List.ofArray vrows)

    //执行副作用
    rows
    |> List.rev //先插入下方的空行
    |> List.iter(fun r ->
        let row = rg.Worksheet.Range(sprintf "%d:%d" r r)

        if unbox <| row.Insert(XlDirection.xlDown) then
            row.Offset(-1,0).ClearFormats() |> ignore
    )

    rg

///删除空行
let removeBlank (rg: Range) =
    let cellArray = getCellArrayOfRange rg

    //找到空行的行号
    let rows =
        cellArray
        |> Array.choose(fun row ->
            if row |> Array.forall(isEmpty) then
                Some row.[0].Row
            else
                None
        )

    rows
    |> Array.rev //先删除下面的行，保持上面的行号不变
    |> Array.iter(fun r ->
        let row = rg.Worksheet.Range(sprintf "%d:%d" r r)
        row.Delete(XlDirection.xlUp)
        |> ignore
    )

    rg

