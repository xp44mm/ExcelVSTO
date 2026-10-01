open System
open System.IO
open Microsoft.Office.Interop.Excel
open ExcelNumericalMethods

/// 模拟发布辅助功能的 Excel COM 冒烟测试：
/// 1) 单格数组公式的写入方式探测（Value2 / Formula / ClearContents+Value2 / FormulaArray）
/// 2) 更新默认值：普通目标单元格 + 单格数组公式更新、错误值跳过、非常量兜底跳过、多格数组公式跳过、普通公式保留、固定绿色标色
/// 3) 脱公式：单格数组公式固化、普通目标固化、普通公式保留
/// 4) 真实工作簿副本：燃烧器功率计算39.xlsx（标记函数 电机额定功率）

let defaultYellow = 10284031.0   // 0x009CEBFF (FFEB9C 的 BGR)
let green = 5296274.0            // RGB(146,208,80) 的 BGR 0x0050D092（Publishing.UpdateDefaultsColor）

let cellOf (ws: Worksheet) (r: int) (c: int) : Range = ws.Cells.[r, c] :?> Range

let setv (ws: Worksheet) (r: int) (c: int) (v: obj) = (cellOf ws r c).Value2 <- v

let expect (cond: bool) (msg: string) =
    if cond then printfn "  ✓ %s" msg
    else printfn "  ✗ 失败：%s" msg

[<EntryPoint>]
let main _ =
    let mutable failed = 0
    let mutable app : Application = null
    try
        let appInst = ApplicationClass()
        appInst.Visible <- false
        appInst.DisplayAlerts <- false
        app <- appInst
        // ================= Part 1：单格数组公式写入方式探测 =================
        printfn "== Part 1 单格数组公式写入探测 =="
        let wb1 = app.Workbooks.Add()
        let ws1 = wb1.Worksheets.[1] :?> Worksheet
        ws1.Name <- "设备选型"
        setv ws1 11 4 9.0   // D11 = 9
        let d12 = cellOf ws1 12 4
        d12.FormulaArray <- "=IFERROR(SQRT(D11),22)"
        printfn "  创建后 HasArray=%b FormulaArray=%s 值=%A" (unbox<bool> d12.HasArray) (string d12.FormulaArray) d12.Value2
        try
            d12.Value2 <- 3.0
            printfn "  Value2 直接赋值成功 HasArray=%b" (unbox<bool> d12.HasArray)
        with e -> printfn "  Value2 直接赋值抛异常：%s" e.Message
        // 重建数组公式再试 ClearContents
        d12.FormulaArray <- "=IFERROR(SQRT(D11),22)"
        try
            d12.ClearContents()
            printfn "  ClearContents 成功 HasArray=%b" (unbox<bool> d12.HasArray)
        with e -> printfn "  ClearContents 抛异常：%s" e.Message
        // 重建数组公式再试 Formula 赋值
        d12.FormulaArray <- "=IFERROR(SQRT(D11),22)"
        try
            d12.Formula <- "=IFERROR(SQRT(D11),3)"
            printfn "  Formula 赋值成功 HasArray=%b" (unbox<bool> d12.HasArray)
        with e -> printfn "  Formula 赋值抛异常：%s" e.Message
        wb1.Close(false) |> ignore

        // ================= Part 2：更新默认值（含固定绿色标色） =================
        printfn "== Part 2 更新默认值 =="
        let wb2 = app.Workbooks.Add()
        let ws2 = wb2.Worksheets.[1] :?> Worksheet
        ws2.Name <- "设备选型"
        setv ws2 1 1 4.0                      // A1 = 4
        (cellOf ws2 1 2).Formula <- "=IFERROR(SQRT(A1),22)"      // B1 普通目标，真值 2
        (cellOf ws2 2 2).Formula <- "=IFERROR(SQRT(-1),1/0)"     // B2 兜底为错误 → 单元格 #DIV/0!，跳过
        (cellOf ws2 3 2).Formula <- "=A1*3"                      // B3 普通公式，保留
        (cellOf ws2 4 2).Formula <- "=IFERROR(SQRT(A1),B1)"      // B4 兜底为引用 → 非常量，跳过
        setv ws2 11 4 9.0                                        // D11 = 9
        (cellOf ws2 12 4).FormulaArray <- "=IFERROR(SQRT(D11),22)" // D12 单格数组公式，真值 3
        ws2.Range("E5:F6").FormulaArray <- "=SQRT(A1:A2)"        // E5:F6 多格数组公式，跳过
        let d12dbg = cellOf ws2 12 4
        printfn "  调试 D12.Pattern=%A int=%d xlNone=%d sampleColor=%A" d12dbg.Interior.Pattern (int (d12dbg.Interior.Pattern :?> XlPattern)) (int XlPattern.xlPatternNone) (Publishing.sampleColor wb2)
        let r1 = Publishing.run(true, wb2, "SQRT")
        printfn "  run(true) processed=%d（期望 2：B1、D12）" r1.ProcessedCount
        printfn "  ArrayFormulaCells=%d（期望 4：E5:F6 按单元格计数）" r1.ArrayFormulaCells.Length
        printfn "  ErrorCells=%d（期望 1：B2 #DIV/0!）" r1.ErrorCells.Length
        printfn "  UnconformCells=%d（期望 1：B4 兜底为引用）" r1.UnconformCells.Length
        expect (r1.ProcessedCount = 2) "processed = 2"
        expect (r1.ArrayFormulaCells.Length = 4) "多格数组公式 E5:F6（4 格）跳过"
        expect (r1.ErrorCells.Length = 1) "错误值 B2 跳过"
        expect (r1.UnconformCells.Length = 1) "非常量兜底 B4 跳过"
        let b1 = cellOf ws2 1 2
        let b4 = cellOf ws2 4 2
        let d12b = cellOf ws2 12 4
        let b3 = cellOf ws2 3 2
        expect (string b1.Formula = "=IFERROR(SQRT(A1),2)") (sprintf "B1 兜底更新：%s" (string b1.Formula))
        expect (string b4.Formula = "=IFERROR(SQRT(A1),B1)") "B4 非常量兜底公式未改动"
        expect (unbox<bool> d12b.HasArray) "D12 仍为数组公式"
        expect (string d12b.FormulaArray = "=IFERROR(SQRT(D11),3)") (sprintf "D12 兜底更新（FormulaArray 写入）：%s" (string d12b.FormulaArray))
        expect (string b3.Formula = "=A1*3") "B3 普通公式保留"
        expect (b1.Interior.Color = green) (sprintf "B1 标绿色：%A" b1.Interior.Color)
        expect (d12b.Interior.Color = green) (sprintf "D12 标绿色：%A" d12b.Interior.Color)
        expect (b3.Interior.Color <> green) "B3 未被标色"
        // 固定绿色标记测试：标记色与 D12 的填充色无关（先把 D12 改成浅黄，再跑一次新目标 B5 仍标固定绿）
        d12b.Interior.Color <- defaultYellow
        (cellOf ws2 5 2).Formula <- "=IFERROR(SQRT(A1),9)"       // B5 新目标
        let r2 = Publishing.run(true, wb2, "SQRT")
        let b5 = cellOf ws2 5 2
        expect (r2.ProcessedCount = 3) (sprintf "第二次 processed=3（B1/B5/D12）：%d" r2.ProcessedCount)
        expect (b5.Interior.Color = green) (sprintf "B5 标固定绿：%A" b5.Interior.Color)
        expect (b1.Interior.Color = green) (sprintf "B1 也被重新标绿：%A" b1.Interior.Color)
        wb2.Close(false) |> ignore

        // ================= Part 3：脱公式 =================
        printfn "== Part 3 脱公式 =="
        let wb3 = app.Workbooks.Add()
        let ws3 = wb3.Worksheets.[1] :?> Worksheet
        ws3.Name <- "设备选型"
        setv ws3 1 1 4.0
        (cellOf ws3 1 2).Formula <- "=IFERROR(SQRT(A1),22)"
        (cellOf ws3 2 2).Formula <- "=IFERROR(SQRT(-1),1/0)"
        (cellOf ws3 3 2).Formula <- "=A1*3"
        setv ws3 11 4 9.0
        (cellOf ws3 12 4).FormulaArray <- "=IFERROR(SQRT(D11),22)"
        ws3.Range("E5:F6").FormulaArray <- "=SQRT(A1:A2)"
        let r3 = Publishing.run(false, wb3, "SQRT")
        printfn "  run(false) processed=%d（期望 2：B1、D12）" r3.ProcessedCount
        printfn "  ArrayFormulaCells=%d（期望 4）" r3.ArrayFormulaCells.Length
        printfn "  ErrorCells=%d（期望 1）" r3.ErrorCells.Length
        expect (r3.ProcessedCount = 2) "processed = 2"
        expect (r3.ArrayFormulaCells.Length = 4) "多格数组公式跳过（4 格）"
        expect (r3.ErrorCells.Length = 1) "错误值跳过"
        let b1s = cellOf ws3 1 2
        let d12s = cellOf ws3 12 4
        expect (not (unbox<bool> b1s.HasFormula)) (sprintf "B1 已固化（无公式）：%A" b1s.Value2)
        expect (b1s.Value2 = 2.0) "B1 值 = 2"
        expect (not (unbox<bool> d12s.HasFormula)) (sprintf "D12 单格数组公式已固化（无公式）：%A" d12s.Value2)
        expect (d12s.Value2 = 3.0) "D12 值 = 3"
        expect (string (cellOf ws3 3 2).Formula = "=A1*3") "B3 普通公式保留"
        wb3.Close(false) |> ignore

        // ================= Part 4：真实工作簿副本（燃烧器功率计算39.xlsx） =================
        printfn "== Part 4 真实工作簿副本 =="
        let src = @"C:\崔胜利\群有\广西盛隆\燃烧器功率计算39.xlsx"
        if File.Exists src then
            let tmp = Path.Combine(Path.GetTempPath(), "发布_冒烟_39.xlsx")
            File.Copy(src, tmp, true)
            let wb4 = app.Workbooks.Open(tmp)
            let ws4 = wb4.Worksheets.["设备选型"] :?> Worksheet
            let d12r = ws4.Cells.[12, 4] :?> Range
            printfn "  处理前 D12 HasArray=%b FormulaArray=%s 值=%A 填充=%A" (unbox<bool> d12r.HasArray) (string d12r.FormulaArray) d12r.Value2 d12r.Interior.Color
            let r4 = Publishing.run(true, wb4, "电机额定功率")
            if r4.FunctionMissing then
                printfn "  FunctionMissing=true（本机未加载该 XLL 函数），处理按设计终止，副本未修改。"
            else
                printfn "  run(true) processed=%d ArrayFormula=%d Error=%d Unconform=%d" r4.ProcessedCount r4.ArrayFormulaCells.Length r4.ErrorCells.Length r4.UnconformCells.Length
                printfn "  处理后 D12 FormulaArray=%s 值=%A 填充=%A" (string d12r.FormulaArray) d12r.Value2 d12r.Interior.Color
                for (addr, f) in r4.ArrayFormulaCells do printfn "    跳过数组：%s %s" addr f
                for (addr, t) in r4.ErrorCells do printfn "    跳过错误：%s %s" addr t
            wb4.Close(false) |> ignore
            File.Delete(tmp)
        else
            printfn "  源文件不存在，跳过真实文件测试。"
    with e ->
        printfn "冒烟异常：%s" e.Message
        failed <- 1
    if app <> null then
        try app.Quit() with _ -> ()
    if failed = 0 then printfn "SMOKE OK" else printfn "SMOKE FAILED"
    failed
