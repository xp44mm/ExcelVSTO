open System
open System.IO
open System.Runtime.InteropServices
open System.Threading
open Microsoft.Office.Interop.Excel
open ExcelNumericalMethods

module Native =
    [<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
    extern nativeint FindWindow(string lpClassName, string lpWindowName)
    [<DllImport("user32.dll", SetLastError = true)>]
    extern bool PostMessage(nativeint hWnd, uint32 Msg, nativeint wParam, nativeint lParam)

/// 冒烟测试里 RootsOfEquations.successive 命中拦截分支会弹 MessageBox，
/// 无人值守跑测试会卡住，故开一个后台线程自动关掉弹出的对话框（#32770 是系统对话框类名）。
let runWithDialogClosed (action: unit -> unit) : unit =
    let stop = ref false
    let closer =
        Thread(fun () ->
            let deadline = DateTime.Now.AddSeconds(10.0)
            while not stop.Value && DateTime.Now < deadline do
                let h = Native.FindWindow("#32770", null)
                if h <> 0n then
                    Native.PostMessage(h, 0x0010u, 0n, 0n) |> ignore  // WM_CLOSE
                Thread.Sleep(150))
    closer.IsBackground <- true
    closer.Start()
    try
        action ()
    finally
        stop.Value <- true

/// 模拟发布辅助功能的 Excel COM 冒烟测试：
/// 1) 单格数组公式的写入方式探测（Value2 / ClearContents+Value2 / FormulaArray / Formula2）
/// 2) 更新默认值：不限函数名处理 IFERROR(函数(...), 常量)（内置 SQRT 与自定义 buckling 均处理）、单格数组公式更新、
///    错误值跳过、非常量兜底跳过、非 IFERROR 公式忽略、固定绿色标色
/// 3) 脱公式：不限函数名，固化所有 IFERROR(函数(...), 常量) 单元格；单格数组公式固化、普通公式保留
/// 4) 真实工作簿副本：燃烧器功率计算39.xlsx
/// 5) 包裹自定义函数：公式包裹为 IFERROR(函数(...), 当前真值)，非公式/错误值/已是 IFERROR 不处理

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
        // 重建数组公式再试 Formula2 赋值
        d12.FormulaArray <- "=IFERROR(SQRT(D11),22)"
        try
            Formula2.set d12 "=IFERROR(SQRT(D11),3)"
            printfn "  Formula2 赋值成功 HasArray=%b" (unbox<bool> d12.HasArray)
        with e -> printfn "  Formula2 赋值抛异常：%s" e.Message
        wb1.Close(false) |> ignore

        // ================= Part 2：更新默认值（不限函数名 + 固定绿色标色） =================
        printfn "== Part 2 更新默认值 =="
        let wb2 = app.Workbooks.Add()
        let ws2 = wb2.Worksheets.[1] :?> Worksheet
        ws2.Name <- "设备选型"
        setv ws2 1 1 4.0                      // A1 = 4
        (cellOf ws2 1 2) |> Formula2.setWith "=IFERROR(SQRT(A1),22)"      // B1 内置函数目标，真值 2
        (cellOf ws2 2 2) |> Formula2.setWith "=IFERROR(SQRT(-1),1/0)"     // B2 兜底为错误 → 单元格 #DIV/0!，跳过
        (cellOf ws2 3 2) |> Formula2.setWith "=A1*3"                      // B3 普通公式（非 IFERROR），忽略
        (cellOf ws2 4 2) |> Formula2.setWith "=IFERROR(SQRT(A1),B1)"      // B4 兜底为引用 → 非常量，跳过
        (cellOf ws2 6 2) |> Formula2.setWith "=IFERROR(buckling(C11),200)" // B6 自定义函数名（本机无此函数，值=兜底 200），不限函数名应被处理
        (cellOf ws2 7 2) |> Formula2.setWith "=IFERROR(SQRT(2),99)"       // B7 真值 1.4142135...，兜底按 .NET "0.##" 格式化为 1.41
        setv ws2 11 4 9.0                                        // D11 = 9
        (cellOf ws2 12 4).FormulaArray <- "=IFERROR(SQRT(D11),22)" // D12 单格数组公式，真值 3
        ws2.Range("E5:F6").FormulaArray <- "=SQRT(A1:A2)"        // E5:F6 非 IFERROR 多格数组公式，不进入处理
        let d12dbg = cellOf ws2 12 4
        printfn "  调试 D12.Pattern=%A int=%d xlNone=%d sampleColor=%A" d12dbg.Interior.Pattern (int (d12dbg.Interior.Pattern :?> XlPattern)) (int XlPattern.xlPatternNone) (Publishing.sampleColor wb2)
        let r1 = Publishing.updateDefaults wb2
        printfn "  updateDefaults processed=%d（期望 4：B1、D12、B6、B7）" r1.ProcessedCount
        printfn "  ArrayFormulaCells=%d（期望 0：E5:F6 非 IFERROR 不进入处理）" r1.ArrayFormulaCells.Length
        printfn "  ErrorCells=%d（期望 1：B2 #DIV/0!）" r1.ErrorCells.Length
        printfn "  UnconformCells=%d（期望 1：B4 兜底为引用）" r1.UnconformCells.Length
        expect (r1.ProcessedCount = 4) "processed = 4"
        expect (r1.ArrayFormulaCells.Length = 0) "非 IFERROR 多格数组公式 E5:F6 不进入处理"
        expect (r1.ErrorCells.Length = 1) "错误值 B2 跳过"
        expect (r1.UnconformCells.Length = 1) "非常量兜底 B4 跳过"
        let b1 = cellOf ws2 1 2
        let b4 = cellOf ws2 4 2
        let b6 = cellOf ws2 6 2
        let b7 = cellOf ws2 7 2
        let d12b = cellOf ws2 12 4
        let b3 = cellOf ws2 3 2
        expect (string b1.Formula = "=IFERROR(SQRT(A1),2)") (sprintf "B1 兜底更新：%s" (string b1.Formula))
        expect (string b4.Formula = "=IFERROR(SQRT(A1),B1)") "B4 非常量兜底公式未改动"
        expect (string b6.Formula = "=IFERROR(buckling(C11),200)") (sprintf "B6（buckling）被处理，兜底保持最新值：%s" (string b6.Formula))
        expect (string b7.Formula = "=IFERROR(SQRT(2),1.41)") (sprintf "B7 兜底按 0.## 格式化：%s" (string b7.Formula))
        expect (unbox<bool> d12b.HasArray) "D12 仍为数组公式"
        expect (string d12b.FormulaArray = "=IFERROR(SQRT(D11),3)") (sprintf "D12 兜底更新（FormulaArray 写入）：%s" (string d12b.FormulaArray))
        expect (string b3.Formula = "=A1*3") "B3 普通公式保留"
        expect (b6.Interior.Color = green) (sprintf "B6（buckling）标绿色：%A" b6.Interior.Color)
        expect (b1.Interior.Color = green) (sprintf "B1 标绿色：%A" b1.Interior.Color)
        expect (d12b.Interior.Color = green) (sprintf "D12 标绿色：%A" d12b.Interior.Color)
        expect (b3.Interior.Color <> green) "B3 未被标色"
        // 固定绿色标记测试：标记色与 D12 的填充色无关（先把 D12 改成浅黄，再跑一次新目标 B5 仍标固定绿）
        d12b.Interior.Color <- defaultYellow
        (cellOf ws2 5 2) |> Formula2.setWith "=IFERROR(SQRT(A1),9)"       // B5 新目标
        let r2 = Publishing.updateDefaults wb2
        let b5 = cellOf ws2 5 2
        expect (r2.ProcessedCount = 5) (sprintf "第二次 processed=5（B1/B5/D12/B6/B7）：%d" r2.ProcessedCount)
        expect (b5.Interior.Color = green) (sprintf "B5 标固定绿：%A" b5.Interior.Color)
        expect (b1.Interior.Color = green) (sprintf "B1 也被重新标绿：%A" b1.Interior.Color)
        wb2.Close(false) |> ignore

        // ================= Part 3：脱公式 =================
        printfn "== Part 3 脱公式 =="
        let wb3 = app.Workbooks.Add()
        let ws3 = wb3.Worksheets.[1] :?> Worksheet
        ws3.Name <- "设备选型"
        setv ws3 1 1 4.0
        (cellOf ws3 1 2) |> Formula2.setWith "=IFERROR(SQRT(A1),22)"
        (cellOf ws3 2 2) |> Formula2.setWith "=IFERROR(SQRT(-1),1/0)"
        (cellOf ws3 3 2) |> Formula2.setWith "=A1*3"
        (cellOf ws3 3 3) |> Formula2.setWith "=IFERROR(SQRT(A1),B3)"   // C3 兜底为引用 → 非常量兜底跳过
        setv ws3 11 4 9.0
        (cellOf ws3 12 4).FormulaArray <- "=IFERROR(SQRT(D11),22)"
        ws3.Range("E5:F6").FormulaArray <- "=SQRT(A1:A2)"      // 非 IFERROR 数组公式 → 不进入处理
        let r3 = Publishing.stripFormulas wb3
        printfn "  stripFormulas processed=%d（期望 2：B1、D12）" r3.ProcessedCount
        printfn "  ArrayFormulaCells=%d（期望 0：E5:F6 非 IFERROR 不进入处理）" r3.ArrayFormulaCells.Length
        printfn "  ErrorCells=%d（期望 1）" r3.ErrorCells.Length
        printfn "  UnconformCells=%d（期望 1：C3 兜底为引用）" r3.UnconformCells.Length
        expect (r3.ProcessedCount = 2) "processed = 2"
        expect (r3.ArrayFormulaCells.Length = 0) "非 IFERROR 多格数组公式不进入处理"
        expect (r3.ErrorCells.Length = 1) "错误值跳过"
        expect (r3.UnconformCells.Length = 1) "非常量兜底（引用）跳过"
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
            let r4 = Publishing.updateDefaults wb4
            printfn "  updateDefaults processed=%d ArrayFormula=%d Error=%d Unconform=%d" r4.ProcessedCount r4.ArrayFormulaCells.Length r4.ErrorCells.Length r4.UnconformCells.Length
            printfn "  处理后 D12 FormulaArray=%s Formula=%s 值=%A 填充=%A" (string d12r.FormulaArray) (string d12r.Formula) d12r.Value2 d12r.Interior.Color
            for (addr, f) in r4.ArrayFormulaCells do printfn "    跳过数组：%s %s" addr f
            for (addr, t) in r4.ErrorCells do printfn "    跳过错误：%s %s" addr t
            wb4.Close(false) |> ignore
            File.Delete(tmp)
        else
            printfn "  源文件不存在，跳过真实文件测试。"

        // ================= Part 5：包裹自定义函数 =================
        printfn "== Part 5 包裹自定义函数 =="
        let wb5 = app.Workbooks.Add()
        let ws5 = wb5.Worksheets.[1] :?> Worksheet
        setv ws5 1 1 4.0                                   // A1 = 4
        (cellOf ws5 2 1) |> Formula2.setWith "=SQRT(A1)"            // A2 真值 2 → 应包裹为 =IFERROR(SQRT(A1),2)
        (cellOf ws5 3 1) |> Formula2.setWith "=buckling(B1)"        // A3 自定义函数（本机无此函数 → #NAME?）→ 错误值不包裹
        setv ws5 4 1 5.0                                   // A4 常量，非公式 → 不包裹
        (cellOf ws5 5 1) |> Formula2.setWith "=IFERROR(SQRT(A1),9)" // A5 已是 IFERROR → 不重复包裹
        (cellOf ws5 7 1) |> Formula2.setWith "=SQRT(2)"             // A7 真值 1.4142135... → 兜底按 0.## 格式化为 1.41
        let rw1 = Publishing.wrapFunction (cellOf ws5 2 1)
        printfn "  wrap A2 Wrapped=%b %s" rw1.Wrapped rw1.Message
        expect (rw1.Wrapped) "A2 包裹成功"
        expect (string (cellOf ws5 2 1).Formula = "=IFERROR(SQRT(A1),2)") (sprintf "A2 公式：%s" (string (cellOf ws5 2 1).Formula))
        let rw3 = Publishing.wrapFunction (cellOf ws5 3 1)
        printfn "  wrap A3 Wrapped=%b %s" rw3.Wrapped rw3.Message
        expect (not rw3.Wrapped) "A3 错误值不包裹"
        let rw4 = Publishing.wrapFunction (cellOf ws5 4 1)
        expect (not rw4.Wrapped) "A4 常量（非公式）不包裹"
        let rw5 = Publishing.wrapFunction (cellOf ws5 5 1)
        expect (not rw5.Wrapped) "A5 已是 IFERROR 不重复包裹"
        let rw6 = Publishing.wrapFunction (cellOf ws5 6 1)
        expect (not rw6.Wrapped) "A6 空单元格（非公式）不包裹"
        let rw7 = Publishing.wrapFunction (cellOf ws5 7 1)
        printfn "  wrap A7 Wrapped=%b %s" rw7.Wrapped rw7.Message
        expect (rw7.Wrapped) "A7 包裹成功"
        expect (string (cellOf ws5 7 1).Formula = "=IFERROR(SQRT(2),1.41)") (sprintf "A7 兜底按 0.## 格式化：%s" (string (cellOf ws5 7 1).Formula))
        wb5.Close(false) |> ignore

        // ================= Part 6：数组公式判定（旧式 CSE／新式溢出） =================
        printfn "== Part 6 数组公式判定 =="
        let wb6 = app.Workbooks.Add()
        let ws6 = wb6.Worksheets.[1] :?> Worksheet
        for i in 1 .. 3 do setv ws6 i 1 (float i)
        (cellOf ws6 5 1) |> Formula2.setWith "=A1*2"                    // A5 普通公式
        ws6.Range("C1:C3").FormulaArray <- "=A1:A3*2"                   // C1:C3 旧式 CSE 多格
        (cellOf ws6 1 5) |> Formula2.setWith "=A1:A3*3"                 // E1:E3 新式溢出
        (cellOf ws6 1 7).FormulaArray <- "=SUM(A1:A3)"                  // G1 单格 CSE
        let c1_6, c2_6 = cellOf ws6 1 3, cellOf ws6 2 3
        let e1_6, e2_6 = cellOf ws6 1 5, cellOf ws6 2 5
        let a5_6, g1_6 = cellOf ws6 5 1, cellOf ws6 1 7
        expect (ArrayFormula.hasLegacyArray c1_6 && ArrayFormula.hasLegacyArray c2_6) "旧式 CSE 整片 HasArray"
        expect (not (ArrayFormula.hasSpill c1_6)) "旧式 CSE HasSpill=false"
        expect (ArrayFormula.hasSpill e1_6 && ArrayFormula.hasSpill e2_6) "新式溢出整片 HasSpill"
        expect (not (ArrayFormula.hasLegacyArray e1_6)) "新式溢出 HasArray=false"
        expect (ArrayFormula.isArrayPart e2_6) "溢出区内单元格也算数组公式的一部分"
        expect (ArrayFormula.isArrayMaster c1_6 && not (ArrayFormula.isArrayMaster c2_6)) "旧式 CSE 只有左上角是主格"
        expect (ArrayFormula.isArrayMaster e1_6 && not (ArrayFormula.isArrayMaster e2_6)) "新式溢出只有锚点是主格"
        expect (ArrayFormula.isArrayMaster a5_6) "普通公式恒为主格"
        expect (ArrayFormula.isMultiCellArray c1_6 && ArrayFormula.isMultiCellArray e1_6) "多格数组"
        expect (not (ArrayFormula.isMultiCellArray g1_6)) "单格 CSE 不算多格"
        wb6.Close(false) |> ignore

        // ================= Part 7：更新默认值（新式溢出 + 旧式多格数组） =================
        printfn "== Part 7 更新默认值（数组公式兜底用数组常数） =="
        let wb7 = app.Workbooks.Add()
        let ws7 = wb7.Worksheets.[1] :?> Worksheet
        for i in 1 .. 3 do setv ws7 i 1 (float i)
        (cellOf ws7 1 5) |> Formula2.setWith "=IFERROR(SQRT(A1:A3),0)"          // E1:E3 新式溢出
        ws7.Range("G1:G3").FormulaArray <- "=IFERROR(SQRT(A1:A3),0)"            // G1:G3 旧式 CSE
        let r7 = Publishing.updateDefaults wb7
        printfn "  processed=%d error=%d unconform=%d arraySkipped=%d"
            r7.ProcessedCount r7.ErrorCells.Length r7.UnconformCells.Length r7.ArrayFormulaCells.Length
        let e1_7, e3_7 = cellOf ws7 1 5, cellOf ws7 3 5
        let g1_7, g3_7 = cellOf ws7 1 7, cellOf ws7 3 7
        printfn "  E1.Formula=%s" (string e1_7.Formula)
        printfn "  G1.FormulaArray=%s" (string g1_7.FormulaArray)
        expect (string e1_7.Formula = "=IFERROR(SQRT(A1:A3),{1;1.41;1.73})") (sprintf "溢出区兜底换成数组常数：%s" (string e1_7.Formula))
        expect (ArrayFormula.hasSpill e1_7) "E1 仍是溢出公式"
        expect (abs (unbox<double> e3_7.Value2 - 1.73) < 0.005) (sprintf "E3 值不变：%A" e3_7.Value2)
        expect (string g1_7.FormulaArray = "=IFERROR(SQRT(A1:A3),{1;1.41;1.73})") (sprintf "旧式多格数组兜底换成数组常数：%s" (string g1_7.FormulaArray))
        expect (ArrayFormula.hasLegacyArray g1_7) "G1 仍是旧式数组公式"
        expect (abs (unbox<double> g3_7.Value2 - 1.73) < 0.005) (sprintf "G3 值不变：%A" g3_7.Value2)
        wb7.Close(false) |> ignore

        // ================= Part 8：脱公式（新式溢出 + 旧式多格数组整块固化） =================
        printfn "== Part 8 脱公式（数组公式整块固化） =="
        let wb8 = app.Workbooks.Add()
        let ws8 = wb8.Worksheets.[1] :?> Worksheet
        for i in 1 .. 3 do setv ws8 i 1 (float i)
        (cellOf ws8 1 5) |> Formula2.setWith "=IFERROR(SQRT(A1:A3),0)"
        ws8.Range("G1:G3").FormulaArray <- "=IFERROR(SQRT(A1:A3),0)"
        let r8 = Publishing.stripFormulas wb8
        printfn "  processed=%d error=%d arraySkipped=%d" r8.ProcessedCount r8.ErrorCells.Length r8.ArrayFormulaCells.Length
        let e1_8, e3_8 = cellOf ws8 1 5, cellOf ws8 3 5
        let g1_8, g3_8 = cellOf ws8 1 7, cellOf ws8 3 7
        expect (not (unbox<bool> e1_8.HasFormula)) "E1 已固化"
        expect (abs (unbox<double> e3_8.Value2 - 1.73) < 0.005) (sprintf "E3 固化值：%A" e3_8.Value2)
        expect (not (unbox<bool> g1_8.HasFormula) && not (unbox<bool> g3_8.HasFormula)) "G1:G3 整块已固化"
        expect (abs (unbox<double> g3_8.Value2 - 1.73) < 0.005) (sprintf "G3 固化值：%A" g3_8.Value2)
        wb8.Close(false) |> ignore

        // ================= Part 9：旧式 CSE 一键转换为新式溢出 =================
        printfn "== Part 9 旧式 CSE 转新式溢出 =="
        let wb9 = app.Workbooks.Add()
        let ws9 = wb9.Worksheets.[1] :?> Worksheet
        for i in 1 .. 3 do setv ws9 i 1 (float i)
        ws9.Range("C1:C3").FormulaArray <- "=A1:A3*2"          // 多格 CSE
        (cellOf ws9 1 7).FormulaArray <- "=SUM(A1:A3)"         // 单格 CSE
        let r9 = ArrayFormula.convertCseToSpill wb9
        printfn "  转换 %d 处，失败 %d 处：%s"
            r9.ConvertedAddresses.Length r9.FailedCells.Length (String.Join("、", r9.ConvertedAddresses))
        for (a, m) in r9.FailedCells do printfn "    失败 %s：%s" a m
        let c1_9, c3_9 = cellOf ws9 1 3, cellOf ws9 3 3
        let g1_9 = cellOf ws9 1 7
        expect (r9.ConvertedAddresses.Length = 2) "两处 CSE 都已转换"
        expect (not (ArrayFormula.hasLegacyArray c1_9)) "C1 不再是旧式数组公式"
        expect (ArrayFormula.hasSpill c1_9) "C1 变成溢出公式"
        expect (unbox<double> c3_9.Value2 = 6.0) (sprintf "C3 值不变：%A" c3_9.Value2)
        expect (not (ArrayFormula.hasLegacyArray g1_9) && unbox<double> g1_9.Value2 = 6.0) (sprintf "G1 单格 CSE 已转换：%A" g1_9.Value2)
        wb9.Close(false) |> ignore

        // ================= Part 10：包裹当前工作簿的全部自定义函数 =================
        printfn "== Part 10 包裹全部自定义函数 =="
        let wb10 = app.Workbooks.Add()
        let ws10 = wb10.Worksheets.[1] :?> Worksheet
        for i in 1 .. 3 do setv ws10 i 1 (float i)
        // 用 LAMBDA 定义名称模拟 dna6 自定义函数 myUdf(x) = x*2（本机没有真正的 XLL）
        try
            wb10.Names.Add("myUdf", "=LAMBDA(x, x*2)") |> ignore
        with e -> printfn "  定义 LAMBDA 名称失败：%s" e.Message
        (cellOf ws10 1 3) |> Formula2.setWith "=myUdf(A1)"       // C1 单格自定义函数，真值 2
        (cellOf ws10 1 5) |> Formula2.setWith "=myUdf(A1:A3)"    // E1:E3 溢出自定义函数，真值 {2;4;6}
        (cellOf ws10 1 7) |> Formula2.setWith "=SQRT(A1)"        // G1 内置函数，不该动
        printfn "  包裹前 C1=%A E1=%A E2=%A G1=%A"
            (cellOf ws10 1 3).Value2 (cellOf ws10 1 5).Value2 (cellOf ws10 2 5).Value2 (cellOf ws10 1 7).Value2
        let r10 = Publishing.wrapWorkbook wb10
        printfn "  已包裹 %d，跳过 %d" r10.WrappedCount r10.SkippedCells.Length
        for (_, m) in r10.WrappedCells do printfn "    %s" m
        for (a, m) in r10.SkippedCells do printfn "    跳过 %s：%s" a m
        expect (string (cellOf ws10 1 3).Formula = "=IFERROR(myUdf(A1),2)") (sprintf "C1 已包裹：%s" (string (cellOf ws10 1 3).Formula))
        expect (string (cellOf ws10 1 5).Formula = "=IFERROR(myUdf(A1:A3),{2;4;6})") (sprintf "E1 兜底用数组常数：%s" (string (cellOf ws10 1 5).Formula))
        expect (string (cellOf ws10 1 7).Formula = "=SQRT(A1)") "G1 内置函数未被动"
        wb10.Close(false) |> ignore

        // ================= Part 11：代入法归零拦截数组公式 =================
        printfn "== Part 11 代入法归零拦截数组公式 =="
        let wb11 = app.Workbooks.Add()
        let ws11 = wb11.Worksheets.[1] :?> Worksheet
        setv ws11 1 1 5.0                                        // A1 字面量（旧值）
        (cellOf ws11 2 1) |> Formula2.setWith "=A1*1.4"          // A2 公式（新值 7）
        (cellOf ws11 3 1) |> Formula2.setWith "=A2-A1"           // A3 误差格：应为 A1 <- 7
        runWithDialogClosed (fun () -> RootsOfEquations.successive (cellOf ws11 3 1))
        expect (unbox<double> (cellOf ws11 1 1).Value2 = 7.0) (sprintf "普通情形正常追赶：A1=%A" (cellOf ws11 1 1).Value2)
        // B1:B2 是溢出公式，C1 = B2-B1 应被拦截，B 列一个字都不许动
        (cellOf ws11 1 2) |> Formula2.setWith "={3;4}"
        let b1Value = (cellOf ws11 1 2).Value2
        let b2Value = (cellOf ws11 2 2).Value2
        (cellOf ws11 1 3) |> Formula2.setWith "=B2-B1"
        runWithDialogClosed (fun () -> RootsOfEquations.successive (cellOf ws11 1 3))
        expect ((cellOf ws11 1 2).Value2 = b1Value && (cellOf ws11 2 2).Value2 = b2Value)
            (sprintf "溢出公式区域内的单元格未被改动：B1=%A B2=%A" (cellOf ws11 1 2).Value2 (cellOf ws11 2 2).Value2)
        wb11.Close(false) |> ignore
    with e ->
        printfn "冒烟异常：%s" e.Message
        failed <- 1
    if app <> null then
        try app.Quit() with _ -> ()
    if failed = 0 then printfn "SMOKE OK" else printfn "SMOKE FAILED"
    failed
