namespace ExcelNumericalMethods

open System
open System.Globalization
open Microsoft.Office.Interop.Excel

/// 发布辅助：「脱公式」与「更新默认值」批量处理，以及自定义函数的 IFERROR 包裹。
/// 前两者处理对象一致：不限函数名，凡 =IFERROR(函数(参数...), 常量) 结构（第一参数为函数调用、第二参数为常量）均处理；
/// 更新默认值仅替换兜底常量为最新真值并以绿色标记（当前工作簿上位修改）；
/// 脱公式将目标单元格固化为当前计算结果（在副本上进行）。
///
/// 数组公式（新旧两式）一视同仁地处理：多格区域整块读写，
/// 旧式 CSE 走 FormulaArray，新式溢出只写锚点，区域内部单元格一律跳过（见 ArrayFormula）。
module Publishing =

    /// 批量处理结果
    type RunResult =
        { /// 已处理的目标单元格数
          ProcessedCount: int
          /// 当前为错误值而跳过的单元格（工作表!地址, 显示文本）
          ErrorCells: (string * string) array
          /// 不符合 IFERROR(函数(...), 常量/兜底) 处理条件而跳过的单元格（工作表!地址, 公式）
          UnconformCells: (string * string) array
          /// 数组公式（旧式 CSE 区域 / 新式溢出区）写入失败而跳过的单元格（工作表!地址, 公式）。
          /// 正常路径下数组公式是整块处理的：多格按块读写、单格按普通单元格处理，只有写入抛异常时才记到这里
          ArrayFormulaCells: (string * string) array
          /// 受保护且无法解除保护而跳过的工作表
          ProtectedSheets: string array }

    let emptyResult =
        { ProcessedCount = 0
          ErrorCells = [||]
          UnconformCells = [||]
          ArrayFormulaCells = [||]
          ProtectedSheets = [||] }

    /// 包裹单个单元格的结果
    type WrapResult =
        { /// 是否已包裹成功
          Wrapped: bool
          /// 单元格地址（工作表名!列字母行号）
          Address: string
          /// 状态栏提示文本（成功时含新公式，失败时含原因）
          Message: string }

    /// 列号转列字母：1 -> A，27 -> AA，703 -> AAA
    let rec columnLetters (n: int) : string =
        let n' = n - 1
        let c = char (int 'A' + n' % 26)
        if n' < 26 then string c
        else columnLetters (n' / 26) + string c

    /// 单元格地址：工作表名!列字母行号
    let cellAddress (ws: Worksheet) (cell: Range) : string =
        sprintf "%s!%s%d" ws.Name (columnLetters cell.Column) cell.Row

    /// 工作表中全部含公式的单元格（SpecialCells 优化；无公式时为空序列）。
    /// 实现见 Traversal.getFormulaCells（新式溢出公式只有锚点会出现，旧式 CSE 则整片区域都会出现）。
    let getFormulaCells (ws: Worksheet) : seq<Range> =
        Traversal.getFormulaCells ws

    /// 字符是否属于 Excel 函数名字符（字母、数字、下划线、点）。
    /// 用于识别 IFERROR 第一参数开头的函数名（isFunctionCallArg）。
    let isFunctionNameChar (c: char) : bool =
        Char.IsLetterOrDigit c || c = '_' || c = '.'

    /// 公式是否为 =IFERROR(...) 结构（不区分大小写）。
    /// 更新默认值与脱公式以此作预过滤：不限函数名，凡 IFERROR 公式都进入结构判定。
    let isIfErrorFormula (cell: Range) : bool =
        let formula = string (cell.Formula)
        formula.TrimStart().StartsWith("=IFERROR(", StringComparison.OrdinalIgnoreCase)

    /// IFERROR 第一参数是否为以函数调用开头的文本（函数名 + 左括号），
    /// 如 buckling(C11)、IF(A1>0,1,0)；单元格引用、括号表达式、跨表引用等开头均非函数调用。
    /// 容忍 Excel 365 自动插入的隐式交集运算符 @ 前缀（如 =IFERROR(@函数(...), 常量)）。
    let isFunctionCallArg (arg1: string) : bool =
        let t0 = arg1.TrimStart()
        let t =
            if t0.StartsWith("@", StringComparison.Ordinal) then t0.Substring(1).TrimStart()
            else t0
        if t.Length = 0 || (not (Char.IsLetter t.[0]) && t.[0] <> '_') then false
        else
            let rec nameEnd i =
                if i < t.Length && isFunctionNameChar t.[i] then nameEnd (i + 1)
                else i
            let n = nameEnd 0
            n < t.Length && t.[n] = '('

    /// 数组公式区域是否仅一个单元格。
    /// 单格数组公式可安全处理（写入用 FormulaArray 保持数组属性），多格数组公式跳过避免破坏。
    let isSingleCellArray (cell: Range) : bool =
        let arr = cell.CurrentArray
        arr.Rows.Count = 1 && arr.Columns.Count = 1

    /// 默认样板色：浅黄 FFEB9C（Excel COM 颜色为 BGR 顺序，即 0x009CEBFF）
    let [<Literal>] DefaultSampleColor = 0x009CEBFF

    /// 更新默认值的标记色：绿色（RGB(146,208,80) 以 BGR 顺序编码为 0x0050D092）
    let [<Literal>] UpdateDefaultsColor = 0x0050D092

    /// 取样板背景色：工作簿存在「设备选型!D12」且该单元格有背景填充时用其颜色，否则用默认浅黄。
    /// 更新默认值的已更新标记色现为固定绿色 UpdateDefaultsColor，此函数仅供调试取样色。
    let sampleColor (wb: Workbook) : float =
        let sheet =
            Traversal.getWorksheets wb
            |> Seq.tryFind (fun w -> w.Name = "设备选型")
        match sheet with
        | None -> float DefaultSampleColor
        | Some w ->
            try
                let cell = w.Cells.[12, 4] :?> Range
                let pattern = cell.Interior.Pattern :?> XlPattern
                if pattern = XlPattern.xlPatternNone then float DefaultSampleColor
                else
                    match box cell.Interior.Color with
                    | :? float as f -> f
                    | :? int as i -> float i
                    | _ -> float DefaultSampleColor
            with _ -> float DefaultSampleColor

    /// 是否为合并区域的左上角单元格（合并区域仅处理左上角，其余跳过）
    let isMergeTopLeft (cell: Range) : bool =
        if not (unbox cell.MergeCells) then true
        else
            let area = cell.MergeArea
            cell.Row = area.Row && cell.Column = area.Column

    /// Excel COM 错误码转错误显示文本
    let errorText (n: int) : string =
        match n with
        | -2146826288 -> "#NULL!"
        | -2146826281 -> "#DIV/0!"
        | -2146826273 -> "#VALUE!"
        | -2146826265 -> "#REF!"
        | -2146826259 -> "#NAME?"
        | -2146826252 -> "#NUM!"
        | -2146826246 -> "#N/A"
        | _ -> sprintf "#ERR(%d)" n

    /// 单元格当前是否为错误值。
    /// 本机 Excel 的 COM 互操作读取错误单元格不抛异常，而是返回负 Int32 错误码（如 -2146826259 为 #NAME?），
    /// 正常单元格经 Value2 读取只会是 double / string / bool，不会出现负 Int32。
    let isErrorCell (cell: Range) : bool =
        try
            match cell.Value2 with
            | :? int as n -> n < 0
            | _ -> false
        with _ -> true

    /// 解析 =IFERROR(参数1, 参数2) 公式，返回 (参数1, 参数2)。
    /// 按括号配对找第一个顶层逗号（跳过嵌套函数与字符串字面量内的逗号），避免误切；
    /// 不符合结构时返回 None。
    let splitIfError (formula: string) : (string * string) option =
        let s = formula.Trim()
        if s.Length = 0 || s.[0] <> '=' then None
        else
            let body = s.Substring(1).TrimStart()
            let nameLen = "IFERROR".Length
            if body.Length < nameLen then None
            else
                let head = body.Substring(0, nameLen)
                if not (String.Equals(head, "IFERROR", StringComparison.OrdinalIgnoreCase)) then None
                else
                    let rest = body.Substring(nameLen).TrimStart()
                    if rest.Length = 0 || rest.[0] <> '(' then None
                    else
                        let mutable depth = 0
                        let mutable comma = -1
                        let mutable close = -1
                        let mutable i = 0
                        let mutable inString = false
                        let n = rest.Length
                        // 找到第一个顶层逗号与配对结尾括号后停止
                        while i < n && (comma < 0 || close < 0) do
                            let c = rest.[i]
                            if inString then
                                // 字符串字面量内："" 是转义的双引号
                                if c = '"' then
                                    if i + 1 < n && rest.[i + 1] = '"' then i <- i + 1
                                    else inString <- false
                            else
                                match c with
                                | '"' -> inString <- true
                                | '(' -> depth <- depth + 1
                                | ')' ->
                                    depth <- depth - 1
                                    if depth = 0 && close < 0 then close <- i
                                | ',' when depth = 1 && comma < 0 -> comma <- i
                                | _ -> ()
                            i <- i + 1
                        if comma < 0 || close < 0 || close < comma then None
                        else
                            let arg1 = rest.Substring(1, comma - 1).Trim()
                            let arg2 = rest.Substring(comma + 1, close - comma - 1).Trim()
                            if arg1.Length = 0 || arg2.Length = 0 then None
                            else Some(arg1, arg2)

    /// 判断 IFERROR 第二参数是否为常量字面量：
    /// 数字（含小数、指数、正负号、百分号）、带双引号的字符串、TRUE/FALSE。
    /// 单元格引用、命名区域、函数调用、运算符表达式等均视为非常量（返回 false）。
    let isConstantLiteral (s: string) : bool =
        let t = s.Trim()
        if t.Length = 0 then false
        elif t.StartsWith("\"", StringComparison.Ordinal) then
            // 字符串字面量：以双引号开始、以双引号结束
            t.Length >= 2 && t.EndsWith("\"", StringComparison.Ordinal)
        elif String.Equals(t, "TRUE", StringComparison.OrdinalIgnoreCase)
             || String.Equals(t, "FALSE", StringComparison.OrdinalIgnoreCase) then true
        else
            // 百分号是合法数字字面量后缀（如 50%），先去掉再解析
            let body =
                if t.EndsWith("%", StringComparison.Ordinal) then t.Substring(0, t.Length - 1)
                else t
            match Double.TryParse(body, NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, _ -> true
            | _ -> false

    /// 将真值格式化为公式字面量：
    /// 文本加双引号（内部双引号转义为两个双引号），布尔为 TRUE/FALSE。
    /// 数字用 .NET 的 "0.##" 格式：最多两位小数、四舍五入、去掉尾随零与多余小数点
    /// （如 0.169584876486935 -> 0.17、2.0 -> 2）。注意这是 .NET 数字格式，与 Excel 单元格数字格式行为不同。
    let formatLiteral (value: obj) : string =
        match value with
        | null -> "\"\""
        | :? string as s -> Quotation.quote s
        | :? bool as b -> if b then "TRUE" else "FALSE"
        | :? float as f -> f.ToString("0.##", CultureInfo.InvariantCulture)
        | :? decimal as d -> d.ToString("0.##", CultureInfo.InvariantCulture)
        | :? DateTime as dt ->
            // 使用 .Value 读取日期时出现；用 DATE 函数保证日期语义
            let datePart = sprintf "DATE(%d,%d,%d)" dt.Year dt.Month dt.Day
            if dt.TimeOfDay = TimeSpan.Zero then datePart
            else
                let frac = dt.TimeOfDay.TotalSeconds / 86400.0
                sprintf "%s+%s" datePart (frac.ToString("R", CultureInfo.InvariantCulture))
        | :? IConvertible as c -> c.ToString(CultureInfo.InvariantCulture)
        | other -> Quotation.quote (string other)

    /// 把二维值数组格式化为 Excel 数组常数，如 {1,2;3,4}：逗号分列、分号换行。
    /// 多格数组公式（旧式 CSE 区域 / 新式溢出区）的兜底值必须用同形数组常数，
    /// 否则兜底值的形状与区域对不上，出错时只会填出一个格子。
    let formatArrayLiteral (values: obj[,]) : string =
        let r0 = values.GetLowerBound 0
        let c0 = values.GetLowerBound 1
        let rows =
            [ for r in 0 .. values.GetLength 0 - 1 ->
                [ for c in 0 .. values.GetLength 1 - 1 -> formatLiteral values.[r0 + r, c0 + c] ]
                |> String.concat "," ]
        "{" + String.concat ";" rows + "}"

    /// 读取单元格当前值：错误值返回 Choice2Of2 显示文本，正常值返回 Choice1Of2。
    let private cellValue (cell: Range) : Choice<obj, string> =
        try
            match cell.Value2 with
            | :? int as n when n < 0 -> Choice2Of2 (errorText n)
            | v -> Choice1Of2 v
        with _ -> Choice2Of2 "#ERR"

    /// 公式中是否含有自定义函数（Excel-DNA / XLL / VBA 宏表函数等）。
    /// 判定口径见 BuiltinFunctions.hasCustomFunction。
    let hasCustomFunction (formula: string) : bool =
        BuiltinFunctions.hasCustomFunction formula

    /// 批量处理工作簿：不限函数名，处理所有 =IFERROR(函数(...), 常量) 结构的单元格
    /// （第一参数为函数调用、第二参数为常量）。
    /// updateDefaults = true 时更新默认值：保留 IFERROR 结构，将兜底常量替换为最新真值，并以绿色标记；
    /// updateDefaults = false 时脱公式：将目标单元格固化为当前计算结果（清除公式）。
    let run (updateDefaults: bool, wb: Workbook) : RunResult =
        wb.Application.Calculate()
        let mutable processed = 0
        let errors = ResizeArray<string * string>()
        let unconform = ResizeArray<string * string>()
        let arrays = ResizeArray<string * string>()
        let protectedSheets = ResizeArray<string>()
        for ws in Traversal.getWorksheets wb do
            let address = cellAddress ws
            let isProtected = unbox ws.ProtectContents
            if isProtected then
                // 尝试无密码解除保护（副本上处理）
                try ws.Unprotect()
                with _ -> ()
            if unbox ws.ProtectContents then
                // 解除保护失败，整表跳过并记录
                protectedSheets.Add ws.Name
            else
                for cell in getFormulaCells ws do
                    // 不限函数名：凡 IFERROR 公式都进入结构判定（更新默认值与脱公式一致）
                    if isIfErrorFormula cell then
                        if not (ArrayFormula.isArrayMaster cell) then
                            // 数组公式区域内部单元格：由区域主单元格统一处理，单独改会破坏数组
                            ()
                        elif not (isMergeTopLeft cell) then
                            () // 合并区域仅处理左上角
                        else
                            let formula = string (cell.Formula)
                            // 先查错误值（取不到真值就不必再看结构），与结构判定分两步：
                            // 多格数组公式要整块查，任一格出错就跳过整块。
                            let errorInBlock : (Range * string) option =
                                if ArrayFormula.isMultiCellArray cell then
                                    match ArrayFormula.arrayRange cell with
                                    | Some range ->
                                        Traversal.getCellsOfRange range
                                        |> Seq.tryPick (fun c ->
                                            match cellValue c with
                                            | Choice2Of2 text -> Some(c, text)
                                            | _ -> None)
                                    | None -> None
                                else
                                    match cellValue cell with
                                    | Choice2Of2 text -> Some(cell, text)
                                    | _ -> None
                            match errorInBlock with
                            | Some (b, text) -> errors.Add(address b, text)
                            | None ->
                                match splitIfError formula with
                                | Some (arg1, arg2) when
                                    isFunctionCallArg arg1
                                    && isConstantLiteral arg2 ->
                                    if ArrayFormula.isMultiCellArray cell then
                                        // ---- 多格数组公式（旧式 CSE 区域 / 新式溢出区）：整块读取、整块写回 ----
                                        match ArrayFormula.arrayRange cell with
                                        | None -> unconform.Add(address cell, formula)
                                        | Some range ->
                                            let values = range.Value2 :?> obj[,]
                                            try
                                                if updateDefaults then
                                                    // 兜底值换成同形数组常数；旧式用 FormulaArray 整块写入保持数组属性
                                                    let newFormula = sprintf "=IFERROR(%s,%s)" arg1 (formatArrayLiteral values)
                                                    if ArrayFormula.hasLegacyArray cell then range.FormulaArray <- newFormula
                                                    else Formula2.set cell newFormula
                                                    cell.Interior.Color <- float UpdateDefaultsColor
                                                else
                                                    // 脱公式：清掉整块数组公式（溢出公式清锚点即可），再把当前结果写回原区域
                                                    (if ArrayFormula.hasLegacyArray cell then range else cell).ClearContents() |> ignore
                                                    range.Value2 <- values
                                                processed <- processed + 1
                                            with _ ->
                                                arrays.Add(address cell, formula)
                                    else
                                        // ---- 单格（含单格旧式 CSE 数组公式）----
                                        match cellValue cell with
                                        | Choice2Of2 text -> errors.Add(address cell, text)
                                        | Choice1Of2 v ->
                                            try
                                                if updateDefaults then
                                                    // 更新默认值：保留第一个参数，兜底值替换为最新真值；
                                                    // 单格数组公式用 FormulaArray 写入以保持数组公式属性
                                                    let newFormula = sprintf "=IFERROR(%s,%s)" arg1 (formatLiteral v)
                                                    if unbox cell.HasArray then cell.FormulaArray <- newFormula
                                                    else Formula2.set cell newFormula
                                                    // 以绿色标记已更新的目标单元格
                                                    cell.Interior.Color <- float UpdateDefaultsColor
                                                else
                                                    // 脱公式：固化当前计算结果，清除公式
                                                    cell.Value2 <- v
                                                processed <- processed + 1
                                            with _ ->
                                                arrays.Add(address cell, formula)
                                | _ ->
                                    unconform.Add(address cell, formula)
                if isProtected then
                    // 重新保护副本，与源工作表保持一致
                    try ws.Protect()
                    with _ -> ()
        { ProcessedCount = processed
          ErrorCells = errors.ToArray()
          UnconformCells = unconform.ToArray()
          ArrayFormulaCells = arrays.ToArray()
          ProtectedSheets = protectedSheets.ToArray() }

    /// 更新默认值：对当前工作簿在位修改，处理所有 =IFERROR(函数(...), 常量) 结构的单元格，
    /// 将兜底常量替换为最新真值并以绿色标记。不限函数名（内置、XLL、自定义函数均适用）。
    let updateDefaults (wb: Workbook) : RunResult =
        run (true, wb)

    /// 脱公式：处理所有 =IFERROR(函数(...), 常量) 结构的单元格，固化为当前计算结果（清除公式）。
    /// 不限函数名；通常用于副本工作簿（另存副本后处理）。
    let stripFormulas (wb: Workbook) : RunResult =
        run (false, wb)

    /// 包裹一个公式单元格：=函数(...) 包裹为 =IFERROR(函数(...), 兜底值)。
    ///
    /// 兜底值取当前计算结果，按 .NET "0.##" 格式化；
    /// 多格数组公式（旧式 CSE 区域 / 新式溢出区）取整个区域的值，兜底值用同形数组常数（如 {1,2;3,4}），
    /// 单格保持标量兜底值。旧式数组公式用 FormulaArray 整块写入以保持数组属性，新式溢出写锚点即可。
    ///
    /// 非公式、已是 IFERROR、处于数组区域内部（非主单元格）、错误值时返回 Wrapped=false 及原因。
    let tryWrapCell (cell: Range) : WrapResult =
        let ws = cell.Worksheet
        let addr = cellAddress ws cell
        if not (unbox cell.HasFormula) then
            { Wrapped = false; Address = addr; Message = "不是公式，未包裹。" }
        elif isIfErrorFormula cell then
            { Wrapped = false; Address = addr; Message = "已是 IFERROR 公式，无需重复包裹。" }
        elif not (ArrayFormula.isArrayMaster cell) then
            { Wrapped = false; Address = addr; Message = "处于数组公式区域内部，须改区域主单元格，未包裹。" }
        elif ArrayFormula.isMultiCellArray cell then
            match ArrayFormula.arrayRange cell with
            | None ->
                { Wrapped = false; Address = addr; Message = "取不到数组公式区域，未包裹。" }
            | Some range ->
                let bad =
                    Traversal.getCellsOfRange range
                    |> Seq.tryPick (fun c ->
                        match cellValue c with
                        | Choice2Of2 text -> Some(c, text)
                        | _ -> None)
                match bad with
                | Some (b, text) ->
                    { Wrapped = false
                      Address = addr
                      Message = sprintf "数组区域内 %s 为错误值（%s），无法取兜底值，未包裹。" (cellAddress ws b) text }
                | None ->
                    let formula = string (cell.Formula)
                    let body = formula.TrimStart().TrimStart('=')
                    let newFormula = sprintf "=IFERROR(%s,%s)" body (formatArrayLiteral (range.Value2 :?> obj[,]))
                    try
                        if ArrayFormula.hasLegacyArray cell then range.FormulaArray <- newFormula
                        else Formula2.set cell newFormula
                        { Wrapped = true
                          Address = addr
                          Message = sprintf "已包裹 %s 为 %s" addr newFormula }
                    with ex ->
                        { Wrapped = false; Address = addr; Message = sprintf "写入失败：%s" ex.Message }
        else
            match cellValue cell with
            | Choice2Of2 text ->
                { Wrapped = false
                  Address = addr
                  Message = sprintf "当前单元格为错误值（%s），无法取兜底值，未包裹。" text }
            | Choice1Of2 v ->
                let formula = string (cell.Formula)
                let body = formula.TrimStart().TrimStart('=')
                let newFormula = sprintf "=IFERROR(%s,%s)" body (formatLiteral v)
                try
                    if unbox cell.HasArray then cell.FormulaArray <- newFormula
                    else Formula2.set cell newFormula
                    { Wrapped = true
                      Address = addr
                      Message = sprintf "已包裹 %s 为 %s" addr newFormula }
                with ex ->
                    { Wrapped = false; Address = addr; Message = sprintf "写入失败：%s" ex.Message }

    /// 包裹自定义函数：把单个单元格公式 =函数(...) 包裹为 =IFERROR(函数(...), 当前真值)。
    /// 详细规则见 tryWrapCell。
    let wrapFunction (cell: Range) : WrapResult =
        tryWrapCell cell

    /// 工作簿批量包裹结果
    type WrapRunResult =
        { /// 已包裹的单元格数
          WrappedCount: int
          /// 已包裹的（工作表!地址, 新公式）
          WrappedCells: (string * string) array
          /// 含自定义函数但未能包裹的（工作表!地址, 原因）
          SkippedCells: (string * string) array
          /// 受保护且无法解除保护而跳过的工作表
          ProtectedSheets: string array }

    /// 包裹当前工作簿中所有含自定义函数（Excel-DNA / XLL）的公式：
    /// =函数(...) 包裹为 =IFERROR(函数(...), 兜底值)。
    ///
    /// 与单个单元格版的区别在范围与兜底值形状：
    ///   - 范围：扫描当前工作簿全部工作表，只动含自定义函数的公式（内置函数公式原样保留）；
    ///   - 兜底值：多格数组公式（旧式 CSE 区域 / 新式溢出区）用同形数组常数，单格用标量。
    /// 已是 IFERROR、合并区域非左上角、数组区域内部、错误值、受保护工作表均跳过。
    let wrapWorkbook (wb: Workbook) : WrapRunResult =
        let wrapped = ResizeArray<string * string>()
        let skipped = ResizeArray<string * string>()
        let protectedSheets = ResizeArray<string>()
        try
            wb.Application.Calculate()
        with _ -> ()
        for ws in Traversal.getWorksheets wb do
            let wasProtected = unbox ws.ProtectContents
            if wasProtected then
                try ws.Unprotect()
                with _ -> ()
            if unbox ws.ProtectContents then
                protectedSheets.Add ws.Name
            else
                for cell in getFormulaCells ws |> Seq.toList do
                    let formula = string (cell.Formula)
                    if isIfErrorFormula cell then
                        () // 已经是 IFERROR，不重复包裹
                    elif not (isMergeTopLeft cell) then
                        () // 合并区域仅处理左上角
                    elif not (ArrayFormula.isArrayMaster cell) then
                        () // 数组区域内部单元格由主单元格统一处理
                    elif hasCustomFunction formula then
                        let result = tryWrapCell cell
                        if result.Wrapped then
                            wrapped.Add(result.Address, result.Message)
                        else
                            skipped.Add(result.Address, result.Message)
            if wasProtected then
                try ws.Protect()
                with _ -> ()
        { WrappedCount = wrapped.Count
          WrappedCells = wrapped.ToArray()
          SkippedCells = skipped.ToArray()
          ProtectedSheets = protectedSheets.ToArray() }

    /// 生成处理结果提示文本
    let formatSummary (updateDefaults: bool, path: string, result: RunResult) : string =
        let sb = Text.StringBuilder()
        let action = if updateDefaults then "更新默认值" else "脱公式"
        sb.AppendLine(sprintf "%s完成。" action) |> ignore
        sb.AppendLine(sprintf "已生成副本：%s" path) |> ignore
        sb.AppendLine(sprintf "已处理目标单元格 %d 个。" result.ProcessedCount) |> ignore
        if result.ErrorCells.Length > 0 then
            sb.AppendLine(sprintf "跳过错误值单元格 %d 个，请人工检查：" result.ErrorCells.Length) |> ignore
            for (addr, text) in result.ErrorCells do
                sb.AppendLine(sprintf "  %s（%s）" addr text) |> ignore
        if result.UnconformCells.Length > 0 then
            let structDesc =
                if updateDefaults then "IFERROR(函数(...), 常量)"
                else "IFERROR(函数(...), 兜底)"
            sb.AppendLine(sprintf "跳过不符合 %s 结构的单元格 %d 个：" structDesc result.UnconformCells.Length) |> ignore
            for (addr, formula) in result.UnconformCells do
                sb.AppendLine(sprintf "  %s：%s" addr formula) |> ignore
        if result.ArrayFormulaCells.Length > 0 then
            sb.AppendLine(sprintf "数组公式写入失败而跳过的单元格 %d 个：" result.ArrayFormulaCells.Length) |> ignore
            for (addr, formula) in result.ArrayFormulaCells do
                sb.AppendLine(sprintf "  %s：%s" addr formula) |> ignore
        if result.ProtectedSheets.Length > 0 then
            sb.AppendLine(sprintf "受保护且无法解除保护而跳过的工作表 %d 个：%s" result.ProtectedSheets.Length (String.concat "、" result.ProtectedSheets)) |> ignore
        sb.ToString().TrimEnd()
