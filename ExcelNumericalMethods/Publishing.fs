namespace ExcelNumericalMethods

open System
open System.Globalization
open Microsoft.Office.Interop.Excel

/// 发布辅助：「脱公式」与「更新默认值」批量处理。
/// 处理对象统一为 =IFERROR(函数(参数...), 兜底值) 结构；
/// 更新默认值不限函数名：处理所有该结构单元格，且兜底值须为常量字面量，仅替换该常量并以绿色标记；
/// 脱公式按 marker 指定的函数名过滤并在副本上进行，更新默认值在当前工作簿上位修改。
module Publishing =

    /// 批量处理结果
    type RunResult =
        { /// 已处理的目标单元格数
          ProcessedCount: int
          /// 本机无此函数（ERROR.TYPE 探测为 #NAME?），处理已终止
          FunctionMissing: bool
          /// 当前为错误值而跳过的单元格（工作表!地址, 显示文本）
          ErrorCells: (string * string) array
          /// 不符合 IFERROR(函数(...), 常量/兜底) 处理条件而跳过的单元格（工作表!地址, 公式）
          UnconformCells: (string * string) array
          /// 多格数组公式而跳过的单元格（工作表!地址, 公式）；单格数组公式按普通单元格处理
          ArrayFormulaCells: (string * string) array
          /// 受保护且无法解除保护而跳过的工作表
          ProtectedSheets: string array }

    let emptyResult =
        { ProcessedCount = 0
          FunctionMissing = false
          ErrorCells = [||]
          UnconformCells = [||]
          ArrayFormulaCells = [||]
          ProtectedSheets = [||] }

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
    /// SpecialCells 的结果可能是多个不连续区域，必须按区域逐个遍历，否则取不到全部单元格。
    let getFormulaCells (ws: Worksheet) : seq<Range> =
        try
            let used = ws.UsedRange
            let formulas = used.SpecialCells(XlCellType.xlCellTypeFormulas)
            seq {
                for a in 1 .. formulas.Areas.Count do
                    let area = formulas.Areas.[a]
                    yield! Traversal.getCellsOfRange area
            }
        with _ -> Seq.empty

    /// 字符是否属于 Excel 函数名字符（字母、数字、下划线、点）。
    /// 用于判断 marker 是否以完整函数名出现，避免子串误匹配（如 marker="RT" 误中 SQRT）。
    let isFunctionNameChar (c: char) : bool =
        Char.IsLetterOrDigit c || c = '_' || c = '.'

    /// 文本中是否出现以 marker 为函数名的调用（不区分大小写）。
    /// marker 是符合 Excel 函数名称格式的通用文本：内置函数（SQRT、SUM…）、XLL 函数、
    /// 自定义函数（电机额定功率…）均可，不绑定任何具体函数。
    /// 仅当 marker 后紧跟左括号（函数调用形式）且其前一个字符不是函数名字符（完整函数名边界）时视为命中。
    let hasFunctionCall (marker: string) (text: string) : bool =
        if String.IsNullOrEmpty marker then false
        else
            let pat = marker + "("
            let rec loop (start: int) =
                let i = text.IndexOf(pat, start, StringComparison.OrdinalIgnoreCase)
                if i < 0 then false
                elif i = 0 || not (isFunctionNameChar text.[i - 1]) then true
                else loop (i + 1)
            loop 0

    /// 单元格公式文本中是否出现以 marker 为函数名的调用
    let containsMarker (marker: string) (cell: Range) : bool =
        hasFunctionCall marker (string (cell.Formula))

    /// 公式是否为 =IFERROR(...) 结构（不区分大小写）。
    /// 更新默认值以此作预过滤：不限函数名，凡 IFERROR 公式都进入结构判定。
    let isIfErrorFormula (cell: Range) : bool =
        let formula = string (cell.Formula)
        formula.TrimStart().StartsWith("=IFERROR(", StringComparison.OrdinalIgnoreCase)

    /// IFERROR 第一参数是否为以函数调用开头的文本（函数名 + 左括号），
    /// 如 buckling(C11)、IF(A1>0,1,0)；单元格引用、括号表达式、跨表引用等开头均非函数调用。
    let isFunctionCallArg (arg1: string) : bool =
        let t = arg1.TrimStart()
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
    /// 文本加双引号（内部双引号转义为两个双引号），数字按不变区域设置输出，布尔为 TRUE/FALSE。
    let formatLiteral (value: obj) : string =
        match value with
        | null -> "\"\""
        | :? string as s -> Quotation.quote s
        | :? bool as b -> if b then "TRUE" else "FALSE"
        | :? float as f -> f.ToString("R", CultureInfo.InvariantCulture)
        | :? decimal as d -> d.ToString(CultureInfo.InvariantCulture)
        | :? DateTime as dt ->
            // 使用 .Value 读取日期时出现；用 DATE 函数保证日期语义
            let datePart = sprintf "DATE(%d,%d,%d)" dt.Year dt.Month dt.Day
            if dt.TimeOfDay = TimeSpan.Zero then datePart
            else
                let frac = dt.TimeOfDay.TotalSeconds / 86400.0
                sprintf "%s+%s" datePart (frac.ToString("R", CultureInfo.InvariantCulture))
        | :? IConvertible as c -> c.ToString(CultureInfo.InvariantCulture)
        | other -> Quotation.quote (string other)

    /// 预检（只读，不修改工作簿）：
    /// 统计目标单元格数，并判断是否全部为错误值（可能本机无此标记函数）。
    /// 返回 (目标单元格数, 是否全部为错误值)。
    let precheck (wb: Workbook, marker: string) : int * bool =
        wb.Application.Calculate()
        let mutable count = 0
        let mutable errorCount = 0
        for ws in Traversal.getWorksheets wb do
            for cell in getFormulaCells ws do
                if containsMarker marker cell then
                    count <- count + 1
                    if isMergeTopLeft cell && isErrorCell cell then errorCount <- errorCount + 1
        count, (count > 0 && errorCount = count)

    /// 探测 marker 所指函数是否可用：取第一个目标单元格的 IFERROR 第一参数，
    /// 用 =ERROR.TYPE(第一参数) 求值，结果为 5（#NAME?）即函数名不存在。
    /// 返回 null 表示找不到可探测的目标（无法判断），true 表示函数可用，false 表示本机无此函数。
    /// （用 Nullable<bool> 以便 C# 调用方以 HasValue/Value 访问。）
    let checkFunctionExists (wb: Workbook, marker: string) : System.Nullable<bool> =
        let probeArg =
            Traversal.getWorksheets wb
            |> Seq.tryPick (fun ws ->
                getFormulaCells ws
                |> Seq.tryPick (fun cell ->
                    if containsMarker marker cell && isMergeTopLeft cell then
                        match splitIfError (string (cell.Formula)) with
                        | Some (arg1, _) when hasFunctionCall marker arg1 ->
                            Some arg1
                        | _ -> None
                    else None))
        match probeArg with
        | None -> System.Nullable()
        | Some arg1 ->
            try
                let v = wb.Application.Evaluate("=ERROR.TYPE(" + arg1 + ")")
                match v with
                | :? float as f -> System.Nullable(f <> 5.0)
                | _ -> System.Nullable(true)
            with _ -> System.Nullable(true)

    /// 批量处理工作簿：
    /// updateDefaults = true 时更新默认值：不限函数名，处理所有 =IFERROR(函数(...), 常量) 结构，
    /// 保留第一参数，仅当第二参数为常量时替换兜底值为最新真值，并以绿色标记；
    /// updateDefaults = false 时脱公式：按 marker 过滤目标函数，将目标单元格固化为当前计算结果。
    let run (updateDefaults: bool, wb: Workbook, marker: string) : RunResult =
        wb.Application.Calculate()
        // 更新默认值不限函数名，无需探测函数存在性；脱公式按 marker 过滤，
        // 处理前检查 marker 所指函数是否存在，不存在则终止，避免误固化
        let blocked =
            not updateDefaults
            && (let exists = checkFunctionExists (wb, marker)
                in exists.HasValue && not exists.Value)
        if blocked then
            { emptyResult with FunctionMissing = true }
        else
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
                    // 更新默认值不限函数名，凡 IFERROR 公式都进入结构判定；脱公式按 marker 过滤
                    let isTarget =
                        if updateDefaults then isIfErrorFormula cell
                        else containsMarker marker cell
                    if isTarget then
                        if unbox cell.HasArray && not (isSingleCellArray cell) then
                            // 多格数组公式跳过，避免破坏；单格数组公式按普通单元格处理
                            arrays.Add(address cell, string (cell.Formula))
                        elif not (isMergeTopLeft cell) then
                            () // 合并区域仅处理左上角
                        else
                            // 读取当前计算结果（真值）；错误值跳过并记录
                            let value =
                                try
                                    match cell.Value2 with
                                    | :? int as n when n < 0 -> Choice2Of2 (errorText n)
                                    | v -> Choice1Of2 v
                                with _ -> Choice2Of2 "#ERR"
                            match value with
                            | Choice2Of2 text -> errors.Add(address cell, text)
                            | Choice1Of2 v ->
                                if updateDefaults then
                                    let formula = string (cell.Formula)
                                    match splitIfError formula with
                                    | Some (arg1, arg2) when
                                        isFunctionCallArg arg1
                                        && isConstantLiteral arg2 ->
                                        // 仅当第二参数为常量时处理：保留第一个参数，兜底值替换为最新真值；
                                        // 单格数组公式用 FormulaArray 写入以保持数组公式属性
                                        let newFormula = sprintf "=IFERROR(%s,%s)" arg1 (formatLiteral v)
                                        if unbox cell.HasArray then cell.FormulaArray <- newFormula
                                        else cell.Formula <- newFormula
                                        // 以绿色标记已更新的目标单元格
                                        cell.Interior.Color <- float UpdateDefaultsColor
                                        processed <- processed + 1
                                    | _ ->
                                        unconform.Add(address cell, formula)
                                else
                                    // 固化当前计算结果，清除公式
                                    cell.Value2 <- v
                                    processed <- processed + 1
                if isProtected then
                    // 重新保护副本，与源工作表保持一致
                    try ws.Protect()
                    with _ -> ()
        { ProcessedCount = processed
          FunctionMissing = false
          ErrorCells = errors.ToArray()
          UnconformCells = unconform.ToArray()
          ArrayFormulaCells = arrays.ToArray()
          ProtectedSheets = protectedSheets.ToArray() }

    /// 更新默认值：对当前工作簿在位修改，处理所有 =IFERROR(函数(...), 常量) 结构的单元格，
    /// 将兜底常量替换为最新真值并以绿色标记。不限函数名（内置、XLL、自定义函数均适用）。
    let updateDefaults (wb: Workbook) : RunResult =
        run (true, wb, "")

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
            sb.AppendLine(sprintf "跳过数组公式单元格 %d 个：" result.ArrayFormulaCells.Length) |> ignore
            for (addr, formula) in result.ArrayFormulaCells do
                sb.AppendLine(sprintf "  %s：%s" addr formula) |> ignore
        if result.ProtectedSheets.Length > 0 then
            sb.AppendLine(sprintf "受保护且无法解除保护而跳过的工作表 %d 个：%s" result.ProtectedSheets.Length (String.concat "、" result.ProtectedSheets)) |> ignore
        sb.ToString().TrimEnd()
