///「删除中转单元格」：把当前单元格（中转格）的公式内联回所有依赖它的公式，然后清空该单元格。
///
/// 受不受理的判据是「一元表达式」：本单元格公式整体必须是一个原子——字面量、单元格引用、
/// 定义名称、函数调用、整体括起来的一组表达式、数组常量。这类文本内联到任何位置都不必加括号，
/// 也不会因运算符优先级改变语义；反之 A1+B1、-A1、A1% 这类带顶层运算符的公式内联到 =B1*2
/// 会变成 =A1+B1*2，语义改变，一律不受理。
///
/// 处理范围：当前工作簿全部工作表里引用本单元格的公式，把「本单元格地址」这段文本整体替换为
/// 本单元格的公式文本（去掉前导 =），随后清空本单元格。引用文本原样搬运、不做行列偏移，
/// 内联后的表达式仍指向原来那些单元格（公式文本原地解析，相对/绝对写法因此不影响结果）；
/// 只有搬到其他工作表时，本表局部引用会被解释成那张表的单元格，故逐个补上本表前缀。
///
/// 不受理时不做任何改动（避免改了一半留下坏引用）：
/// - 本单元格不是公式、属于多格数组公式、是合并单元格；
/// - 本单元格公式不是一元表达式，或公式里引用了本单元格自身；
/// - 有公式以区域形式引用包含本单元格的范围（如 A40 被 SUM(A40:A42) 引用），无法逐格替换；
/// - 内联后某个公式超过 8192 字符。
/// 名称定义、条件格式、数据验证、图表系列、其他工作簿的链接不在扫描范围内。
module ExcelNumericalMethods.RelayDelete

open System
open System.Collections.Generic
open System.Text.RegularExpressions
open Microsoft.Office.Interop.Excel

/// 工作表的最大行号 / 列号，用于把整行整列引用展开成矩形
let [<Literal>] private MaxSheetRow = 1048576

/// 工作表的最大列号
let [<Literal>] private MaxSheetCol = 16384

// ---------------------------------------------------------------------------
// 地址正则
// ---------------------------------------------------------------------------

/// 单格或矩形区域：A1、$A$1、A1:B2
let private areaRgx =
    Regex(@"^(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)", RegexOptions.CultureInvariant)

/// 整列引用：A:A、$A:$B
let private colsRgx =
    Regex(@"^(\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})", RegexOptions.CultureInvariant)

/// 整行引用：1:1、$1:$3
let private rowsRgx =
    Regex(@"^(\$?\d+:\$?\d+)", RegexOptions.CultureInvariant)

/// 单格：A1、$A$1
let private cellPartRgx =
    Regex(@"^\$?([A-Za-z]{1,3})\$?(\d+)$", RegexOptions.CultureInvariant)

/// 单列：A、$A
let private colPartRgx =
    Regex(@"^\$?([A-Za-z]{1,3})$", RegexOptions.CultureInvariant)

/// 单行：1、$1
let private rowPartRgx =
    Regex(@"^\$?(\d+)$", RegexOptions.CultureInvariant)

/// 列字母转列号：A -> 1，AA -> 27
let private colIndexOf (letters: string) : int =
    letters.ToUpperInvariant().ToCharArray()
    |> Array.fold (fun acc ch -> acc * 26 + (int ch - int 'A' + 1)) 0

/// 未加引号的名字允许的字符：字母、数字、下划线、点
let private isNameChar (c: char) =
    Char.IsLetterOrDigit c || c = '_' || c = '.'

/// 从 k 处前向读一个名字 token，返回（文本，结束下标）。支持 '引号名'、[工作簿]名、裸名。
let private readNameFwd (s: string) (k: int) : (string * int) option =
    let n = s.Length
    if k < 0 || k >= n then None
    elif s.[k] = '\'' then
        // 引号名：内部 '' 表示一个单引号
        let sb = Text.StringBuilder()
        let mutable j = k + 1
        let mutable closed = false
        while not closed && j < n do
            if s.[j] = '\'' then
                if j + 1 < n && s.[j + 1] = '\'' then
                    sb.Append('\'') |> ignore
                    j <- j + 2
                else
                    closed <- true
                    j <- j + 1
            else
                sb.Append(s.[j]) |> ignore
                j <- j + 1
        if closed then Some(sb.ToString(), j) else None
    elif s.[k] = '[' then
        // [Book1.xlsx]Sheet1（引号外的工作簿前缀）
        let e = s.IndexOf(']', k)
        if e < 0 then None
        else
            let mutable j = e + 1
            while j < n && isNameChar s.[j] do
                j <- j + 1
            Some(s.Substring(k, j - k), j)
    elif isNameChar s.[k] then
        let mutable j = k
        while j < n && isNameChar s.[j] do
            j <- j + 1
        Some(s.Substring(k, j - k), j)
    else None

/// 从 k 处读一个地址（单格 / 矩形 / 整列 / 整行），返回（地址文本，结束下标）。
/// 后面紧跟字母数字、$、#、括号的字面不算地址（避开函数名 LOG10( 之类）。
let private readAddress (s: string) (k: int) : (string * int) option =
    let n = s.Length
    if k < 0 || k >= n then None
    else
        let rest = s.Substring k
        let m =
            let a = areaRgx.Match rest
            if a.Success then a
            else
                let b = colsRgx.Match rest
                if b.Success then b
                else rowsRgx.Match rest
        if not m.Success then None
        else
            let v = m.Groups.[1].Value
            let after = k + v.Length
            let boundaryOk =
                if after >= n then true
                else
                    let c = s.[after]
                    not (Char.IsLetterOrDigit c)
                    && c <> '_'
                    && c <> '.'
                    && c <> '$'
                    && c <> '#'
                    && c <> '('
            if boundaryOk then Some(v, after) else None

/// 名字 token 里的工作簿部分：`[Book1.xlsx]Sheet1` -> Some "Book1.xlsx"
let private bookNameOf (name: string) : string option =
    let lb = name.IndexOf '['
    let rb = name.IndexOf ']'
    if lb >= 0 && rb > lb then Some(name.Substring(lb + 1, rb - lb - 1)) else None

/// 名字 token 里的工作表名列表：去掉工作簿部分，按冒号拆 3D 引用
let private sheetNamesOf (name: string) : string list =
    match bookNameOf name with
    | Some _ ->
        let rb = name.IndexOf ']'
        name.Substring(rb + 1).Split(':')
        |> Array.map (fun x -> x.Trim())
        |> Array.filter (fun x -> x.Length > 0)
        |> List.ofArray
    | None ->
        name.Split(':')
        |> Array.map (fun x -> x.Trim())
        |> Array.filter (fun x -> x.Length > 0)
        |> List.ofArray

/// 公式里一处单元格引用
type CellRef =
    { /// 引用原文在公式中的起始下标
      Start: int
      /// 引用原文长度
      Length: int
      /// 引用原文，如 A1、$A$1、Sheet2!B2、'My Sheet'!$B$2、[Book1.xlsx]Sheet1!A1
      Text: string
      /// 引用中写出的工作表名（未写前缀为空；3D 引用为起止两端）
      Sheets: string list
      /// 是否 3D 引用（Sheet1:Sheet3!A1）
      Is3D: bool
      /// 外部工作簿名（未写工作簿前缀为 None）
      BookName: string option
      /// 地址部分，如 A1、$A$1、A1:B2、A:A、1:1
      Address: string }

/// 解析公式中的全部单元格引用（本表引用也包含），跳过字符串字面量。
let referencesOf (formula: string) : CellRef list =
    let s = if isNull formula then "" else formula
    let n = s.Length
    let result = ResizeArray<CellRef>()
    let mutable i = 0
    let mutable inString = false

    let add (startIdx: int) (afterIdx: int) (names: string list) (is3d: bool) (book: string option) (addr: string) =
        result.Add
            { Start = startIdx
              Length = afterIdx - startIdx
              Text = s.Substring(startIdx, afterIdx - startIdx)
              Sheets = names
              Is3D = is3d
              BookName = book
              Address = addr }

    while i < n do
        let c = s.[i]
        if inString then
            if c = '"' then
                if i + 1 < n && s.[i + 1] = '"' then i <- i + 1
                else inString <- false
            i <- i + 1
        elif c = '"' then
            inString <- true
            i <- i + 1
        elif c = '\'' || c = '[' || Char.IsLetter c || c = '_' then
            match readNameFwd s i with
            | None -> i <- i + 1
            | Some (name1, e1) ->
                if e1 < n && s.[e1] = '!' then
                    match readAddress s (e1 + 1) with
                    | Some (addr, after) ->
                        add i after (sheetNamesOf name1) false (bookNameOf name1) addr
                        i <- after
                    | None -> i <- e1 + 1
                elif e1 < n && s.[e1] = ':' then
                    // 3D 引用 Sheet1:Sheet3!A1，或本表区域 A1:B2
                    match readNameFwd s (e1 + 1) with
                    | Some (name2, e2) when e2 < n && s.[e2] = '!' ->
                        match readAddress s (e2 + 1) with
                        | Some (addr, after) ->
                            let book =
                                match bookNameOf name1 with
                                | Some _ as b -> b
                                | None -> bookNameOf name2
                            add i after (sheetNamesOf name1 @ sheetNamesOf name2) true book addr
                            i <- after
                        | None -> i <- e2 + 1
                    | _ ->
                        match readAddress s i with
                        | Some (addr, after) ->
                            add i after [] false None addr
                            i <- after
                        | None -> i <- max (i + 1) e1
                else
                    match readAddress s i with
                    | Some (addr, after) ->
                        add i after [] false None addr
                        i <- after
                    | None -> i <- max (i + 1) e1
        else
            match readAddress s i with
            | Some (addr, after) ->
                add i after [] false None addr
                i <- after
            | None -> i <- i + 1
    List.ofSeq result

/// 地址转矩形（起始行, 起始列, 结束行, 结束列），已归一化。
/// 整列按 1..MaxSheetRow 展开，整行按 1..MaxSheetCol 展开，因此「整行/整列」不会等于单格。
let rectOf (address: string) : (int * int * int * int) option =
    let part (t: string) : (int * int * int * int) option =
        let p = t.Replace("$", "").Trim()
        let mc = cellPartRgx.Match p
        if mc.Success then
            let c = colIndexOf mc.Groups.[1].Value
            let r = Int32.Parse mc.Groups.[2].Value
            Some(r, c, r, c)
        else
            let mcol = colPartRgx.Match p
            if mcol.Success then
                let c = colIndexOf mcol.Groups.[1].Value
                Some(1, c, MaxSheetRow, c)
            else
                let mrow = rowPartRgx.Match p
                if mrow.Success then
                    let r = Int32.Parse mrow.Groups.[1].Value
                    Some(r, 1, r, MaxSheetCol)
                else None
    match address.Split(':') with
    | [| one |] -> part one
    | [| a; b |] ->
        match part a, part b with
        | Some (r1, c1, _, _), Some (_, _, r2, c2) ->
            Some(min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    | _ -> None

/// 一处引用与目标单元格的关系
type RefRelation =
    /// 恰好引用目标单元格（可替换）
    | IsTarget
    /// 区域引用包含目标单元格（无法逐格替换）
    | Encloses
    /// 与目标单元格无关
    | Unrelated

/// 判断引用是否指向目标单元格、或以区域形式包含它。
/// selfBook 为当前工作簿名，用于把「写全工作簿名的自引用」也认作本工作簿。
let relationTo (selfBook: string) (targetSheet: string) (targetRow: int) (targetCol: int) (referrerSheet: string) (r: CellRef) : RefRelation =
    let same (a: string) (b: string) = String.Equals(a, b, StringComparison.OrdinalIgnoreCase)
    let sheets = if r.Sheets.IsEmpty then [ referrerSheet ] else r.Sheets
    let inThisBook =
        match r.BookName with
        | Some bk -> same bk selfBook
        | None -> true
    if not inThisBook then Unrelated
    elif r.Is3D then
        if sheets |> List.exists (fun nm -> same nm targetSheet) then Encloses else Unrelated
    elif not (sheets |> List.forall (fun nm -> same nm targetSheet)) then Unrelated
    else
        match rectOf r.Address with
        | None -> Unrelated
        | Some (r1, c1, r2, c2) ->
            if r1 = targetRow && c1 = targetCol && r2 = targetRow && c2 = targetCol then IsTarget
            elif targetRow >= r1 && targetRow <= r2 && targetCol >= c1 && targetCol <= c2 then Encloses
            else Unrelated

// ---------------------------------------------------------------------------
// 一元表达式判定
// ---------------------------------------------------------------------------

/// 顶层运算符字符（空格是交集运算符，单独判）
let private opChars = "+-*/^&=<>%,;"

/// 字符串字面量："abc"、""""（内部双引号成对转义）
let private stringLiteralRgx = Regex("^\"([^\"]|\"\")*\"$", RegexOptions.CultureInvariant)

/// 数字字面量：123、1.5、.5、1.5E-3（不接受正负号，负号是顶层一元运算）
let private numberLiteralRgx = Regex(@"^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$", RegexOptions.CultureInvariant)

/// 布尔字面量
let private boolLiteralRgx = Regex("^(TRUE|FALSE)$", RegexOptions.IgnoreCase)

/// s 是否整体被一对 openCh..closeCh 括住（字符串与 '引号名' 内的括号不计数）
let private balancedWhole (openCh: char) (closeCh: char) (s: string) : bool =
    if s.Length < 2 || s.[0] <> openCh then false
    else
        let n = s.Length
        let mutable i = 0
        let mutable depth = 0
        let mutable inStr = false
        let mutable inName = false
        let mutable ok = true
        while ok && i < n do
            let c = s.[i]
            if inStr then
                if c = '"' then
                    if i + 1 < n && s.[i + 1] = '"' then i <- i + 1
                    else inStr <- false
            elif inName then
                if c = '\'' then
                    if i + 1 < n && s.[i + 1] = '\'' then i <- i + 1
                    else inName <- false
            else
                match c with
                | '"' -> inStr <- true
                | '\'' -> inName <- true
                | _ when c = openCh -> depth <- depth + 1
                | _ when c = closeCh ->
                    depth <- depth - 1
                    if depth = 0 && i <> n - 1 then ok <- false
                | _ -> ()
            i <- i + 1
        ok && depth = 0 && not inStr && not inName

/// 名字/引用文本里没有顶层运算符（引号内不算），也没有括号
let private noTopLevelOperator (s: string) : bool =
    let n = s.Length
    let mutable i = 0
    let mutable inStr = false
    let mutable inName = false
    let mutable ok = true
    while ok && i < n do
        let c = s.[i]
        if inStr then
            if c = '"' then
                if i + 1 < n && s.[i + 1] = '"' then i <- i + 1
                else inStr <- false
        elif inName then
            if c = '\'' then
                if i + 1 < n && s.[i + 1] = '\'' then i <- i + 1
                else inName <- false
        else
            if c = '"' then inStr <- true
            elif c = '\'' then inName <- true
            elif c = '(' || c = ')' then ok <- false
            elif Char.IsWhiteSpace c then ok <- false
            elif opChars.IndexOf c >= 0 then ok <- false
        i <- i + 1
    ok && not inStr && not inName

/// 公式主体（不含前导 =）是否一元表达式：
/// 内联到其他公式的任何位置都不必加括号，也不会因运算符优先级改变语义。
/// 受理：字面量、单元格引用（含整行整列与跨表前缀）、定义名称、函数调用、
/// 整体括起来的表达式、数组常量。
let isUnitExpression (body: string) : bool =
    let b = (if isNull body then "" else body).Trim()
    let b = if b.StartsWith("@", StringComparison.Ordinal) then b.Substring(1).Trim() else b
    if b.Length = 0 then false
    elif b.StartsWith("\"", StringComparison.Ordinal) then stringLiteralRgx.IsMatch b
    elif b.StartsWith("(", StringComparison.Ordinal) then balancedWhole '(' ')' b
    elif b.StartsWith("{", StringComparison.Ordinal) then balancedWhole '{' '}' b
    elif numberLiteralRgx.IsMatch b then true
    elif boolLiteralRgx.IsMatch b then true
    else
        match readNameFwd b 0 with
        | Some (_, e) when e < b.Length && b.[e] = '(' -> balancedWhole '(' ')' (b.Substring e)
        | _ -> noTopLevelOperator b

/// 工作表名按公式引用规则加引号：Sheet1、Sheet_1 不加；含空格、中划线、以数字开头等加 '...'（内部 ' 写成 ''）
let quoteSheetName (name: string) : string =
    let needsQuote =
        name.Length = 0
        || Char.IsDigit name.[0]
        || name |> Seq.exists (fun c -> not (Char.IsLetterOrDigit c) && c <> '_' && c <> '.')
    if needsQuote then "'" + name.Replace("'", "''") + "'" else name

/// 内联文本：本单元格公式搬到 destSheet 时用的文本。
/// 引用文本原样搬到同表即可；搬到别的表时，本表局部引用（不带表前缀的 A1、A1:B2、A:A、1:1）
/// 会被解释成那张表的单元格，故逐个补上本表前缀，其余（已写表名、外部工作簿、名称、字面量）不动。
let inlineText (sourceSheet: string) (destSheet: string) (body: string) : string =
    if String.Equals(destSheet, sourceSheet, StringComparison.OrdinalIgnoreCase) then body
    else
        let reps =
            referencesOf body
            |> List.filter (fun r -> r.Sheets.IsEmpty)
            |> List.map (fun r -> r.Start, r.Length, quoteSheetName sourceSheet + "!" + r.Text)
        if reps.IsEmpty then body
        else WorksheetRefs.rewriteFormula body reps

// ---------------------------------------------------------------------------
// 拆除
// ---------------------------------------------------------------------------

/// 拆除结果
type DeleteResult =
    { /// 是否已删除（依赖公式已改写且本单元格已清空）
      Deleted: bool
      /// 本单元格地址（工作表名!列字母行号）
      Address: string
      /// 用于内联的公式文本（不含前导 =）
      Body: string
      /// 已改写的依赖单元格地址
      Rewritten: string array
      /// 阻止删除的原因；有值时未做任何改动
      Blockers: string array
      /// 状态栏提示文本
      Message: string }

/// 待改写的依赖公式
type private Job =
    { Target: Range
      /// 目标是否多格 CSE 数组（需整块写入）
      IsArrayRange: bool
      Address: string
      NewFormula: string }

/// 删除中转单元格：把本单元格的公式内联回全部依赖公式，再清空本单元格。
/// 本单元格不是公式、属于多格数组、合并单元格、公式不是一元表达式、
/// 有区域引用包含本单元格、内联后公式超长时，返回 Deleted=false 与原因，且不做任何改动。
let delete (cell: Range) : DeleteResult =
    let ws = cell.Worksheet
    let wb =
        match box ws.Parent with
        | :? Workbook as w -> w
        | _ -> Unchecked.defaultof<Workbook>
    let addr = Publishing.cellAddress ws cell
    let fail (reason: string) =
        { Deleted = false
          Address = addr
          Body = ""
          Rewritten = [||]
          Blockers = [| reason |]
          Message = sprintf "%s：%s" addr reason }

    if isNull (box wb) then
        fail "当前工作表不属于任何工作簿。"
    elif not (unbox cell.HasFormula) then
        fail "当前单元格不是公式，不是中转单元格。"
    elif unbox cell.MergeCells then
        fail "当前单元格是合并单元格，不能作为中转单元格删除。"
    elif unbox cell.HasArray && not (Publishing.isSingleCellArray cell) then
        fail "当前单元格属于多格数组公式，不能单独删除。"
    else
        let formula =
            try string (Formula2.getRaw cell)
            with _ -> string cell.Formula
        let body =
            let t = formula.Trim()
            if t.StartsWith("=", StringComparison.Ordinal) then t.Substring(1).Trim() else t

        if not (isUnitExpression body) then
            fail (sprintf "公式 =%s 不是一元表达式：内联到其他公式需加括号或会改变语义。" body)
        elif referencesOf body |> List.exists (fun r -> relationTo wb.Name ws.Name cell.Row cell.Column ws.Name r = IsTarget) then
            fail (sprintf "公式 =%s 引用了本单元格自身。" body)
        else
            let jobs = ResizeArray<Job>()
            let blockers = ResizeArray<string>()

            for sheet in Traversal.getWorksheets wb do
                for referrer in Publishing.getFormulaCells sheet do
                    let isSelf =
                        String.Equals(sheet.Name, ws.Name, StringComparison.OrdinalIgnoreCase)
                        && referrer.Row = cell.Row
                        && referrer.Column = cell.Column
                    if not isSelf then
                        let f =
                            try string (Formula2.getRaw referrer)
                            with _ -> ""
                        if f.Length > 0 then
                            let referrerAddr = Publishing.cellAddress sheet referrer
                            let reps = ResizeArray<int * int * string>()
                            let inlined = inlineText ws.Name sheet.Name body
                            for r in referencesOf f do
                                match relationTo wb.Name ws.Name cell.Row cell.Column sheet.Name r with
                                | IsTarget -> reps.Add(r.Start, r.Length, inlined)
                                | Encloses ->
                                    blockers.Add(
                                        sprintf "%s 的区域引用 %s 包含本单元格，无法逐格替换。" referrerAddr r.Text)
                                | Unrelated -> ()
                            if reps.Count > 0 then
                                let target, isArrayRange, isAnchor = WorksheetRefs.arrayUnitOf referrer
                                if isAnchor then
                                    let newFormula = WorksheetRefs.rewriteFormula f (List.ofSeq reps)
                                    if newFormula.Length > 8192 then
                                        blockers.Add(sprintf "%s 内联后公式超过 8192 字符。" referrerAddr)
                                    elif newFormula <> f then
                                        jobs.Add
                                            { Target = target
                                              IsArrayRange = isArrayRange
                                              Address = referrerAddr
                                              NewFormula = newFormula }

            if blockers.Count > 0 then
                { Deleted = false
                  Address = addr
                  Body = body
                  Rewritten = [||]
                  Blockers = blockers.ToArray()
                  Message = sprintf "%s 未删除：%d 处依赖无法内联。" addr blockers.Count }
            else
                let app = wb.Application
                let oldCalc = app.Calculation
                let failed = ResizeArray<string>()
                let rewritten = ResizeArray<string>()
                try
                    app.Calculation <- XlCalculation.xlCalculationManual
                    for j in jobs do
                        let ok =
                            try
                                if j.IsArrayRange then
                                    (try
                                        j.Target.FormulaArray <- j.NewFormula
                                        true
                                     with _ -> false)
                                    || (try
                                            Formula2.set j.Target j.NewFormula
                                            true
                                        with _ -> false)
                                else
                                    Formula2.set j.Target j.NewFormula
                                    true
                            with _ -> false
                        if ok then rewritten.Add j.Address else failed.Add j.Address

                    // 依赖公式全部写好后才清空本单元格，避免中途失败留下断链
                    if failed.Count = 0 then
                        cell.ClearContents() |> ignore
                finally
                    app.Calculation <- oldCalc
                    app.Calculate()

                if failed.Count = 0 then
                    { Deleted = true
                      Address = addr
                      Body = body
                      Rewritten = rewritten.ToArray()
                      Blockers = [||]
                      Message =
                        sprintf "已删除中转单元格 %s：内联 =%s，改写 %d 个依赖公式。"
                            addr body rewritten.Count }
                else
                    { Deleted = false
                      Address = addr
                      Body = body
                      Rewritten = rewritten.ToArray()
                      Blockers =
                        failed |> Seq.map (fun a -> sprintf "%s 写入失败。" a) |> Seq.toArray
                      Message =
                        sprintf "写入失败 %d 处，%s 未清空（已改写的公式需人工检查）。" failed.Count addr }
