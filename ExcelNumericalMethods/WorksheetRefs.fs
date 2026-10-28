namespace ExcelNumericalMethods

open System
open System.Collections.Generic
open System.Text
open System.Text.RegularExpressions
open Microsoft.Office.Interop.Excel

///「工作表引用」：把本工作表公式对外部地址的引用集中到工作表底部的中转区，
/// 原公式改为引用中转区，使复杂公式内部不再嵌套外部地址。
///
/// 外部地址指本工作表以外的地址：其他工作表（Sheet2!A1）、其他工作簿（[Book1.xlsx]Sheet1!A1）。
/// 可中转的形状：任意矩形区域（单格 Sheet2!C5、竖向 C7:C9、横向 A1:C1、二维 A1:B2）。
/// 中转块以区域的左上单元格为锚点，在 A 列写**一条**整块区域公式（如 =Sheet2!$C$7:$C$9），
/// 由动态数组向右下溢出成与原区域相同的形状；原公式中该引用替换为同形状的绝对区域
/// （$A$40:$A$42），语义与形状都不变。
/// 由于靠溢出铺开，目标矩形必须为空：中转区写在 UsedRange 之下，其右下角天然空闲，
/// 块按行带依次向下堆叠即可；写入前再对矩形 ClearContents 一次，兜住旧提示行等残留。
/// 不中转的情形：整行整列（A:A、1:1，范围不确定）、3D 引用（Sheet1:Sheet3!A1）、
/// 超大区域——这些在块尾部逐行写「【未中转】…」提示，原公式保持不动。
///
/// 多格数组公式（Ctrl+Shift+Enter 数组，或动态数组溢出成多格）按**数组单元**处理：
/// 整个数组只有左上锚点持有真公式，其余格只是结果占位格；收集与改写都以锚点为单位做一次，
/// 写入落在整个数组区域上，避免「不能更改数组的一部分」而漏掉其中的外部引用。
/// 本表完全没有可中转的外部地址时，也在表尾写一行「【提示】没有可中转的外部地址」。
///
/// 布局：中转区写在 UsedRange 末行下方空一行处，块与块按行带依次向下堆叠，
/// 提示行写在本次块之后。
/// 每次运行都**从 UsedRange 下方重新追加**，不识别也不复用上一次写的中转区；
/// 重复中转（旧中转格又被当作外部引用转一次）由其使用者自行处理。
module WorksheetRefs =

    /// 公式中一处外部地址引用（带位置，用于文本替换）
    type RefSpan =
        { /// 引用原文在公式字符串中的起始下标
          Start: int
          /// 引用原文长度
          Length: int
          /// 引用原文，如 Sheet2!A1、'My Sheet'!$B$2、[Book1.xlsx]Sheet1!A1
          Text: string
          /// 引用中写出的工作表名（已去引号；3D 引用为起止两端；外部工作簿名不计入）
          Sheets: string list
          /// 是否指向其他工作簿
          IsExternalBook: bool
          /// 地址部分，如 A1、$A$1、A1:B2、A:A、1:1
          Address: string }

    /// 中转块：以外部区域左上格为锚点铺开的矩形
    type RelayBlock =
        { /// 块高（行数）
          Rows: int
          /// 块宽（列数）
          Cols: int
          /// 锚点格公式：整块一条区域引用（如 =大气与烟气参数!$C$7:$C$9），
          /// 写入后由动态数组向右下溢出占满 Rows × Cols
          Formula: string }

    /// 中转结果
    type RelayResult =
        { /// 本次写入的中转单元格数（各块 Rows × Cols 之和）
          NewRelayCount: int
          /// 中转区单元格总数（不做复用，与 NewRelayCount 相同）
          RelayTotal: int
          /// 中转区起始行（无中转格时为 0）
          RelayStartRow: int
          /// 中转区结束行（无中转格时为 0）
          RelayEndRow: int
          /// 中转区占用行数
          RelayRows: int
          /// 中转区最大宽度（列数）
          RelayCols: int
          /// 改写的公式单元格数
          RewrittenCells: int
          /// 未中转的引用数（整行整列、3D 引用、超大区域）
          SkippedRefs: int
          /// 未中转的引用原文（去重，便于提示）
          SkippedTexts: string array
          /// 写在表尾的提示单元格数
          HintCount: int
          /// 提示单元格起始行（无提示时为 0）
          HintStartRow: int
          /// 本表没有可中转的外部地址时，写在表尾的说明所在行（未写则为 0）
          NoRefNoticeRow: int
          /// 按多格数组单元整体改写的数组锚点地址（A19 这类；空数组表示本表没有多格数组）
          ArrayUnits: string array
          /// 写入或改写失败的单元格地址
          FailedCells: string array }

    /// 未中转提示的文字前缀，同时用于识别表尾旧提示格
    let HintMark = "【未中转】"

    /// 表尾其他说明的文字前缀（同样用于识别旧提示格，便于重建与清理）
    let NoteMark = "【提示】"

    /// 本表没有可中转的外部地址时写在表尾的说明
    let NoRefText = NoteMark + "没有可中转的外部地址"

    // ---------------------------------------------------------------------
    // 地址正则
    // ---------------------------------------------------------------------

    /// 单格或矩形区域：A1、$A$1、A1:B2
    let private areaRgx =
        Regex(@"^(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)", RegexOptions.CultureInvariant)

    /// 整列引用：A:A、$A:$B
    let private colsRgx =
        Regex(@"^(\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})", RegexOptions.CultureInvariant)

    /// 整行引用：1:1、$1:$3
    let private rowsRgx =
        Regex(@"^(\$?\d+:\$?\d+)", RegexOptions.CultureInvariant)

    /// 单格地址（非区域、非整行整列）
    let private singleCellRgx =
        Regex(@"^\$?[A-Za-z]{1,3}\$?\d+$", RegexOptions.CultureInvariant)

    let private colRowRgx =
        Regex(@"^([A-Za-z]{1,3})(\d+)$", RegexOptions.CultureInvariant)

    let private colsOnlyRgx =
        Regex(@"^([A-Za-z]{1,3})$", RegexOptions.CultureInvariant)

    let private rowsOnlyRgx =
        Regex(@"^(\d+)$", RegexOptions.CultureInvariant)

    // ---------------------------------------------------------------------
    // 行列换算
    // ---------------------------------------------------------------------

    let private colIndexOf (letters: string) : int =
        letters.ToUpperInvariant().ToCharArray()
        |> Array.fold (fun acc ch -> acc * 26 + (int ch - int 'A' + 1)) 0

    /// 列号转列名：1 -> A，27 -> AA
    let colNameOf (index: int) : string =
        let rec loop n acc =
            if n <= 0 then acc
            else loop ((n - 1) / 26) (string (char (int 'A' + (n - 1) % 26)) :: acc)
        String.Concat(loop index [])

    // ---------------------------------------------------------------------
    // 公式解析：只做「表前缀!地址」的文本识别，不建立完整表达式树
    // ---------------------------------------------------------------------

    /// 工作表名（不加引号时）允许的字符：字母、数字、下划线、点
    let private isNameChar (c: char) =
        Char.IsLetterOrDigit c || c = '_' || c = '.'

    /// 从 endIdx 位置向前读出一个 token（名字 / '引号名' / [工作簿]），返回 token 起点。
    let private backToken (s: string) (endIdx: int) : int option =
        if endIdx < 0 || endIdx >= s.Length then None
        else
            let c = s.[endIdx]
            if c = '\'' then
                // 引号名：向前找配对的起始引号，跳过成对的 '' 转义
                let mutable j = endIdx - 1
                let mutable start = -1
                while j >= 0 && start < 0 do
                    if s.[j] = '\'' then
                        if j - 1 >= 0 && s.[j - 1] = '\'' then j <- j - 2
                        else start <- j
                    else
                        j <- j - 1
                if start < 0 then None else Some start
            elif c = ']' then
                // [Book1.xlsx]
                let mutable j = endIdx - 1
                while j >= 0 && s.[j] <> '[' do
                    j <- j - 1
                if j < 0 then None else Some j
            elif isNameChar c then
                let mutable j = endIdx
                while j >= 0 && isNameChar s.[j] do
                    j <- j - 1
                Some(j + 1)
            else None

    /// 从 `!` 前一个字符向前读出完整的表前缀（可含 [工作簿] 与 3D 冒号），返回前缀起点。
    let private prefixStart (s: string) (endIdx: int) : int option =
        let rec loop (idx: int) (acc: int option) =
            match backToken s idx with
            | None -> acc
            | Some st ->
                let acc = Some st
                let k = st - 1
                if k >= 0 && s.[k] = ':' then loop (k - 1) acc
                elif k >= 0 && s.[k] = ']' then loop k acc
                else acc
        loop endIdx None

    /// 去掉整体包裹的单引号，'' 还原为一个 '
    let private unquote (t: string) =
        if t.Length >= 2 && t.[0] = '\'' && t.[t.Length - 1] = '\'' then
            t.Substring(1, t.Length - 2).Replace("''", "'")
        else t

    /// 解析表前缀：返回（引用的工作表名列表，是否外部工作簿）
    let private parsePrefix (prefixText: string) : string list * bool =
        let t = unquote (prefixText.Trim())
        let lb = t.LastIndexOf '['
        let rb = t.LastIndexOf ']'
        let splitNames (text: string) =
            text.Split(':')
            |> Array.map (fun x -> x.Trim().Trim('\''))
            |> Array.filter (fun x -> x.Length > 0)
            |> List.ofArray
        if lb >= 0 && rb > lb then
            // 外部工作簿：[Book1.xlsx]Sheet1；引号内还可能带路径 C:\dir\[Book1.xlsx]Sheet1
            splitNames (t.Substring(rb + 1)), true
        else
            splitNames t, false

    /// 地址起始位置处的地址文本（带边界检查；跳过 #REF!、Sheet1!A1X 之类）
    let private addressAt (s: string) (startIdx: int) : string option =
        if startIdx >= s.Length then None
        else
            let rest = s.Substring startIdx
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
                let after = startIdx + v.Length
                let ok =
                    if after >= s.Length then true
                    else
                        let c = s.[after]
                        not (Char.IsLetterOrDigit c)
                        && c <> '_'
                        && c <> '.'
                        && c <> '('
                        && c <> '$'
                        && c <> '#'
                if ok then Some v else None

    /// 解析公式中全部「外部地址」引用（带位置）。跳过字符串字面量（含 "" 转义）与非法地址。
    /// 是否属于本表由调用方判断。
    let referencesOf (formula: string) : RefSpan list =
        let result = ResizeArray<RefSpan>()
        let s = if isNull formula then "" else formula
        let n = s.Length
        let mutable i = 0
        let mutable inString = false
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
            elif c = '!' then
                match prefixStart s (i - 1) with
                | None -> i <- i + 1
                | Some start ->
                    match addressAt s (i + 1) with
                    | None -> i <- i + 1
                    | Some addr ->
                        let length = i + 1 + addr.Length - start
                        let text = s.Substring(start, length)
                        let prefixText = s.Substring(start, i - start)
                        let names, isBook = parsePrefix prefixText
                        result.Add
                            { Start = start
                              Length = length
                              Text = text
                              Sheets = names
                              IsExternalBook = isBook
                              Address = addr }
                        i <- i + 1 + addr.Length
            else
                i <- i + 1
        List.ofSeq result

    /// 把地址里的行列都加上 $：A1 -> $A$1，A1:B2 -> $A$1:$B$2，A:A -> $A:$A
    let absolutizeAddress (address: string) : string =
        let one (part: string) =
            let p = part.Replace("$", "")
            let m = colRowRgx.Match p
            if m.Success then
                sprintf "$%s$%s" (m.Groups.[1].Value.ToUpperInvariant()) m.Groups.[2].Value
            else
                let mc = colsOnlyRgx.Match p
                if mc.Success then "$" + mc.Groups.[1].Value.ToUpperInvariant()
                else
                    let mr = rowsOnlyRgx.Match p
                    if mr.Success then "$" + mr.Groups.[1].Value
                    else part
        address.Split(':') |> Array.map one |> String.concat ":"

    /// 按替换表改写公式：(起始下标, 原长度, 新文本)
    let rewriteFormula (formula: string) (replacements: (int * int * string) list) : string =
        let sb = StringBuilder()
        let mutable pos = 0
        for (start, length, text) in replacements |> List.sortBy (fun (st, _, _) -> st) do
            if start >= pos && start + length <= formula.Length then
                sb.Append(formula.Substring(pos, start - pos)).Append(text) |> ignore
                pos <- start + length
        sb.Append(formula.Substring pos) |> ignore
        sb.ToString()

    /// 引用是否指向本工作表（显式写自身表名也算本表）
    let isSelfRef (sheetName: string) (r: RefSpan) : bool =
        not r.IsExternalBook
        && not r.Sheets.IsEmpty
        && r.Sheets |> List.forall (fun nm -> String.Equals(nm, sheetName, StringComparison.OrdinalIgnoreCase))

    let private isSingleCell (address: string) = singleCellRgx.IsMatch address

    /// 解析矩形地址：返回（起始行, 起始列, 结束行, 结束列），已按左上到右下归一化。
    /// 单格、A1:B2、$A$1:$B$2 均可；整行整列（A:A、1:1）与其他写法返回 None。
    let rectOf (address: string) : (int * int * int * int) option =
        let parsePart (t: string) =
            let m = colRowRgx.Match(t.Replace("$", ""))
            if m.Success then Some(colIndexOf m.Groups.[1].Value, Int32.Parse m.Groups.[2].Value)
            else None
        match address.Split(':') with
        | [| one |] -> parsePart one |> Option.map (fun (c, r) -> (r, c, r, c))
        | [| a; b |] ->
            match parsePart a, parsePart b with
            | Some(c1, r1), Some(c2, r2) -> Some(min r1 r2, min c1 c2, max r1 r2, max c1 c2)
            | _ -> None
        | _ -> None

    /// 中转块上限：超过则视为过大区域，不中转
    let [<Literal>] MaxBlockRows = 1000

    /// 中转块列数上限
    let [<Literal>] MaxBlockCols = 100

    /// 中转块单元格总数上限
    let [<Literal>] MaxBlockCells = 10000

    /// 块是否在允许范围内
    let blockFits (b: RelayBlock) : bool =
        b.Rows <= MaxBlockRows
        && b.Cols <= MaxBlockCols
        && b.Rows * b.Cols <= MaxBlockCells

    /// 一处外部地址对应的中转块：以左上单元格为锚点，整块写一条区域公式，
    /// 由动态数组向右下溢出成原形状。
    /// 单格给 1×1；C7:C9 给 3×1；A1:C1 给 1×3；A1:B2 给 2×2。
    /// 整行整列（A:A、1:1）无法确定范围，返回 None。
    let relayBlockOf (prefixText: string) (address: string) : RelayBlock option =
        rectOf address
        |> Option.map (fun (r1, c1, r2, c2) ->
            let rows = r2 - r1 + 1
            let cols = c2 - c1 + 1
            let startRef = sprintf "$%s$%d" (colNameOf c1) r1
            let endRef = sprintf "$%s$%d" (colNameOf c2) r2
            // 起止倒写的区域按左上到右下归一化
            let target = if rows = 1 && cols = 1 then startRef else startRef + ":" + endRef
            { Rows = rows
              Cols = cols
              Formula = "=" + prefixText + "!" + target })

    /// 改写文本：块占 rows 行 cols 列，锚点行 startRow、锚点列 A
    let relayTargetText (rows: int) (cols: int) (startRow: int) : string =
        let r2 = startRow + rows - 1
        if rows = 1 && cols = 1 then sprintf "$A$%d" startRow
        elif cols = 1 then sprintf "$A$%d:$A$%d" startRow r2
        else sprintf "$A$%d:$%s$%d" startRow (colNameOf cols) r2

    /// 公式是否「整体就是一个外部单格引用」，即旧式逐格中转格
    let isRelayFormula (formula: string) : bool =
        if String.IsNullOrWhiteSpace formula then false
        else
            let body = formula.Trim()
            let body = if body.StartsWith "=" then body.Substring 1 else body
            match referencesOf body with
            | [ r ] -> r.Start = 0 && r.Length = body.Length && isSingleCell r.Address
            | _ -> false

    /// 公式是否「整体就是一个外部引用」，并给出它占用的形状（行数, 列数）。
    /// 本模块写的中转锚点（单格或区域公式）都满足，其他公式返回 None。
    /// 区域公式的溢出占位格公式为空，只能用锚点的形状推断覆盖范围。
    let relayAnchorShape (formula: string) : (int * int) option =
        if String.IsNullOrWhiteSpace formula then None
        else
            let body = formula.Trim()
            let body = if body.StartsWith "=" then body.Substring 1 else body
            match referencesOf body with
            | [ r ] when r.Start = 0 && r.Length = body.Length ->
                rectOf r.Address
                |> Option.map (fun (r1, c1, r2, c2) -> (r2 - r1 + 1, c2 - c1 + 1))
            | _ -> None

    /// 单元格文本是否本模块写的表尾提示（未中转提示 / 其他说明）
    let isHintText (text: string) : bool =
        if String.IsNullOrWhiteSpace text then false
        else
            let t = text.TrimStart()
            t.StartsWith(HintMark, StringComparison.Ordinal)
            || t.StartsWith(NoteMark, StringComparison.Ordinal)

    // ---------------------------------------------------------------------
    // 构建中转区并改写公式
    // ---------------------------------------------------------------------

    /// 单元格所属「数组单元」：返回（写入目标区域, 是否多格 CSE 数组, 是否该单元左上锚点）。
    /// CSE 多格数组的结果格共用同一条公式，必须整块写入，否则报「不能更改数组的一部分」；
    /// 动态数组只有锚点有真公式，写锚点即可，溢出格写不进去也不用写。
    /// 非数组格：目标是自身、非多格、算锚点。
    let arrayUnitOf (cell: Range) : Range * bool * bool =
        let currentArray =
            try
                let a = cell.CurrentArray
                if a.Count > 1 then Some a else None
            with _ -> None
        match currentArray with
        | None -> cell, false, true
        | Some arr ->
            let isAnchor = cell.Row = arr.Row && cell.Column = arr.Column
            let isCse = try unbox cell.HasArray with _ -> false
            // CSE 数组写整个区域；动态数组写锚点，其余溢出格跳过
            (if isCse then arr else cell), isCse, isAnchor

    /// 在当前工作表建立中转区：表尾从 A 列起按块铺开，原公式改为引用对应区域。
    let relay (ws: Worksheet) : RelayResult =
        let sheetName = ws.Name
        let failed = ResizeArray<string>()
        let mutable rewrittenCells = 0
        let mutable skippedRefs = 0

        let addressOf (cell: Range) =
            try Publishing.cellAddress ws cell
            with _ -> "?"

        let formulaOf (cell: Range) =
            try string (cell.Formula)
            with _ -> ""

        // 1. 公式快照
        let cells =
            [ for cell in Publishing.getFormulaCells ws do
                let f = formulaOf cell
                if not (String.IsNullOrWhiteSpace f) then yield (cell, f) ]

        // 2. 中转区位置：本次运行一律从 UsedRange 末行下方空一行处开始，向下顺序堆叠。
        //    不识别、也不复用表尾上一次写的中转区——重复中转由使用者自行处理。
        let used = ws.UsedRange
        let lastUsedRow = used.Row + used.Rows.Count - 1

        // 3. 收集外部引用，按「前缀 + 绝对地址」去重，按出现顺序分配中转块
        let anchorOfKey = Dictionary<string, int>(StringComparer.Ordinal)
        let newBlocks = ResizeArray<int * RelayBlock>()
        // 改写计划：（写入目标, 是否多格数组需整块写, 原公式, 替换表）
        let plans = ResizeArray<Range * bool * string * (int * int * string) list>()
        let hints = ResizeArray<string>()
        let arrayUnits = ResizeArray<string>()
        let mutable nextRow = lastUsedRow + 2

        for (cell, f) in cells do
            // 数组单元只处理左上锚点，同一数组的其他格共用这条公式，跳过即可
            let target, isArrayRange, isAnchor = arrayUnitOf cell
            if isAnchor then
                let replacements = ResizeArray<int * int * string>()
                for r in referencesOf f do
                    if not (isSelfRef sheetName r) then
                        let prefixText =
                            if r.Text.Length > r.Address.Length then
                                r.Text.Substring(0, r.Text.Length - r.Address.Length - 1)
                            else ""
                        let blockOpt =
                            if r.Sheets.Length = 1 then relayBlockOf prefixText r.Address else None
                        match blockOpt with
                        | Some b when blockFits b ->
                            let key =
                                prefixText.Replace("$", "").ToUpperInvariant()
                                + "!"
                                + absolutizeAddress r.Address
                            let startRow =
                                match anchorOfKey.TryGetValue key with
                                | true, rr -> rr
                                | _ ->
                                    let rr = nextRow
                                    nextRow <- nextRow + b.Rows
                                    newBlocks.Add(rr, b)
                                    anchorOfKey.[key] <- rr
                                    rr
                            replacements.Add(r.Start, r.Length, relayTargetText b.Rows b.Cols startRow)
                        | _ ->
                            skippedRefs <- skippedRefs + 1
                            let reason =
                                match blockOpt with
                                | Some b -> sprintf "区域过大 %d 行 × %d 列" b.Rows b.Cols
                                | None ->
                                    if r.Sheets.Length > 1 then "3D 引用，无法逐格取值"
                                    else "整行/整列，范围不确定"
                            let text = HintMark + r.Text + "：" + reason
                            if not (hints.Contains text) then hints.Add text
                if replacements.Count > 0 then
                    // 只登记真正改写的多格数组，没外部引用的数组不算
                    if isArrayRange then arrayUnits.Add(addressOf cell)
                    plans.Add(target, isArrayRange, f, List.ofSeq replacements)

        // 4. 几何统计
        let relayTotal =
            newBlocks |> Seq.sumBy (fun (_, b) -> b.Rows * b.Cols)
        let newRelayCount = relayTotal
        let areaStart, areaEnd, areaCols =
            if newBlocks.Count = 0 then (0, 0, 0)
            else
                let starts = newBlocks |> Seq.map fst
                let ends = newBlocks |> Seq.map (fun (a, b) -> a + b.Rows - 1)
                let cols = newBlocks |> Seq.map (fun (_, b) -> b.Cols)
                (Seq.min starts, Seq.max ends, Seq.max cols)

        let hintStart =
            if newBlocks.Count > 0 then areaEnd + 1 else nextRow

        // 本表既没有可中转的引用、也没有未中转的引用：在表尾写一行说明，避免用户以为按钮没反应。
        let noExternalRefs =
            newBlocks.Count = 0
            && hints.Count = 0

        // 表尾要写的全部行：未中转提示 + 无外部依赖说明
        let tailTexts =
            List.ofSeq hints @ (if noExternalRefs then [ NoRefText ] else [])

        // 5. 写盘：中转块按分配位置写入，提示行写在本次块之后
        let needWrite = newBlocks.Count > 0 || not tailTexts.IsEmpty

        if needWrite then
            let tailEnd =
                if tailTexts.IsEmpty then 0 else hintStart + tailTexts.Length - 1
            let blockEnd =
                if newBlocks.Count = 0 then 0
                else newBlocks |> Seq.map (fun (a, b) -> a + b.Rows - 1) |> Seq.max
            let lastRow = max blockEnd tailEnd
            if lastRow > ws.Rows.Count then
                invalidOp "工作表底部没有足够空间放置中转单元格。"

            for (anchor, b) in newBlocks do
                let cell = ws.Cells.[anchor, 1] :?> Range
                try
                    // 先腾空目标矩形，再写区域公式：动态数组溢出时撞上已有内容会得到 #SPILL!
                    let rg = ws.Range(ws.Cells.[anchor, 1], ws.Cells.[anchor + b.Rows - 1, b.Cols])
                    rg.ClearContents() |> ignore
                    Formula2.set cell b.Formula
                with _ ->
                    failed.Add(sprintf "%s!A%d" sheetName anchor)

            for i in 0 .. tailTexts.Length - 1 do
                let row = hintStart + i
                let cell = ws.Cells.[row, 1] :?> Range
                try
                    cell.Value2 <- box tailTexts.[i]
                with _ ->
                    failed.Add(sprintf "%s!A%d" sheetName row)

        // 中转格没写全时不动原公式，避免公式指向空格子
        let relayWriteOk = failed.Count = 0

        // 6. 改写原公式
        for (target, isArrayRange, f, replacements) in plans do
            if relayWriteOk then
                let newFormula = rewriteFormula f replacements
                if newFormula <> f then
                    let write (rg: Range) =
                        try
                            Formula2.set rg newFormula
                            true
                        with _ -> false
                    let ok =
                        if isArrayRange then
                            // 多格数组必须整块写入：先试 FormulaArray（保持数组属性），
                            // 再退回整块 Formula2，最后退到锚点单格
                            let anchor = try target.Cells.[1, 1] :?> Range with _ -> target
                            (try
                                target.FormulaArray <- newFormula
                                true
                             with _ -> false)
                            || write target
                            || write anchor
                        else
                            write target
                    if ok then rewrittenCells <- rewrittenCells + 1
                    else failed.Add(addressOf target)

        { NewRelayCount = newRelayCount
          RelayTotal = relayTotal
          RelayStartRow = areaStart
          RelayEndRow = areaEnd
          RelayRows = (if newBlocks.Count > 0 then areaEnd - areaStart + 1 else 0)
          RelayCols = areaCols
          RewrittenCells = rewrittenCells
          SkippedRefs = skippedRefs
          SkippedTexts = (hints |> Seq.map (fun t -> t.Substring HintMark.Length) |> Seq.toArray)
          HintCount = hints.Count
          HintStartRow = (if hints.Count > 0 then hintStart else 0)
          NoRefNoticeRow = (if noExternalRefs then hintStart + hints.Count else 0)
          ArrayUnits = arrayUnits.ToArray()
          FailedCells = failed.ToArray() }
