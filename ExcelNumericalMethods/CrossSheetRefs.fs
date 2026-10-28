///「标记跨表引用」：扫描当前工作簿的全部公式，凡是「本工作表以外」的单元格引用了某个单元格，
/// 就把被引用的单元格背景色标成独特颜色（洋红），用于快速识别对外提供数据的源头单元格。
/// 只处理工作簿内部的「工作表!地址」引用；外部工作簿引用（[Book1.xlsx]Sheet1!A1）与名称引用不处理。
module ExcelNumericalMethods.CrossSheetRefs

open System
open System.Collections.Generic
open System.Text.RegularExpressions
open Microsoft.Office.Interop.Excel

/// 标记色：洋红 RGB(255,0,255)；Excel COM 的 Interior.Color 为 BGR 顺序，故为 0x00FF00FF
let [<Literal>] MarkColor = 0x00FF00FF

/// 一处「工作表!地址」引用
type Ref =
    { /// 引用中写出的工作表名；3D 引用（Sheet1:Sheet3!A1）为起止两端两个名字
      Sheets: string list
      /// 是否 3D 引用，需按工作表顺序展开
      Is3D: bool
      /// 被引用的区域地址（A1、A1:B2、A:A、1:1）
      Address: string }

/// 标记结果
type MarkResult =
    { /// 被标色的引用区域数（同一区域多次引用只计一次）
      MarkedAreas: int
      /// 被标色的工作表名
      MarkedSheets: string array
      /// 公式中引用但工作簿里不存在的工作表名
      MissingSheets: string array
      /// 受保护而无法标色的工作表名
      ProtectedSheets: string array }

// ---------------------------------------------------------------------------
// 公式解析：只做「工作表名!地址」的文本识别，不建立完整表达式树。
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

/// 工作表名（不加引号时）允许的字符：字母、数字、下划线、点
let private isNameChar (c: char) =
    Char.IsLetterOrDigit c || c = '_' || c = '.'

/// 从 endIdx 位置向前读出一个工作表名，返回（名称，名称首字符下标）。
/// 支持两种写法：不加引号的 Sheet2，加引号的 'My Sheet'（内部 '' 表示一个单引号）。
let private readNameBack (s: string) (endIdx: int) : (string * int) option =
    if endIdx < 0 || endIdx >= s.Length then None
    elif s.[endIdx] = '\'' then
        // 引号名：向前找配对的起始引号，跳过成对的 '' 转义
        let mutable j = endIdx - 1
        let mutable start = -1
        while j >= 0 && start < 0 do
            if s.[j] = '\'' then
                if j - 1 >= 0 && s.[j - 1] = '\'' then j <- j - 2
                else start <- j
            else
                j <- j - 1
        if start < 0 then None
        else Some(s.Substring(start + 1, endIdx - start - 1).Replace("''", "'"), start)
    else
        let mutable j = endIdx
        while j >= 0 && isNameChar s.[j] do
            j <- j - 1
        if j = endIdx then None
        else Some(s.Substring(j + 1, endIdx - j), j + 1)

/// 从公式中解析出全部「工作表!地址」引用。
/// 跳过字符串字面量（含 "" 转义）、外部工作簿引用与 #REF! 等非法地址。
let referencesOf (formula: string) : Ref list =
    let result = ResizeArray<Ref>()
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
        else
            if c = '!' then
                let names = ResizeArray<string>()
                let mutable ok = true
                match readNameBack s (i - 1) with
                | None -> ok <- false
                | Some(name, startIdx) ->
                    names.Add name
                    let mutable k = startIdx - 1
                    // 3D 引用：Sheet1:Sheet3!A1
                    if k >= 0 && s.[k] = ':' then
                        match readNameBack s (k - 1) with
                        | Some(name2, startIdx2) ->
                            names.Insert(0, name2)
                            k <- startIdx2 - 1
                        | None -> ok <- false
                    // 外部工作簿引用：[Book1.xlsx]Sheet1!A1
                    if k >= 0 && s.[k] = ']' then ok <- false
                if ok then
                    let rest = s.Substring(i + 1)
                    let m =
                        let m1 = areaRgx.Match rest
                        if m1.Success then m1
                        else
                            let m2 = colsRgx.Match rest
                            if m2.Success then m2
                            else rowsRgx.Match rest
                    if m.Success then
                        // 引号名内部也可能含冒号（'Sheet 1:Sheet 3'!A1）与外部前缀，统一在此拆分过滤
                        let expanded =
                            names
                            |> Seq.collect (fun nm -> nm.Split(':'))
                            |> Seq.map (fun nm -> nm.Trim())
                            |> Seq.filter (fun nm -> nm.Length > 0)
                            |> Seq.toList
                        let hasExternal =
                            expanded |> List.exists (fun nm -> nm.Contains "[" || nm.Contains "]")
                        if not hasExternal && not expanded.IsEmpty then
                            result.Add
                                { Sheets = expanded
                                  Is3D = expanded.Length > 1
                                  Address = m.Groups.[1].Value }
            i <- i + 1
    List.ofSeq result

// ---------------------------------------------------------------------------
// 工作簿扫描与标色
// ---------------------------------------------------------------------------

/// 工作簿中的工作表（按工作表标签顺序；跳过图表工作表等非 Worksheet 对象）
let private worksheetsOf (wb: Workbook) : Worksheet list =
    let list = ResizeArray<Worksheet>()
    for i in 1 .. wb.Worksheets.Count do
        match box (wb.Worksheets.[i]) with
        | :? Worksheet as ws -> list.Add ws
        | _ -> ()
    List.ofSeq list

/// 给一个工作表上的若干地址涂色。地址拼成长串一次涂色；失败时退化为逐个单元格涂色，
/// 仍失败的地址记入 failed。
let private applyColor (ws: Worksheet) (addresses: string list) (failed: ResizeArray<string>) =
    let chunkSize = 50
    let rec loop (rest: string list) =
        match rest with
        | [] -> ()
        | _ ->
            let take = min chunkSize (List.length rest)
            let head = List.take take rest
            let tail = List.skip take rest
            try
                let area = ws.Range(String.Join(",", List.toArray head))
                area.Interior.Color <- float MarkColor
            with _ ->
                for addr in head do
                    try
                        let cell = ws.Range(addr)
                        cell.Interior.Color <- float MarkColor
                    with _ -> failed.Add addr
            loop tail
    loop addresses

/// 扫描当前工作簿，给「被其他工作表引用」的单元格涂上标记色。
/// 公式所在工作表与被引用工作表相同（含显式写自身表名，如 Sheet1 里写 Sheet1!A1）不标记。
let mark (wb: Workbook) : MarkResult =
    let sheets = worksheetsOf wb

    // 名称 -> 工作表（工作表名不区分大小写）；同名只取第一个
    let byName = Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase)
    for ws in sheets do
        if not (byName.ContainsKey ws.Name) then byName.[ws.Name] <- ws

    // 工作表标签顺序，用于展开 3D 引用
    let order = sheets |> List.indexed

    let missing = HashSet<string>(StringComparer.OrdinalIgnoreCase)
    // 目标工作表 -> 待涂色地址集合
    let targets = Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase)

    /// 把一条引用解析成目标工作表列表
    let resolve (r: Ref) : Worksheet list =
        match r.Sheets with
        | [] -> []
        | [one] ->
            match byName.TryGetValue one with
            | true, ws -> [ ws ]
            | _ ->
                missing.Add one |> ignore
                []
        | first :: last :: _ ->
            match byName.TryGetValue first, byName.TryGetValue last with
            | (true, a), (true, b) ->
                let ia, ib =
                    let find (name: string) =
                        order |> List.pick (fun (i, w) -> if String.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase) then Some i else None)
                    find a.Name, find b.Name
                let lo, hi = min ia ib, max ia ib
                order |> List.filter (fun (i, _) -> i >= lo && i <= hi) |> List.map snd
            | _ ->
                missing.Add first |> ignore
                missing.Add last |> ignore
                []

    for ws in sheets do
        for cell in Publishing.getFormulaCells ws do
            let formula =
                try string (cell.Formula)
                with _ -> ""
            if formula.Length > 0 then
                for r in referencesOf formula do
                    for target in resolve r do
                        if not (String.Equals(target.Name, ws.Name, StringComparison.OrdinalIgnoreCase)) then
                            let set =
                                match targets.TryGetValue target.Name with
                                | true, s -> s
                                | _ ->
                                    let s = HashSet<string>(StringComparer.Ordinal)
                                    targets.[target.Name] <- s
                                    s
                            set.Add r.Address |> ignore

    let markedSheets = ResizeArray<string>()
    let protectedSheets = ResizeArray<string>()
    let mutable markedAreas = 0
    for KeyValue(name, set) in targets do
        match byName.TryGetValue name with
        | true, ws ->
            let addresses = List.ofSeq set
            let failed = ResizeArray<string>()
            applyColor ws addresses failed
            let ok = List.length addresses - failed.Count
            if ok > 0 then
                markedSheets.Add ws.Name
                markedAreas <- markedAreas + ok
            if failed.Count > 0 && ok = 0 then protectedSheets.Add ws.Name
        | _ -> ()

    { MarkedAreas = markedAreas
      MarkedSheets = markedSheets.ToArray()
      MissingSheets = Seq.toArray missing
      ProtectedSheets = protectedSheets.ToArray() }
