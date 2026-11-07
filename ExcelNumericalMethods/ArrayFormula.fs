namespace ExcelNumericalMethods

open System
open System.Reflection
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel

/// 数组公式（旧式 CSE 与新式溢出）的判定与转换。
///
/// 两种数组公式在 Excel COM 上的实际表现（Excel 365 实测，见 tools/ArrayProbe）：
///
/// | 项目              | 旧式 CSE（Ctrl+Shift+Enter）        | 新式溢出（spill）                    |
/// |-------------------|------------------------------------|--------------------------------------|
/// | HasArray          | true                               | false                                |
/// | HasSpill          | false                              | true（锚点与溢出区单元格都为 true）  |
/// | SpillingToRange   | 空                                 | 仅锚点非空                           |
/// | CurrentArray      | 整个数组区域                       | 抛异常「未找到单元格」               |
/// | 非主单元格 Formula| 与主单元格相同的完整公式           | 空字符串                             |
///
/// 关键风险：向溢出区的非锚点单元格写入（Formula2 / Value2）不会报错，而是静默破坏溢出区，
/// 使锚点变成 #SPILL!。因此凡是要写入单元格的地方，都必须先用本模块的 isArrayMaster 过滤。
///
/// HasSpill / SpillingToRange 只存在于 Excel 2021 与 Microsoft 365 的对象模型上，
/// 本仓库的互操作程序集是 Office15 版（无这两个成员），故一律用后期绑定访问，缺失时按「非溢出」处理。
module ArrayFormula =

    /// DISP_E_MEMBERNOTFOUND（0x80020003）：宿主 Excel 较旧，对象上没有该成员
    let private dispMemberNotFound = -2147352573

    /// 后期绑定读取属性；成员不存在、取到空或读取失败时返回 None。
    let private tryGetProperty (target: obj) (name: string) : obj option =
        try
            match target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null) with
            | null -> None
            | :? Missing -> None
            | value -> Some value
        with
        | :? MissingMethodException -> None
        | :? COMException as ex when ex.HResult = dispMemberNotFound -> None
        | _ -> None

    /// 单元格当前是否为错误值。
    /// 本机 Excel 的 COM 互操作读取错误单元格不抛异常，而是返回负 Int32 错误码；
    /// 正常单元格经 Value2 读取只会是 double / string / bool，不会出现负 Int32。
    /// （与 Publishing.isErrorCell 同义，此处独立实现以避免模块循环依赖。）
    let isErrorValue (cell: Range) : bool =
        try
            match cell.Value2 with
            | :? int as n -> n < 0
            | _ -> false
        with _ -> true

    // ---------- 判定 ----------

    /// 是否为旧式 CSE 数组公式（Ctrl+Shift+Enter 录入）的一部分。
    let hasLegacyArray (cell: Range) : bool =
        try unbox cell.HasArray with _ -> false

    /// 是否处于新式溢出（spill）数组公式的区域内：锚点与溢出区单元格都为 true。
    /// 宿主不支持 HasSpill 时恒为 false。
    let hasSpill (cell: Range) : bool =
        match tryGetProperty cell "HasSpill" with
        | Some value -> try unbox value with _ -> false
        | None -> false

    /// 单元格是否属于数组公式（旧式 CSE 或新式溢出）。
    /// 属于数组公式意味着它不能单独修改：旧式要整块改，新式只能改锚点。
    let isArrayPart (cell: Range) : bool =
        hasLegacyArray cell || hasSpill cell

    /// 溢出区域（仅锚点单元格取得到；非锚点与不支持溢出语义的宿主返回 None）。
    let spillRange (cell: Range) : Range option =
        match tryGetProperty cell "SpillingToRange" with
        | Some (:? Range as range) -> Some range
        | _ -> None

    /// 是否为溢出区域的锚点单元格（溢出公式真正所在的格子）。
    let isSpillAnchor (cell: Range) : bool =
        (spillRange cell).IsSome

    /// 单元格所属数组公式的整体区域；不属于任何数组公式时返回 None。
    /// 旧式 CSE 取 CurrentArray；新式溢出只有锚点能取到 SpillingToRange。
    let arrayRange (cell: Range) : Range option =
        if hasLegacyArray cell then
            try Some cell.CurrentArray with _ -> None
        else spillRange cell

    /// 是否为数组公式区域中可写的那个单元格：
    /// 旧式 CSE 的左上角、新式溢出的锚点、普通公式恒为 true。
    /// 区域内的其余单元格返回 false，须跳过，由主单元格统一处理。
    let isArrayMaster (cell: Range) : bool =
        if hasLegacyArray cell then
            try
                let area = cell.CurrentArray
                area.Row = cell.Row && area.Column = cell.Column
            with _ -> true
        elif hasSpill cell then
            (spillRange cell).IsSome
        else
            true

    /// 数组公式区域的规模（行数, 列数）；普通公式与取不到区域时为 (1, 1)。
    let arraySize (cell: Range) : int * int =
        match arrayRange cell with
        | Some range ->
            try (range.Rows.Count, range.Columns.Count)
            with _ -> (1, 1)
        | None -> (1, 1)

    /// 数组公式区域是否为多格（需要按块读写、兜底值要用数组常数）。
    let isMultiCellArray (cell: Range) : bool =
        let (rows, cols) = arraySize cell
        rows > 1 || cols > 1

    // ---------- 旧式 CSE 一键转换为新式溢出 ----------

    /// 旧式 CSE 数组公式转换为新式溢出公式的结果
    type ConvertResult =
        { /// 成功转换的数组公式区域地址（工作表!地址）
          ConvertedAddresses: string array
          /// 转换失败已回滚的（工作表!地址, 原因）
          FailedCells: (string * string) array
          /// 受保护且无法解除保护而跳过的工作表
          ProtectedSheets: string array }

    /// 把工作簿内全部旧式 CSE 数组公式（Ctrl+Shift+Enter 录入）一键转换为新式溢出公式。
    ///
    /// 做法：先收集全部旧式数组区域地址，再逐个「整块 ClearContents + 锚点写 Formula2」，
    /// 由 Excel 自动溢出。先收集后转换是必须的——转换会改变公式单元格集合，边遍历边改会漏项。
    ///
    /// 转换后锚点若为错误值（典型是溢出被相邻数据阻挡变成 #SPILL!），自动回滚为原来的旧式数组公式并记为失败。
    /// 工作簿受保护时先尝试无密码解除保护，处理完再按原样保护回去。
    let convertCseToSpill (wb: Workbook) : ConvertResult =
        let converted = ResizeArray<string>()
        let failed = ResizeArray<string * string>()
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
                // 第一阶段：收集本表全部旧式数组区域（只取左上角，去重）
                let addresses = ResizeArray<string>()
                let seen = Collections.Generic.HashSet<string>()
                for cell in Traversal.getFormulaCells ws do
                    if hasLegacyArray cell then
                        try
                            let area = cell.CurrentArray
                            if area.Row = cell.Row && area.Column = cell.Column then
                                let address = string (area.Address())
                                if seen.Add(address) then addresses.Add address
                        with _ -> ()
                // 第二阶段：逐个转换
                for address in addresses do
                    let area = ws.Range(address)
                    let anchor = area.Cells.[1, 1] :?> Range
                    let formula = string anchor.Formula
                    let label = sprintf "%s!%s" ws.Name address
                    try
                        area.ClearContents() |> ignore
                        Formula2.set anchor formula
                        if isErrorValue anchor then
                            // 溢出被阻挡：回滚为原来的旧式数组公式
                            area.FormulaArray <- formula
                            failed.Add(label, "转换后锚点为错误值（多半是溢出被相邻数据阻挡），已回滚为旧式数组公式")
                        else
                            converted.Add label
                    with ex ->
                        // 任何异常都尽量回滚成原样，再记为失败
                        try area.FormulaArray <- formula
                        with _ -> ()
                        failed.Add(label, ex.Message)
            if wasProtected then
                try ws.Protect()
                with _ -> ()
        { ConvertedAddresses = converted.ToArray()
          FailedCells = failed.ToArray()
          ProtectedSheets = protectedSheets.ToArray() }
