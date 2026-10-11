///「复制选择」：把当前选择的单元格地址与公式字符串，以
/// 工作簿【工作簿名】当前选择：["A1","=0";"A2","=A1"]
/// 格式写入系统剪贴板。
module ExcelNumericalMethods.SelectionCopy

open System
open System.Globalization
open Microsoft.Office.Interop.Excel
open FSharp.Idioms.Literal

/// 单个单元格的 "地址","公式" 条目；
/// 公式字符串内部的双引号按 CSV 惯例加倍转义（" -> ""），保证含字符串字面量的公式输出仍可解析。
let private entryOf (cell: Range) =
    let address = cell.get_Address(false, false)
    let formula =
        match box cell.Formula with
        | null -> ""
        | :? string as s -> s
        | v -> Convert.ToString(v, CultureInfo.InvariantCulture)
    address,formula

/// 遍历选择的全部单元格（支持多区域、行优先），生成 "地址","公式" 条目列表。
/// 地址取不带 $ 的 A1 样式（如 A1）；公式取 A1 样式公式字符串（如 =A1），常量单元格输出值字符串，空单元格输出空串。
let entriesOf (sel: Range) : (string * string) list =
    let entries = ResizeArray<string*string>()
    for a in 1 .. sel.Areas.Count do
        let area = sel.Areas.[a]
        for r in 1 .. area.Rows.Count do
            for c in 1 .. area.Columns.Count do
                let cell = area.Cells.[r, c] :?> Range
                entries.Add(entryOf cell)
    List.ofSeq entries

/// 拼装剪贴板文本：工作簿【工作簿名】当前选择：[条目1;条目2;...]
let format (workbookName: string) (entries: (string * string) list) : string =
    sprintf "工作簿【%s】当前选择：%s" workbookName (stringify entries)

/// 将文本写入系统剪贴板；剪贴板被其他程序占用时重试数次。
let setClipboardText (text: string) : unit =
    let rec retry attempt =
        if attempt <= 0 then
            raise (Runtime.InteropServices.ExternalException "无法写入剪贴板，可能被其他程序占用，请重试。")
        else
            try
                System.Windows.Forms.Clipboard.SetDataObject(text, true, 5, 100)
            with :? Runtime.InteropServices.ExternalException ->
                System.Threading.Thread.Sleep(100)
                retry (attempt - 1)
    retry 5

/// 把当前选择复制到剪贴板，返回生成的文本（供状态栏提示等使用）。
let copySelection (sel: Range) (workbookName: string) : string =
    let text = format workbookName (entriesOf sel)
    setClipboardText text
    text
