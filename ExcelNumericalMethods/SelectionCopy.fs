///「复制选择」：把当前选择的区域地址与当前值，以
/// "工作簿全路径"
/// '工作表名'!A1:C3->
/// [
/// [a1;b1;c1];
/// [a2;b2;c2];
/// [a3;b3;c3];
/// ]
/// 格式写入系统剪贴板。二维数组是 Value2 当前值，不含公式；
/// 单格选择时数组简化为该值的标量字面量。
module ExcelNumericalMethods.SelectionCopy

open System
open Microsoft.Office.Interop.Excel
open FSharp.Idioms.Literal

/// 读取单元格当前值。错误单元格的 Value2 是负 Int32 错误码，改取显示文本（如 #N/A）。
let private valueOf (cell: Range) : obj =
    try
        match cell.Value2 with
        | :? int as n when n < 0 -> box (string cell.Text)
        | v -> v
    with _ -> null

/// 选择区域地址：A1 样式、不带 $；多区域由 Excel 以逗号连接。
let addressOf (sel: Range) : string =
    string (sel.get_Address(false, false))

/// 遍历选择的全部单元格（支持多区域、行优先），生成按行分组的当前值二维数组。
/// 空单元格为 null，字符串/数字/布尔保留原类型，便于以 F# 字面量输出。
let valuesOf (sel: Range) : obj list list =
    let rows = ResizeArray<obj list>()
    for a in 1 .. sel.Areas.Count do
        let area = sel.Areas.[a]
        for r in 1 .. area.Rows.Count do
            let row = ResizeArray<obj>()
            for c in 1 .. area.Columns.Count do
                let cell = area.Cells.[r, c] :?> Range
                row.Add(valueOf cell)
            rows.Add(List.ofSeq row)
    List.ofSeq rows

/// 拼装剪贴板文本：
/// "工作簿全路径"
/// '工作表名'!A1:C3->
/// [
/// [a1;b1;c1];
/// [a2;b2;c2];
/// ]
/// 工作表名内的单引号按 Excel 惯例加倍；每行用 stringify 生成 F# list 字面量；
/// 单格时数据块直接输出标量，空选择输出 []。
let format
    (workbookPath: string)
    (sheetName: string)
    (address: string)
    (values: obj list list)
    : string
    =
    let nl = Environment.NewLine
    let sheet = sheetName.Replace("'", "''")
    let header = sprintf "\"%s\"%s'%s'!%s->" workbookPath nl sheet address
    let body =
        match values with
        | [[v]] -> stringify v
        | [] -> "[]"
        | _ ->
            let rows = values |> List.map (fun row -> stringify row + ";")
            "[" + nl + String.concat nl rows + nl + "]"
    header + nl + body + ";;"

/// 将文本写入系统剪贴板；剪贴板被其他程序占用时重试数次。
let setClipboardText (text: string) : unit =
    let rec retry attempt =
        if attempt <= 0 then
            raise(Runtime.InteropServices.ExternalException "无法写入剪贴板，可能被其他程序占用，请重试。")
        else
            try
                System.Windows.Forms.Clipboard.SetDataObject(text, true, 5, 100)
            with :? Runtime.InteropServices.ExternalException ->
                System.Threading.Thread.Sleep(100)
                retry(attempt - 1)
    retry 5

/// 把当前选择复制到剪贴板，返回生成的文本（供状态栏提示等使用）。
let copySelection (sel: Range) (workbookPath: string) (sheetName: string) : string =
    let text = format workbookPath sheetName (addressOf sel) (valuesOf sel)
    setClipboardText text
    text
