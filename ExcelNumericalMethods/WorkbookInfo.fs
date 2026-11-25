///「复制当前工作簿信息」：把当前 Excel 实例的后台定位信息（采集时间 / ROT 注册名）
/// 与当前工作簿的名称、完整路径写入剪贴板，供大模型精确定位并访问这个工作簿。
module ExcelNumericalMethods.WorkbookInfo

open System
open System.Diagnostics
open System.Globalization
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open System.Text
open Microsoft.Office.Interop.Excel

[<DllImport("ole32.dll")>]
extern int GetRunningObjectTable(int reserved, [<Out>] IRunningObjectTable& pprot)

[<DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)>]
extern int CreateItemMoniker(string lpszDelim, string lpszObject, [<Out>] IMoniker& ppmk)

/// 本实例在 ROT 中的注册名；进程内只注册一次。
let mutable private rotName : string = null

/// 把当前 Excel Application 以 !ExcelVSTO.Instance.<PID> 注册到 ROT。
/// flag=0 为弱注册：ROT 不持有强引用，Excel 退出后条目自动消失。
let private ensureRotRegistration (app: Application) (pid: int) =
    if not (isNull rotName) then rotName
    else
        let name = "!ExcelVSTO.Instance." + pid.ToString(CultureInfo.InvariantCulture)
        let mutable table = Unchecked.defaultof<IRunningObjectTable>
        let hr = GetRunningObjectTable(0, &table)
        if hr < 0 then Marshal.ThrowExceptionForHR(hr)
        let mutable moniker = Unchecked.defaultof<IMoniker>
        let hr2 = CreateItemMoniker("!", name, &moniker)
        if hr2 < 0 then Marshal.ThrowExceptionForHR(hr2)
        table.Register(0, box app, moniker) |> ignore
        rotName <- name
        name

let private safeGet (f: unit -> string) =
    try
        match f () with
        | null -> "(无)"
        | s when String.IsNullOrEmpty s -> "(无)"
        | s -> s
    with _ -> "(读取失败)"

/// 生成剪贴板文本。输出行：采集时间 / ROT / 工作簿 / 路径。
let private formatText (app: Application) (pid: int) =
    let name = ensureRotRegistration app pid
    let sb = StringBuilder()
    sb.AppendLine("[ExcelVSTO 当前 Excel 实例] 采集时间 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)) |> ignore
    sb.AppendLine("ROT : " + name) |> ignore
    match app.ActiveWorkbook with
    | null ->
        sb.AppendLine("工作簿 : (无打开的工作簿)") |> ignore
        sb.AppendLine("路径 : (无)") |> ignore
    | wb ->
        sb.AppendLine("工作簿 : " + safeGet (fun () -> wb.Name)) |> ignore
        sb.AppendLine("路径 : " + safeGet (fun () -> if String.IsNullOrEmpty(wb.Path) then "(未保存)" else wb.FullName)) |> ignore
    sb.ToString()

/// 复制当前工作簿信息到剪贴板；异常在模块内提示。
let copyCurrent (app: Application) =
    try
        use proc = Process.GetCurrentProcess()
        SelectionCopy.setClipboardText (formatText app proc.Id)
    with ex ->
        System.Windows.Forms.MessageBox.Show(
            ex.Message,
            "复制当前工作簿",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
        |> ignore
