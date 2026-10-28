namespace ExcelNumericalMethods

open System
open System.Reflection
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel

/// Excel 公式写入统一入口：通过 IDispatch 后期绑定调用 Range.Formula2。
///
/// 为什么不用强类型的 Range.Formula2：
///   本仓库引用的互操作程序集是 Office15 版（AssemblyVersion 15.0.0.0），其 Range 类型上没有 Formula2；
///   含 Formula2 的互操作程序集只有 16.0.0.0 版本，而 VSTO 宿主链（Microsoft.Office.Tools.*）
///   以 15.0.0.0 的类型标识为准。若把本程序集升级到 16.0.0.0，ExcelVSTO 项目在宿主对象与
///   本程序集之间传递 Range/Worksheet 时会因类型标识不一致而无法编译，故保持 15.0.0.0 不变，
///   改在运行时用后期绑定访问 Formula2。
///
/// 语义差异：Formula2 按 US-English 函数名 + 动态数组（溢出）语法解释公式，与 FormulaLocal 的
/// 本地化语义相对；因此写入的公式文本必须使用英文函数名，不能依赖界面上显示的本地化名称。
/// 运行时要求 Excel 2021 或 Microsoft 365；更早的宿主会抛出 NotSupportedException。
module Formula2 =

    /// DISP_E_MEMBERNOTFOUND（0x80020003）：宿主 Excel 较旧，Range 上没有 Formula2 成员
    let private dispMemberNotFound = -2147352573

    let private unsupported (inner: exn) =
        NotSupportedException(
            "当前 Excel 版本不支持 Range.Formula2（需要 Excel 2021 或 Microsoft 365），公式写入已中止。", inner)

    /// 写入公式，等价于 VBA 的 Range.Formula2 = "=..."。
    /// 公式须使用英文函数名与 A1 引用样式，支持动态数组溢出语法。
    let set (range: Range) (formula: string) : unit =
        if obj.ReferenceEquals(range, null) then nullArg "range"
        if isNull formula then nullArg "formula"
        try
            range.GetType().InvokeMember("Formula2", BindingFlags.SetProperty, null, range, [| box formula |])
            |> ignore
        with
        | :? MissingMethodException as ex -> raise (unsupported ex)
        | :? COMException as ex when ex.HResult = dispMemberNotFound -> raise (unsupported ex)

    /// 管道写法：cell |> Formula2.setWith "=SUM(A1:A3)"
    let setWith (formula: string) (range: Range) : unit = set range formula
