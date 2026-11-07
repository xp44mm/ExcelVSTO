namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type SelectionCopyTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``format 生成与示例一致的剪贴板文本``() =
        let entries = [ ("A1", "=0"); ("A2", "=A1") ]
        let text = SelectionCopy.format @"C:\数据\当前工作簿.xlsx" "Sheet1" entries
        Should.equal text @"当前工作簿【C:\数据\当前工作簿.xlsx】工作表【Sheet1】选择：[""A1"",""=0"";""A2"",""=A1""]"

    [<Fact>]
    member this.``format 空选择输出空列表``() =
        let text = SelectionCopy.format @"C:\数据\工作簿1.xlsx" "Sheet1" []
        Should.equal text @"当前工作簿【C:\数据\工作簿1.xlsx】工作表【Sheet1】选择：[]"

    [<Fact>]
    member this.``format 含引号公式时保持条目可解析``() =
        let entries = [ ("A1", "=IF(B1=\"x\",1,0)") ]
        let text = SelectionCopy.format @"C:\数据\工作簿1.xlsx" "Sheet1" entries
        Should.equal text @"当前工作簿【C:\数据\工作簿1.xlsx】工作表【Sheet1】选择：[""A1"",""=IF(B1=\""x\"",1,0)""]"
