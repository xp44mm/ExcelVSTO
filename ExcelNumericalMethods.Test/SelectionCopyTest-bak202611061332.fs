namespace ExcelNumericalMethods.Test

open System
open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type SelectionCopyTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``format 生成与示例一致的剪贴板文本``() =
        let values =
            [ [box "a1"; box "b1"; box "c1"]
              [box "a2"; box "b2"; box "c2"]
              [box "a3"; box "b3"; box "c3"] ]
        let text = SelectionCopy.format @"C:\数据\当前工作簿.xlsx" "Sheet1" "A1:C3" values
        let expected =
            String.concat Environment.NewLine
                [ "\"C:\\数据\\当前工作簿.xlsx\""
                  "'Sheet1'!A1:C3="
                  "["
                  "[\"a1\";\"b1\";\"c1\"];"
                  "[\"a2\";\"b2\";\"c2\"];"
                  "[\"a3\";\"b3\";\"c3\"];"
                  "]" ]
        Should.equal text expected

    [<Fact>]
    member this.``format 混合类型与空单元格保留类型字面量``() =
        let values = [[box 1.0; box true; box null]]
        let text = SelectionCopy.format @"C:\w.xlsx" "Sheet1" "A1:C1" values
        let expected =
            String.concat Environment.NewLine
                [ "\"C:\\w.xlsx\""
                  "'Sheet1'!A1:C1="
                  "["
                  "[1.0;true;null];"
                  "]" ]
        Should.equal text expected

    [<Fact>]
    member this.``format 空选择输出空列表``() =
        let text = SelectionCopy.format @"C:\w.xlsx" "Sheet1" "A1" []
        let expected =
            String.concat Environment.NewLine
                [ "\"C:\\w.xlsx\""
                  "'Sheet1'!A1="
                  "[]" ]
        Should.equal text expected

    [<Fact>]
    member this.``format 工作表名单引号加倍且值内引号转义``() =
        let text = SelectionCopy.format @"C:\w.xlsx" "O'Brien" "A1:B1" [[box "a\"b"; box "x"]]
        Assert.Contains("'O''Brien'!A1:B1=", text)
        Assert.Contains("[\"a\\\"b\";\"x\"];", text)
