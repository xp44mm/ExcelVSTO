namespace ExcelNumericalMethods.Test

open System
open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type SelectionCopyTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``format 只输出文件名和工作表地址``() =
        let text = SelectionCopy.format @"C:\数据\当前工作簿.xlsx" "Sheet1" "A1:C3"
        let expected =
            String.concat Environment.NewLine
                [ "\"当前工作簿.xlsx\""
                  "'Sheet1'!A1:C3" ]
        Should.equal text expected

    [<Fact>]
    member this.``format 单格地址不输出当前值``() =
        let text = SelectionCopy.format @"C:\w.xlsx" "2#酸洗线脱硝设备报价" "B7"
        let expected =
            String.concat Environment.NewLine
                [ "\"w.xlsx\""
                  "'2#酸洗线脱硝设备报价'!B7" ]
        Should.equal text expected
        Assert.DoesNotContain("->", text)

    [<Fact>]
    member this.``format 多区域地址原样保留``() =
        let text = SelectionCopy.format @"C:\w.xlsx" "Sheet1" "A1:B2,C3"
        let expected =
            String.concat Environment.NewLine
                [ "\"w.xlsx\""
                  "'Sheet1'!A1:B2,C3" ]
        Should.equal text expected

    [<Fact>]
    member this.``format 工作表名单引号加倍``() =
        let text = SelectionCopy.format @"C:\w.xlsx" "O'Brien" "A1:B1"
        Assert.Contains("'O''Brien'!A1:B1", text)
