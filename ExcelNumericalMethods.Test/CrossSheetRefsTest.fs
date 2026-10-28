namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type CrossSheetRefsTest(output : ITestOutputHelper) =

    let refsOf (formula: string) =
        CrossSheetRefs.referencesOf formula
        |> List.map (fun r -> r.Sheets, r.Is3D, r.Address)

    [<Fact>]
    member this.``同表引用与常量不算跨表引用``() =
        Should.equal (refsOf "=A1*2") []
        Should.equal (refsOf "=SUM(A1:B3)") []

    [<Fact>]
    member this.``识别简单跨表引用``() =
        Should.equal (refsOf "=Sheet2!A1") [ ([ "Sheet2" ], false, "A1") ]

    [<Fact>]
    member this.``识别带引号的工作表名``() =
        Should.equal (refsOf "='My Sheet'!B2:C3") [ ([ "My Sheet" ], false, "B2:C3") ]
        Should.equal (refsOf "='It''s'!A1") [ ([ "It's" ], false, "A1") ]

    [<Fact>]
    member this.``识别整列与整行引用``() =
        Should.equal (refsOf "=SUM(Sheet2!A:A)") [ ([ "Sheet2" ], false, "A:A") ]
        Should.equal (refsOf "=SUM(Sheet2!1:1)") [ ([ "Sheet2" ], false, "1:1") ]

    [<Fact>]
    member this.``识别 3D 引用``() =
        Should.equal (refsOf "=SUM(Sheet1:Sheet3!A1)") [ ([ "Sheet1"; "Sheet3" ], true, "A1") ]
        Should.equal (refsOf "=SUM('Sheet 1:Sheet 3'!A1)") [ ([ "Sheet 1"; "Sheet 3" ], true, "A1") ]

    [<Fact>]
    member this.``跳过字符串字面量中的感叹号``() =
        Should.equal (refsOf "=\"Sheet2!A1\"&A1") []

    [<Fact>]
    member this.``跳过外部工作簿引用``() =
        Should.equal (refsOf "=[Book1.xlsx]Sheet1!A1") []
        Should.equal (refsOf "='[Book1.xlsx]Sheet1'!A1") []

    [<Fact>]
    member this.``跳过非法地址``() =
        Should.equal (refsOf "=Sheet2!#REF!") []

    [<Fact>]
    member this.``一个公式内多个跨表引用``() =
        Should.equal
            (refsOf "=IF(A1>0,Sheet2!B2,'My Sheet'!$C$3)")
            [ ([ "Sheet2" ], false, "B2"); ([ "My Sheet" ], false, "$C$3") ]
