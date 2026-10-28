namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type WorksheetRefsTest(output: ITestOutputHelper) =

    let spansOf (formula: string) =
        WorksheetRefs.referencesOf formula
        |> List.map (fun r -> r.Text, r.Sheets, r.IsExternalBook, r.Address)

    [<Fact>]
    member this.``识别单格跨表引用并记录位置``() =
        let f = "=A1*Sheet2!B2"
        Should.equal (spansOf f) [ ("Sheet2!B2", [ "Sheet2" ], false, "B2") ]
        let r = (WorksheetRefs.referencesOf f).Head
        Should.equal (f.Substring(r.Start, r.Length)) "Sheet2!B2"
        Should.equal r.Start 4

    [<Fact>]
    member this.``同表引用不产生引用片段``() =
        Should.equal (spansOf "=A1*2") []
        Should.equal (spansOf "=SUM(A1:B3)") []

    [<Fact>]
    member this.``识别带引号的工作表名``() =
        Should.equal (spansOf "='My Sheet'!B2") [ ("'My Sheet'!B2", [ "My Sheet" ], false, "B2") ]
        Should.equal (spansOf "='It''s'!A1") [ ("'It''s'!A1", [ "It's" ], false, "A1") ]

    [<Fact>]
    member this.``识别外部工作簿引用``() =
        Should.equal
            (spansOf "=[Book1.xlsx]Sheet1!A1")
            [ ("[Book1.xlsx]Sheet1!A1", [ "Sheet1" ], true, "A1") ]

        Should.equal
            (spansOf "='[Book1.xlsx]Sheet 1'!A1")
            [ ("'[Book1.xlsx]Sheet 1'!A1", [ "Sheet 1" ], true, "A1") ]

    [<Fact>]
    member this.``识别3D引用``() =
        Should.equal (spansOf "=SUM(Sheet1:Sheet3!A1)") [ ("Sheet1:Sheet3!A1", [ "Sheet1"; "Sheet3" ], false, "A1") ]

    [<Fact>]
    member this.``识别区域与整列整行引用``() =
        Should.equal (spansOf "=SUM(Sheet2!A1:B2)") [ ("Sheet2!A1:B2", [ "Sheet2" ], false, "A1:B2") ]
        Should.equal (spansOf "=SUM(Sheet2!A:A)") [ ("Sheet2!A:A", [ "Sheet2" ], false, "A:A") ]
        Should.equal (spansOf "=SUM(Sheet2!1:1)") [ ("Sheet2!1:1", [ "Sheet2" ], false, "1:1") ]

    [<Fact>]
    member this.``跳过字符串字面量``() =
        Should.equal (spansOf "=\"Sheet2!A1\"&A1") []

    [<Fact>]
    member this.``跳过非法地址``() =
        Should.equal (spansOf "=Sheet2!#REF!") []

    [<Fact>]
    member this.``本表判定``() =
        let f = "=Sheet1!A1+'My Sheet'!B2"
        let refs = WorksheetRefs.referencesOf f
        Should.equal (refs |> List.map (fun r -> WorksheetRefs.isSelfRef "sheet1" r)) [ true; false ]
        Should.equal (refs |> List.map (fun r -> WorksheetRefs.isSelfRef "My Sheet" r)) [ false; true ]

    [<Fact>]
    member this.``地址绝对化``() =
        Should.equal (WorksheetRefs.absolutizeAddress "A1") "$A$1"
        Should.equal (WorksheetRefs.absolutizeAddress "$B$2") "$B$2"
        Should.equal (WorksheetRefs.absolutizeAddress "A1:B2") "$A$1:$B$2"
        Should.equal (WorksheetRefs.absolutizeAddress "A:A") "$A:$A"
        Should.equal (WorksheetRefs.absolutizeAddress "1:1") "$1:$1"

    [<Fact>]
    member this.``按位置改写公式``() =
        let f = "=A1*Sheet2!B2+Sheet2!B2"
        let refs = WorksheetRefs.referencesOf f
        let reps =
            refs |> List.map (fun r -> r.Start, r.Length, "$A$100")
        Should.equal (WorksheetRefs.rewriteFormula f reps) "=A1*$A$100+$A$100"

    [<Fact>]
    member this.``识别中转格公式``() =
        Should.equal (WorksheetRefs.isRelayFormula "=Sheet2!$A$1") true
        Should.equal (WorksheetRefs.isRelayFormula "=[Book1.xlsx]Sheet1!$A$1") true
        Should.equal (WorksheetRefs.isRelayFormula "=Sheet2!A1+1") false
        Should.equal (WorksheetRefs.isRelayFormula "=SUM(Sheet2!A1:B2)") false
        Should.equal (WorksheetRefs.isRelayFormula "=A1") false

    [<Fact>]
    member this.``识别未中转提示行``() =
        Should.equal (WorksheetRefs.isHintText "【未中转】Sheet2!A:A：整行/整列，范围不确定") true
        Should.equal (WorksheetRefs.isHintText "  【未中转】X") true
        Should.equal (WorksheetRefs.isHintText "=Sheet2!$A$1") false
        Should.equal (WorksheetRefs.isHintText "") false

    [<Fact>]
    member this.``识别无外部依赖说明行``() =
        Should.equal WorksheetRefs.NoRefText "【提示】没有可中转的外部地址"
        Should.equal (WorksheetRefs.isHintText WorksheetRefs.NoRefText) true
        Should.equal (WorksheetRefs.isHintText "  没有可中转的外部地址") false

    [<Fact>]
    member this.``列号与列名互转``() =
        Should.equal (WorksheetRefs.colNameOf 1) "A"
        Should.equal (WorksheetRefs.colNameOf 26) "Z"
        Should.equal (WorksheetRefs.colNameOf 27) "AA"
        Should.equal (WorksheetRefs.colNameOf 52) "AZ"
        Should.equal (WorksheetRefs.colNameOf 53) "BA"

    [<Fact>]
    member this.``解析矩形地址``() =
        Should.equal (WorksheetRefs.rectOf "C5") (Some(5, 3, 5, 3))
        Should.equal (WorksheetRefs.rectOf "$C$7:$C$9") (Some(7, 3, 9, 3))
        Should.equal (WorksheetRefs.rectOf "C9:C7") (Some(7, 3, 9, 3))
        Should.equal (WorksheetRefs.rectOf "A1:C1") (Some(1, 1, 1, 3))
        Should.equal (WorksheetRefs.rectOf "A1:B2") (Some(1, 1, 2, 2))
        Should.equal (WorksheetRefs.rectOf "B2:A1") (Some(1, 1, 2, 2))
        Should.equal (WorksheetRefs.rectOf "A:A") None
        Should.equal (WorksheetRefs.rectOf "1:1") None

    [<Fact>]
    member this.``单列竖向区域用一条溢出公式中转``() =
        Should.equal
            (WorksheetRefs.relayBlockOf "大气与烟气参数" "$C$7:$C$9")
            (Some
                { Rows = 3
                  Cols = 1
                  Formula = "=大气与烟气参数!$C$7:$C$9" })

        // 起止倒写归一化为左上到右下
        Should.equal
            (WorksheetRefs.relayBlockOf "Sheet2" "C9:C7")
            (Some { Rows = 3; Cols = 1; Formula = "=Sheet2!$C$7:$C$9" })

    [<Fact>]
    member this.``单格中转块``() =
        Should.equal
            (WorksheetRefs.relayBlockOf "Sheet2" "C5")
            (Some { Rows = 1; Cols = 1; Formula = "=Sheet2!$C$5" })

        Should.equal
            (WorksheetRefs.relayBlockOf "Sheet2" "$C$5")
            (Some { Rows = 1; Cols = 1; Formula = "=Sheet2!$C$5" })

    [<Fact>]
    member this.``横向区域用一条溢出公式中转``() =
        let b = (WorksheetRefs.relayBlockOf "Sheet2" "A1:C1").Value
        Should.equal (b.Rows, b.Cols) (1, 3)
        Should.equal b.Formula "=Sheet2!$A$1:$C$1"

    [<Fact>]
    member this.``二维区域用一条溢出公式中转``() =
        let b = (WorksheetRefs.relayBlockOf "Sheet2" "A1:B2").Value
        Should.equal (b.Rows, b.Cols) (2, 2)
        Should.equal b.Formula "=Sheet2!$A$1:$B$2"

    [<Fact>]
    member this.``整行整列无法中转``() =
        Should.equal (WorksheetRefs.relayBlockOf "Sheet2" "A:A") None
        Should.equal (WorksheetRefs.relayBlockOf "Sheet2" "A:B") None
        Should.equal (WorksheetRefs.relayBlockOf "Sheet2" "1:1") None

    [<Fact>]
    member this.``中转块大小限制``() =
        let fits rows cols =
            WorksheetRefs.blockFits
                { Rows = rows
                  Cols = cols
                  Formula = "" }
        Should.equal (fits 1000 1) true
        Should.equal (fits 1001 1) false
        Should.equal (fits 10 100) true
        Should.equal (fits 10 101) false
        Should.equal (fits 200 60) false

    [<Fact>]
    member this.``中转改写文本``() =
        Should.equal (WorksheetRefs.relayTargetText 1 1 11) "$A$11"
        Should.equal (WorksheetRefs.relayTargetText 3 1 11) "$A$11:$A$13"
        Should.equal (WorksheetRefs.relayTargetText 1 3 11) "$A$11:$C$11"
        Should.equal (WorksheetRefs.relayTargetText 2 2 11) "$A$11:$B$12"
        Should.equal (WorksheetRefs.relayTargetText 2 3 11) "$A$11:$C$12"

    [<Fact>]
    member this.``煤质表真实公式的区域引用被识别``() =
        let f = "=理论空气量(大气与烟气参数!$C$7:$C$9,C14)"
        let refs = WorksheetRefs.referencesOf f
        Should.equal
            (refs |> List.map (fun r -> r.Text, r.Sheets, r.Address))
            [ ("大气与烟气参数!$C$7:$C$9", [ "大气与烟气参数" ], "$C$7:$C$9") ]

        let r = refs.Head
        Should.equal (WorksheetRefs.isSelfRef "煤质" r) false
        Should.equal
            (WorksheetRefs.relayBlockOf "大气与烟气参数" r.Address)
            (Some
                { Rows = 3
                  Cols = 1
                  Formula = "=大气与烟气参数!$C$7:$C$9" })

    [<Fact>]
    member this.``识别中转锚点及其覆盖形状``() =
        // 区域公式锚点：3 行 1 列
        Should.equal
            (WorksheetRefs.relayAnchorShape "=大气与烟气参数!$C$7:$C$9")
            (Some(3, 1))
        // 二维区域：2 × 2
        Should.equal (WorksheetRefs.relayAnchorShape "=Sheet2!$A$1:$B$2") (Some(2, 2))
        // 单格中转
        Should.equal (WorksheetRefs.relayAnchorShape "=Sheet2!$C$5") (Some(1, 1))
        // 溢出占位格公式为空
        Should.equal (WorksheetRefs.relayAnchorShape "") None
        // 普通公式、拼接公式都不是锚点
        Should.equal (WorksheetRefs.relayAnchorShape "=Sheet2!$C$5+1") None
        Should.equal (WorksheetRefs.relayAnchorShape "=SUM(Sheet2!$C$5:$C$9)") None
        // 整列引用无法确定形状
        Should.equal (WorksheetRefs.relayAnchorShape "=Sheet2!$A:$A") None
