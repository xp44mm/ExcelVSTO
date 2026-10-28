namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open FSharp.xUnit

type RelayDeleteTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``一元表达式：字面量与引用``() =
        for b in [ "A1"; "$A$1"; "Sheet2!$B$2"; "'My Sheet'!A1"; "A1:B2"; "A:A"; "MyName"
                   "123"; "1.5"; ".5"; "1.5E-3"; "\"abc\""; "TRUE"; "FALSE"; "(A1+B1)"; "{1,2;3,4}" ] do
            Should.equal (RelayDelete.isUnitExpression b) true

    [<Fact>]
    member this.``一元表达式：函数调用``() =
        for b in [ "SUM(A1:B3)"; "IFERROR(SUM(A1),0)"; "LOG10(A1)"; "@SUM(A1)"; "'My Sheet'!A1" ] do
            Should.equal (RelayDelete.isUnitExpression b) true

    [<Fact>]
    member this.``一元表达式：顶层运算符一律不受理``() =
        for b in [ "A1+B1"; "-A1"; "+A1"; "A1%"; "A1*2"; "A1&\"x\""; "A1=1"; "SUM(A1)+1"
                   "SUM(A1)*2"; "(A1)+(B1)"; "A1 A2"; ""; "  " ] do
            Should.equal (RelayDelete.isUnitExpression b) false

    [<Fact>]
    member this.``引用解析：本表与跨表``() =
        let refs = RelayDelete.referencesOf "=A1*Sheet2!B2"
        Should.equal (refs |> List.map (fun r -> r.Text, r.Sheets, r.Address))
            [ ("A1", [], "A1"); ("Sheet2!B2", [ "Sheet2" ], "B2") ]

    [<Fact>]
    member this.``引用解析：区域与整列整行``() =
        let refs = RelayDelete.referencesOf "=SUM($A$40:$A$42)+SUM(A:A)+SUM(1:1)"
        Should.equal (refs |> List.map (fun r -> r.Text))
            [ "$A$40:$A$42"; "A:A"; "1:1" ]

    [<Fact>]
    member this.``引用解析：跳过字符串与函数名``() =
        Should.equal (RelayDelete.referencesOf "=\"A1\"&B1" |> List.map (fun r -> r.Text)) [ "B1" ]
        Should.equal (RelayDelete.referencesOf "=LOG10(A1)" |> List.map (fun r -> r.Text)) [ "A1" ]

    [<Fact>]
    member this.``引用解析：引号名、3D、外部工作簿``() =
        let q = RelayDelete.referencesOf "='My Sheet'!$B$2"
        Should.equal (q |> List.map (fun r -> r.Sheets)) [ [ "My Sheet" ] ]
        let d = RelayDelete.referencesOf "=SUM(Sheet1:Sheet3!A1)"
        Should.equal (d |> List.map (fun r -> r.Sheets, r.Is3D)) [ ([ "Sheet1"; "Sheet3" ], true) ]
        let e = RelayDelete.referencesOf "=[Book1.xlsx]Sheet1!A1"
        Should.equal (e |> List.map (fun r -> r.BookName)) [ Some "Book1.xlsx" ]

    [<Fact>]
    member this.``内联文本：同表原样，跨表补表前缀``() =
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet1" "SUM(A1:B3)") "SUM(A1:B3)"
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet2" "SUM(A1:B3)") "SUM(Sheet1!A1:B3)"
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet2" "A1*2+$B$2") "Sheet1!A1*2+Sheet1!$B$2"
        // 已写表名、外部工作簿、字面量不加前缀
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet2" "Sheet3!A1") "Sheet3!A1"
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet2" "[B.xlsx]Sheet1!A1") "[B.xlsx]Sheet1!A1"
        Should.equal (RelayDelete.inlineText "Sheet1" "Sheet2" "\"A1\"&A1") "\"A1\"&Sheet1!A1"
        Should.equal (RelayDelete.inlineText "My Sheet" "Sheet2" "A1") "'My Sheet'!A1"

    [<Fact>]
    member this.``关系判定：恰好引用目标``() =
        let rel (formula: string) =
            RelayDelete.referencesOf formula
            |> List.map (RelayDelete.relationTo "Book1.xlsx" "Sheet1" 2 2 "Sheet1")
        Should.equal (rel "=B2*3") [ RelayDelete.IsTarget ]
        Should.equal (rel "=$B$2*3") [ RelayDelete.IsTarget ]
        Should.equal (rel "=Sheet1!B2") [ RelayDelete.IsTarget ]

    [<Fact>]
    member this.``关系判定：区域包含目标``() =
        let rel (formula: string) =
            RelayDelete.referencesOf formula
            |> List.map (RelayDelete.relationTo "Book1.xlsx" "Sheet1" 2 2 "Sheet1")
        Should.equal (rel "=SUM(A1:C3)") [ RelayDelete.Encloses ]
        Should.equal (rel "=SUM(B:B)") [ RelayDelete.Encloses ]
        Should.equal (rel "=SUM(B1:B3)") [ RelayDelete.Encloses ]

    [<Fact>]
    member this.``关系判定：其他表与其他工作簿无关``() =
        let rel (formula: string) =
            RelayDelete.referencesOf formula
            |> List.map (RelayDelete.relationTo "Book1.xlsx" "Sheet1" 2 2 "Sheet1")
        Should.equal (rel "=Sheet2!B2") [ RelayDelete.Unrelated ]
        Should.equal (rel "=[Other.xlsx]Sheet1!B2") [ RelayDelete.Unrelated ]
        Should.equal (rel "=C9") [ RelayDelete.Unrelated ]
        // 写全本工作簿名的自引用仍算本工作簿
        Should.equal (rel "=[Book1.xlsx]Sheet1!B2") [ RelayDelete.IsTarget ]
