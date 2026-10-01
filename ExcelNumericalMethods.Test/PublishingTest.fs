namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open System
open Xunit
open FSharp.xUnit

type PublishingTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``splitIfError 基本结构``() =
        Should.equal (Publishing.splitIfError "=IFERROR(电机额定功率(D11), 22)") (Some("电机额定功率(D11)", "22"))

    [<Fact>]
    member this.``splitIfError 嵌套函数括号与逗号不误切``() =
        Should.equal (Publishing.splitIfError "=IFERROR(电机额定功率(SUM(1,2),3), 22)") (Some("电机额定功率(SUM(1,2),3)", "22"))

    [<Fact>]
    member this.``splitIfError 字符串字面量内逗号不误切``() =
        Should.equal (Publishing.splitIfError "=IFERROR(电机额定功率(\"a,b\"), 22)") (Some("电机额定功率(\"a,b\")", "22"))

    [<Fact>]
    member this.``splitIfError 字符串字面量内转义双引号``() =
        Should.equal (Publishing.splitIfError "=IFERROR(电机额定功率(\"a\"\"b\"), 0)") (Some("电机额定功率(\"a\"\"b\")", "0"))

    [<Fact>]
    member this.``splitIfError 小写 iferror``() =
        Should.equal (Publishing.splitIfError "=iferror(A1, 0)") (Some("A1", "0"))

    [<Fact>]
    member this.``splitIfError 兜底值为引用``() =
        Should.equal (Publishing.splitIfError "=IFERROR(电机额定功率(D11), E11)") (Some("电机额定功率(D11)", "E11"))

    [<Fact>]
    member this.``splitIfError 非 IFERROR 结构返回 None``() =
        Should.equal (Publishing.splitIfError "=SUM(A1, B1)") None

    [<Fact>]
    member this.``splitIfError 无顶层逗号返回 None``() =
        Should.equal (Publishing.splitIfError "=IFERROR(A1)") None

    [<Fact>]
    member this.``splitIfError 不带等号返回 None``() =
        Should.equal (Publishing.splitIfError "IFERROR(A1, 2)") None

    [<Fact>]
    member this.``splitIfError 空参数返回 None``() =
        Should.equal (Publishing.splitIfError "=IFERROR(, 2)") None

    [<Fact>]
    member this.``formatLiteral 文本加双引号并转义内部双引号``() =
        Should.equal (Publishing.formatLiteral (box "22")) "\"22\""
        Should.equal (Publishing.formatLiteral (box "说\"好\"")) "\"说\"\"好\"\"\""

    [<Fact>]
    member this.``formatLiteral 数字``() =
        Should.equal (Publishing.formatLiteral (box 22.0)) "22"
        Should.equal (Publishing.formatLiteral (box 1.5)) "1.5"
        Should.equal (Publishing.formatLiteral (box 123.456)) "123.456"

    [<Fact>]
    member this.``formatLiteral 整数类型``() =
        Should.equal (Publishing.formatLiteral (box 22)) "22"

    [<Fact>]
    member this.``formatLiteral 布尔``() =
        Should.equal (Publishing.formatLiteral (box true)) "TRUE"
        Should.equal (Publishing.formatLiteral (box false)) "FALSE"

    [<Fact>]
    member this.``formatLiteral 空值``() =
        Should.equal (Publishing.formatLiteral null) "\"\""

    [<Fact>]
    member this.``formatLiteral 日期``() =
        Should.equal (Publishing.formatLiteral (box (DateTime(2026, 10, 1)))) "DATE(2026,10,1)"

    [<Fact>]
    member this.``columnLetters 列号转字母``() =
        Should.equal (Publishing.columnLetters 1) "A"
        Should.equal (Publishing.columnLetters 26) "Z"
        Should.equal (Publishing.columnLetters 27) "AA"
        Should.equal (Publishing.columnLetters 52) "AZ"
        Should.equal (Publishing.columnLetters 53) "BA"
        Should.equal (Publishing.columnLetters 703) "AAA"
