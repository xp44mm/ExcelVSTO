namespace ExcelNumericalMethods.Test
open ExcelNumericalMethods
open ExcelCompiler

open System

open Xunit
open Xunit.Abstractions
open FSharp.xUnit
open FSharp.Idioms

type NameOpsTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``normalize names``() =
        let data =
            [|
                "r", "=sheet1!a1"
                "d", "=r*2"
            |]
        let res = NameOps.normalizeNames data
        Should.equal res [|
            ("r", [|LPAREN; REFERENCE (["sheet1"], ["a1"]); RPAREN|]); 
            ("d", [|LPAREN; LPAREN; REFERENCE (["sheet1"], ["a1"]); RPAREN; MUL; NUMBER "2"; RPAREN|])|]

    [<Fact>]
    member this.``normalize names no action``() =
        let data =
            [|
                "r", "=sheet1!a1"
                "w", "=12"
            |]
        let res = NameOps.normalizeNames data
        let exp = [|
            ("r", [|LPAREN; REFERENCE (["sheet1"], ["a1"]); RPAREN|]); 
            ("w", [|LPAREN; NUMBER "12"; RPAREN|])
        |]
        Should.equal res exp

    [<Fact>]
    member this.``remove names in cells``() =
        let names =
            [|
                "sheet1!x", "=sheet1!A1"
                "sheet2!x", "=sheet2!A1"
                "x"       , "=sheet1!A2"
            |]

        let cells = [|
            "sheet1","B1","=x*2"

            "sheet2","B1","=x*3"

            "sheet3","B1","=x*4"
            "sheet3","B2","=sheet1!x*5"
        |]

        let res = 
            NameOps.replaceNames names cells

        let exp = [|
            "sheet1","B1","=sheet1!A1*2";
            "sheet2","B1","=sheet2!A1*3";
            "sheet3","B1","=sheet1!A2*4";
            "sheet3","B2","=sheet1!A1*5";
            |]

        Should.equal res exp
        //output.WriteLine(Render.stringify res)
