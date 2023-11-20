namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open Xunit.Abstractions
open FSharp.xUnit

type WorksheetOpsTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``active worksheet references``() =
        let sheet2 =
            [|
                "a1", "=3"
                "a2", "=a1*2"
                "a3", "=sheet1!a1+a1"
                "a4", "=sheet2!a2*4"
            |]
        let res = 
            WorksheetOps.references "sheet2" sheet2
            |> List.ofSeq
            |> List.map fst
        Should.equal res ["a3"]


    [<Fact>]
    member this.``active worksheet dependents``() =
        let sheet1 = 
            [|
                "sheet1","a1", "=3"
                "sheet1","a2", "=a1*2"
            |]

        let otherWss =
            [|
                "sheet2","a1", "=3"
                "sheet2","a2", "=a1*2"
                "sheet2","a3", "=sheet1!a1*3"
                "sheet2","a4", "=sheet2!a1*4"
                "sheet2","a5", "=sheet2!a2*5"

                "sheet3","a1", "=3"
                "sheet3","a2", "=a1*2"
                "sheet3","a3", "=sheet1!a1*6"
                "sheet3","a4", "=sheet2!a1*7"
            |]

        let res = 
            WorksheetOps.dependents "sheet1" otherWss
            |> List.ofSeq
            |> List.map (fun(ws,addr,formula)-> ws,addr)
        Should.equal res ["sheet2","a3";"sheet3","a3"]

