module Program

open ExcelNumericalMethods


let [<EntryPoint>] main _ = 
    let names =
        [|
            "sheet1!x", "=sheet1!A1"
            "sheet2!x", "=sheet2!A1"
            "x","=sheet1!A2"
        |]

    let cells = [|
        "sheet1","B1","=x*2"

        "sheet2","B1","=x*3"

        "sheet3","B1","=x*4"
        "sheet3","B2","=sheet1!x*5"
    |]

    let res = 
        NameOps.replaceNames names cells


    0
