module ExcelNumericalMethods.Prune

open System.Text.RegularExpressions

/// 比较列地址的大小
let compareAlphabet (a:string) (b:string) =
    match compare a.Length b.Length with
    | 0 -> compare a b
    | x -> x

let minCol (a:string) (b:string) =
    match compareAlphabet a b with
    | x when x > 0 -> b
    | _ -> a
    
let cellRgx = Regex(@"^\$(\w+)\$(\d+)$")
let parseCellAddress (cell:string) =
    let gs = cellRgx.Match(cell)
    let aa = gs.Groups.[1].Value.ToUpper()
    let nn = int gs.Groups.[2].Value
    aa, nn

let colRgx = Regex(@"^\$([a-zA-Z]+)$")
let parseColAddress (col:string) =
    let gs = colRgx.Match(col)
    let aa = gs.Groups.[1].Value
    aa

let rowRgx = Regex(@"^\$(\d+)$")
let parseRowAddress (row:string) =
    let gs = rowRgx.Match(row)
    let nn = int gs.Groups.[1].Value
    nn

/// used range 的右下角
let usedRangeBottomRight (addr:string) =
    let xs = addr.Split(':')
    let br = xs.[xs.Length-1]
    parseCellAddress br

let prune (usedRange:string) (selectedRange:string) =
    if selectedRange.Contains ":" then
        let c0,r0 = usedRangeBottomRight usedRange

        let tl,br = 
            let ls = selectedRange.Split(':')
            ls.[0],ls.[1]

        if rowRgx.IsMatch(tl) then
            let r1 = parseRowAddress br
            let cc = c0
            let rr = min r0 r1
            sprintf "$A%s:$%s$%d" tl cc rr

        elif colRgx.IsMatch(tl) then
            let c1 = parseColAddress br
            let cc = minCol c0 c1
            let rr = r0
            sprintf "%s$1:$%s$%d" tl cc rr

        elif cellRgx.IsMatch(tl) then
            let c1,r1 = parseCellAddress br
            let cc = minCol c0 c1
            let rr = min r0 r1
            sprintf "%s:$%s$%d" tl cc rr

        else
            failwithf "prune:%s,%s" usedRange selectedRange
    else
        selectedRange