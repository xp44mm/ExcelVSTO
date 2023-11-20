namespace ExcelNumericalMethods.Test

open ExcelNumericalMethods
open Xunit
open Xunit.Abstractions
open FSharp.xUnit

type PruneTest(output : ITestOutputHelper) =

    [<Fact>]
    member this.``usedBottomRight 1 Test``() =
        let data = "$A$1:$B$4"
        let aa,nn = Prune.usedRangeBottomRight data

        Should.equal aa "B"
        Should.equal nn 4

    [<Fact>]
    member this.``usedBottomRight 2 Test``() =
        let data = "$B$4"
        let aa,nn = Prune.usedRangeBottomRight data

        Should.equal aa "B"
        Should.equal nn 4

    [<Fact>]
    member this.``compareAlphabet Test``() =
        let x = Prune.compareAlphabet "B" "AA"
        Should.equal x -1

        let x = Prune.compareAlphabet "B" "B"
        Should.equal x 0

        let x = Prune.compareAlphabet "A" "B"
        Should.equal x -1

        let x = Prune.compareAlphabet "B" "A"
        Should.equal x 1

    [<Fact>]
    member this.``min col Test``() =
        let x = Prune.minCol "B" "AA"
        Should.equal x "B"

        let x = Prune.minCol "B" "B"
        Should.equal x "B"

        let x = Prune.minCol "A" "B"
        Should.equal x "A"

        let x = Prune.minCol "B" "A"
        Should.equal x "A"

    [<Fact>]
    member this.``parseCellAddress Test``() =
        let aa,nn = Prune.parseCellAddress "$B$1"
        Should.equal aa "B"
        Should.equal nn 1

    [<Fact>]
    member this.``parseColAddress Test``() =
        let aa = Prune.parseColAddress "$B"
        Should.equal aa "B"

    [<Fact>]
    member this.``parseRowAddress Test``() =
        let nn = Prune.parseRowAddress "$1"
        Should.equal nn 1

    [<Fact>]
    member this.``prune Test``() =
        let usedRanges = [
            "$A$1:$F$8"
            "$F$8"
            "$C$3:$F$8"
        ]

        let selectedRanges = [
            "$1:$999"
            "$2:$4"
            "$A:$G"
            "$C:$E"
            "$A$1:$G$9"
            "$D$4:$E$7"
            "$G$9"
            "$D$4"
        ]
        let inputs = 
            usedRanges
            |> List.collect(fun u -> selectedRanges |> List.map(fun s -> u,s))

        let expected = [
            ("$A$1:$F$8", "$1:$999", "$A$1:$F$8")
            ("$A$1:$F$8", "$2:$4", "$A$2:$F$4")
            ("$A$1:$F$8", "$A:$G", "$A$1:$F$8")
            ("$A$1:$F$8", "$C:$E", "$C$1:$E$8")
            ("$A$1:$F$8", "$A$1:$G$9", "$A$1:$F$8")
            ("$A$1:$F$8", "$D$4:$E$7", "$D$4:$E$7")
            ("$A$1:$F$8", "$G$9", "$G$9")
            ("$A$1:$F$8", "$D$4", "$D$4")
            ("$F$8", "$1:$999", "$A$1:$F$8")
            ("$F$8", "$2:$4", "$A$2:$F$4")
            ("$F$8", "$A:$G", "$A$1:$F$8")
            ("$F$8", "$C:$E", "$C$1:$E$8")
            ("$F$8", "$A$1:$G$9", "$A$1:$F$8")
            ("$F$8", "$D$4:$E$7", "$D$4:$E$7")
            ("$F$8", "$G$9", "$G$9")
            ("$F$8", "$D$4", "$D$4")
            ("$C$3:$F$8", "$1:$999", "$A$1:$F$8")
            ("$C$3:$F$8", "$2:$4", "$A$2:$F$4")
            ("$C$3:$F$8", "$A:$G", "$A$1:$F$8")
            ("$C$3:$F$8", "$C:$E", "$C$1:$E$8")
            ("$C$3:$F$8", "$A$1:$G$9", "$A$1:$F$8")
            ("$C$3:$F$8", "$D$4:$E$7", "$D$4:$E$7")
            ("$C$3:$F$8", "$G$9", "$G$9")
            ("$C$3:$F$8", "$D$4", "$D$4")
        ]
        let actual = [
            for ur,sr in inputs do
                let res = Prune.prune ur sr
                yield (ur,sr,res)
        ]

        Should.equal expected actual