module ExcelNumericalMethods.NameOps

open System

open FSharp.Idioms
open FSharp.Idioms.StringOps
open FslexFsyacc
open ExcelCompiler

///公式变小括号
let parenthesis (tokens:ExcelToken[]) =
    [|
        yield LPAREN
        yield! Seq.tail tokens
        yield RPAREN
    |]

///假设refersTo中的名称使用全称，不会省略前缀，并且不使用转义字符。
///就是名称和nameName中的名称一样。
///清除名称对名称的引用
let normalizeNames (names:(string*string)[]) =
    // 名称分解
    let firstLasts = 
        names 
        |> Array.map fst
        |> Array.map(fun nm -> NameParser.split nm, nm)

    //根据tok找到名称的名称属性值，没找到返回""
    let findName = function
        | REFERENCE([],[y]) -> 
            firstLasts 
            |> Seq.tryFind(fun ((x0,y0),_) -> x0 = "" && y0 == y)
            |> Option.map snd
            |> Option.defaultValue ""

        | REFERENCE([x],[y]) -> 
            firstLasts 
            |> Seq.tryFind(fun ((x0,y0),_) -> x0 == x && y0 == y)
            |> Option.map snd
            |> Option.defaultValue ""

        | _ -> ""


    // p0: 不依赖名称的名称
    // p1: 仅依赖p0组名称的名称
    // p2: 依赖非p0组名称的名称

    //修改refersto，使其不引用其他名称
    let rec loop (p0:(string*ExcelToken [])[]) (pp:(string*ExcelToken [])[]) =
        if Array.isEmpty pp then p0 else

        //不要使用names，因为refersTo是变化的
        let mpRefersTo = Map.ofArray p0

        //替换公式中的名称
        let replaceName (refersTo:ExcelToken[]) =
            refersTo
            |> Array.collect(fun tok ->
                let key = findName tok
                if mpRefersTo.ContainsKey key then
                    mpRefersTo.[key]
                else
                    [|tok|]
            )

        //执行一次替换
        //分成两部分：无依赖的名称组，有依赖的名称组
        let pp0,ppp =
            pp
            |> Array.map(fun (nm, refersTo) -> nm, replaceName refersTo)
            |> Array.partition(fun (nm, refersTo) ->
                refersTo
                |> Seq.map(findName)
                |> Seq.forall(fun nm -> String.IsNullOrEmpty(nm))
            )

        let pp0 = Array.append p0 pp0
        if Array.isEmpty pp0 then
            failwithf "名称循环引用：%A" (Array.map snd ppp)
        loop pp0 ppp

    let names = 
        names
        |> Array.map(fun(nm, refersTo) -> 
            let tokens =
                refersTo  
                |> ExcelTokenUtils.tokenize 0
                //|> ExcelDFA.analyze
                |> ExcelExprCompiler.analyze
                //|> Seq.concat
                |> Seq.map(fun{value=tok}->tok)
                |> Array.ofSeq
            nm, parenthesis tokens
        )

    loop [||] names

///清除单元格对名称的引用
let replaceNames (names:(string*string)[]) (cells:(string*string*string)[]) =
    // 名称分解
    let firstLasts = 
        names 
        |> Seq.map fst
        |> Seq.map(fun nm -> NameParser.split nm, nm)
        |> Set.ofSeq
        
    let names = 
        names
        |> normalizeNames
        |> Map.ofSeq

    //根据tok找到名称的名称属性值，没找到返回""
    let findName (ws:string) (reference:ExcelToken) =
        match reference with
        | REFERENCE([],[y]) -> 
            // 从后往前找，先匹配工作表名称，后匹配工作簿名称。
            firstLasts
            |> Seq.tryFindBack(fun ((x0,y0),_) -> (Apostrophe.smartDeapos x0 == ws || x0 = "") && y0 == y)
            |> Option.map snd
            |> Option.defaultValue ""

        | REFERENCE([x],[y]) ->
            firstLasts 
            |> Seq.tryFind(fun ((x0,y0),_) -> x0 == x && y0 == y)
            |> Option.map snd
            |> Option.defaultValue ""

        | _ -> ""

    //替换公式中的名称
    let replaceName (ws:string) (formula:#seq<ExcelToken>) =
        formula
        |> Seq.collect(fun tok ->
            match findName ws tok with
            | "" -> Array.singleton tok
            | key -> names.[key]
        )
    

    cells
    |> Array.Parallel.choose(fun (ws,addr,formula) ->
        let tokens = 
            formula  
            //|> ExcelTokenUtils.tokenize
            //|> ExcelDFA.analyze
            //|> Seq.concat
            |> ExcelTokenUtils.tokenize 0
            //|> ExcelDFA.analyze
            |> ExcelExprCompiler.analyze
            //|> Seq.concat
            |> Seq.map(fun{value=tok}->tok)

            |> Array.ofSeq

        //Console.WriteLine(FSharp.Literals.Render.stringify(tokens))

        let noName =
            tokens 
            |> Seq.map(fun tok -> findName ws tok)
            |> Seq.forall(fun nm -> String.IsNullOrEmpty(nm))

        //Console.WriteLine(FSharp.Literals.Render.stringify noName)

        if noName then
            None
        else
            let tokens = 
                replaceName ws tokens

            let formula = 
                tokens
                |> Seq.tail // remove first `=` to expr
                |> Seq.map(fun tok -> {index=0;length=0;value=tok})
                
                |> ExcelExprCompiler.parseTokens

                |> ExprRender.norm
                |> (+) "="

            Some(ws, addr, formula)
    )


