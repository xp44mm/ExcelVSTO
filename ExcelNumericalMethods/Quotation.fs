module Quotation

open System

///加双引号:一个引号变两个引号
let quote (s:string) = "\""+ s.Replace("\"","\"\"") + "\""

///去双引号
let dequote(s:string) = s.[1..s.Length - 2].Replace("\"\"","\"")

///如果没有双引号，则加双引号
let smartQuote (s:string) =
    if System.String.IsNullOrEmpty s then 
        s
    elif s.Chars 0 = '"' then 
        quote s
    else s