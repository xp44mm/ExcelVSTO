module ExcelNumericalMethods.Apostrophe

///加双引号:一个引号变两个引号
let quote (s:string) = "\""+ s.Replace("\"","\"\"") + "\""

///去单引号
let deapos(s:string) = s.[1..s.Length-2].Replace("''","'")

///从公式中获取工作表的名称
let smartDeapos (ws:string) =
    if System.String.IsNullOrEmpty ws then 
        ws
    elif ws.Chars 0 = '\'' then 
        deapos ws 
    else ws

