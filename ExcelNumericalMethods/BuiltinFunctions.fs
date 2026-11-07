namespace ExcelNumericalMethods

open System
open System.Collections.Generic

/// Excel 内置工作表函数名清单，用于把「自定义函数」（Excel-DNA / XLL / VBA 宏表函数）与内置函数区分开。
///
/// 用途：脱公式、包裹自定义函数等批量操作只应动自定义函数，内置函数（SUM、IF、SQRT……）必须原样保留。
/// 判定口径：公式中出现 `_xll.` 前缀（Excel 保存 XLL 函数时的写法，如 `_xll.电机额定功率(D11)`），
/// 或者调用了一个不在本清单内的名字，都算自定义函数。
///
/// 本文件被两个解决方案共用：ExcelVSTO 下的 ExcelNumericalMethods 项目直接编译，
/// FormulaStripper 独立解决方案通过 <Compile Include> 链接同一份源文件，避免维护两份清单。
module BuiltinFunctions =

    /// Excel 内置工作表函数名（大写，含带点与带下划线的现代函数名）
    let private builtinNames: string[] =
        [| "ABS"; "ACCRINT"; "ACCRINTM"; "ACOS"; "ACOSH"; "ACOT"; "ACOTH"; "ADDRESS"; "AGGREGATE"
           "AMORDEGRC"; "AMORLINC"; "AND"; "ARABIC"; "AREAS"; "ARRAYTOTEXT"; "ASC"; "ASIN"; "ASINH"
           "ATAN"; "ATAN2"; "ATANH"; "AVEDEV"; "AVERAGE"; "AVERAGEA"; "AVERAGEIF"; "AVERAGEIFS"
           "BAHTTEXT"; "BASE"; "BESSELI"; "BESSELJ"; "BESSELK"; "BESSELY"
           "BETA.DIST"; "BETA.INV"; "BETADIST"; "BETAINV"
           "BIN2DEC"; "BIN2HEX"; "BIN2OCT"; "BINOM.DIST"; "BINOM.DIST.RANGE"; "BINOM.INV"; "BINOMDIST"
           "BITAND"; "BITLSHIFT"; "BITOR"; "BITRSHIFT"; "BITXOR"; "BYCOL"; "BYROW"
           "CALL"; "CEILING"; "CEILING.MATH"; "CEILING.PRECISE"; "CELL"; "CHAR"
           "CHIDIST"; "CHIINV"; "CHISQ.DIST"; "CHISQ.DIST.RT"; "CHISQ.INV"; "CHISQ.INV.RT"; "CHISQ.TEST"; "CHITEST"
           "CHOOSE"; "CHOOSECOLS"; "CHOOSEROWS"; "CLEAN"; "CODE"; "COLUMN"; "COLUMNS"
           "COMBIN"; "COMBINA"; "COMPLEX"; "CONCAT"; "CONCATENATE"
           "CONFIDENCE"; "CONFIDENCE.NORM"; "CONFIDENCE.T"; "CONVERT"; "CORREL"; "COS"; "COSH"; "COT"; "COTH"
           "COUNT"; "COUNTA"; "COUNTBLANK"; "COUNTIF"; "COUNTIFS"
           "COUPDAYBS"; "COUPDAYS"; "COUPDAYSNC"; "COUPNCD"; "COUPPCD"
           "COVAR"; "COVARIANCE.P"; "COVARIANCE.S"; "CRITBINOM"; "CSC"; "CSCH"
           "CUBEKPIMEMBER"; "CUBEMEMBER"; "CUBEMEMBERPROPERTY"; "CUBERANKEDMEMBER"
           "CUBESET"; "CUBESETCOUNT"; "CUBEVALUE"; "CUMIPMT"; "CUMPRINC"
           "DATE"; "DATEDIF"; "DATEVALUE"; "DAVERAGE"; "DAY"; "DAYS"; "DAYS360"; "DB"; "DBCS"
           "DCOUNT"; "DCOUNTA"; "DDB"; "DEC2BIN"; "DEC2HEX"; "DEC2OCT"; "DECIMAL"; "DEGREES"; "DELTA"
           "DEVSQ"; "DGET"; "DISC"; "DMAX"; "DMIN"; "DOLLAR"; "DOLLARDE"; "DOLLARFR"
           "DPRODUCT"; "DSTDEV"; "DSTDEVP"; "DSUM"; "DURATION"; "DVAR"; "DVARP"
           "EDATE"; "EFFECT"; "ENCODEURL"; "EOMONTH"; "ERF"; "ERF.PRECISE"; "ERFC"; "ERFC.PRECISE"
           "ERROR.TYPE"; "EUROCONVERT"; "EVEN"; "EXACT"; "EXP"; "EXPON.DIST"; "EXPONDIST"
           "F.DIST"; "F.DIST.RT"; "F.INV"; "F.INV.RT"; "F.TEST"; "FACT"; "FACTDOUBLE"; "FALSE"; "FDIST"
           "FILTER"; "FILTERXML"; "FIND"; "FINDB"; "FINV"; "FISHER"; "FISHERINV"; "FIXED"
           "FLOOR"; "FLOOR.MATH"; "FLOOR.PRECISE"; "FORMULATEXT"; "FREQUENCY"; "FTEST"; "FV"; "FVSCHEDULE"
           "GAMMA"; "GAMMA.DIST"; "GAMMA.INV"; "GAMMADIST"; "GAMMAINV"; "GAMMALN"; "GAMMALN.PRECISE"
           "GAUSS"; "GCD"; "GEOMEAN"; "GETPIVOTDATA"; "GROWTH"
           "HARMEAN"; "HEX2BIN"; "HEX2DEC"; "HEX2OCT"; "HLOOKUP"; "HOUR"; "HSTACK"; "HYPERLINK"
           "HYPGEOM.DIST"; "HYPGEOMDIST"
           "IF"; "IFERROR"; "IFNA"; "IFS"
           "IMABS"; "IMAGINARY"; "IMARGUMENT"; "IMCONJUGATE"; "IMCOS"; "IMCOSH"; "IMCOT"; "IMCSC"; "IMCSCH"
           "IMDIV"; "IMEXP"; "IMLN"; "IMLOG10"; "IMLOG2"; "IMPOWER"; "IMPRODUCT"; "IMREAL"; "IMSEC"; "IMSECH"
           "IMSIN"; "IMSINH"; "IMSQRT"; "IMSUB"; "IMSUM"; "IMTAN"
           "INDEX"; "INDIRECT"; "INFO"; "INT"; "INTERCEPT"; "INTRATE"; "IPMT"; "IRR"
           "ISBLANK"; "ISERR"; "ISERROR"; "ISEVEN"; "ISFORMULA"; "ISLOGICAL"; "ISNA"; "ISNONTEXT"
           "ISNUMBER"; "ISODD"; "ISOMITTED"; "ISPMT"; "ISREF"; "ISTEXT"
           "ISO.CEILING"; "ISOWEEKNUM"; "JIS"
           "KURT"; "LAMBDA"; "LARGE"; "LCM"; "LEFT"; "LEFTB"; "LEN"; "LENB"; "LET"; "LINEST"
           "LN"; "LOG"; "LOG10"; "LOGEST"; "LOGINV"; "LOGNORM.DIST"; "LOGNORM.INV"; "LOGNORMDIST"
           "LOOKUP"; "LOWER"
           "MAKEARRAY"; "MAP"; "MDETERM"; "MDURATION"; "MEDIAN"; "MID"; "MIDB"
           "MIN"; "MINA"; "MINIFS"; "MINUTE"; "MINVERSE"; "MIRR"; "MMULT"; "MOD"
           "MODE"; "MODE.MULT"; "MODE.SNGL"; "MONTH"; "MROUND"; "MULTINOMIAL"; "MUNIT"
           "N"; "NA"; "NEGBINOM.DIST"; "NEGBINOMDIST"; "NETWORKDAYS"; "NETWORKDAYS.INTL"
           "NOMINAL"; "NORM.DIST"; "NORM.INV"; "NORM.S.DIST"; "NORM.S.INV"
           "NORMDIST"; "NORMINV"; "NORMSDIST"; "NORMSINV"; "NOT"; "NOW"; "NPER"; "NPV"; "NUMBERVALUE"
           "OCT2BIN"; "OCT2DEC"; "OCT2HEX"; "ODD"; "ODDFPRICE"; "ODDFYIELD"; "ODDLPRICE"; "ODDLYIELD"
           "OFFSET"; "OR"
           "PDURATION"; "PEARSON"; "PERCENTILE"; "PERCENTILE.EXC"; "PERCENTILE.INC"
           "PERCENTRANK"; "PERCENTRANK.EXC"; "PERCENTRANK.INC"; "PERMUT"; "PERMUTATIONA"
           "PHI"; "PHONETIC"; "PI"; "PMT"; "POISSON"; "POISSON.DIST"; "POWER"; "PPMT"
           "PRICE"; "PRICEDISC"; "PRICEMAT"; "PROB"; "PRODUCT"; "PROPER"; "PV"
           "QUARTILE"; "QUARTILE.EXC"; "QUARTILE.INC"; "QUOTIENT"
           "RADIANS"; "RAND"; "RANDARRAY"; "RANDBETWEEN"; "RANK"; "RANK.AVG"; "RANK.EQ"; "RATE"
           "RECEIVED"; "REDUCE"; "REGISTER.ID"; "REPLACE"; "REPLACEB"; "REPT"; "RIGHT"; "RIGHTB"
           "ROMAN"; "ROUND"; "ROUNDDOWN"; "ROUNDUP"; "ROW"; "ROWS"; "RRI"; "RSQ"; "RTD"
           "SCAN"; "SEARCH"; "SEARCHB"; "SEC"; "SECH"; "SECOND"; "SEQUENCE"; "SERIESSUM"
           "SHEET"; "SHEETS"; "SIGN"; "SIN"; "SINH"; "SKEW"; "SKEW.P"; "SLN"; "SLOPE"; "SMALL"
           "SORT"; "SORTBY"; "SQRT"; "SQRTPI"; "STANDARDIZE"
           "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "STEYX"
           "SUBSTITUTE"; "SUBTOTAL"; "SUM"; "SUMIF"; "SUMIFS"; "SUMPRODUCT"; "SUMSQ"
           "SUMX2MY2"; "SUMX2PY2"; "SUMXMY2"; "SWITCH"; "SYD"
           "T"; "T.DIST"; "T.DIST.2T"; "T.DIST.RT"; "T.INV"; "T.INV.2T"; "T.TEST"
           "TAN"; "TANH"; "TBILLEQ"; "TBILLPRICE"; "TBILLYIELD"; "TDIST"
           "TEXT"; "TEXTAFTER"; "TEXTBEFORE"; "TEXTJOIN"; "TEXTSPLIT"
           "TIME"; "TIMEVALUE"; "TINV"; "TODAY"; "TOCOL"; "TOROW"; "TRANSPOSE"; "TREND"; "TRIM"; "TRIMMEAN"
           "TRUE"; "TRUNC"; "TTEST"; "TYPE"; "UNICHAR"; "UNICODE"; "UNIQUE"; "UPPER"
           "VALUE"; "VALUETOTEXT"; "VALUETYPE"; "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VDB"
           "VLOOKUP"; "VSTACK"; "WEBSERVICE"; "WEEKDAY"; "WEEKNUM"; "WEIBULL"; "WEIBULL.DIST"
           "WORKDAY"; "WORKDAY.INTL"; "WRAPCOLS"; "WRAPROWS"
           "XIRR"; "XLOOKUP"; "XMATCH"; "XNPV"; "XOR"; "YEAR"; "YEARFRAC"
           "YIELD"; "YIELDDISC"; "YIELDMAT"; "Z.TEST"; "ZTEST" |]

    let private lookup: HashSet<string> =
        let set = HashSet<string>(StringComparer.OrdinalIgnoreCase)
        for name in builtinNames do set.Add name |> ignore
        set

    /// 名称是否为 Excel 内置工作表函数（不区分大小写）。
    let isBuiltin (name: string) : bool =
        not (String.IsNullOrEmpty name) && lookup.Contains name

    /// 是否为 XLL 注册函数的存储写法（形如 `_xll.函数名`）；`_xll.` 前缀本身也算。
    let isXllName (name: string) : bool =
        not (String.IsNullOrEmpty name)
        && name.StartsWith("_xll", StringComparison.OrdinalIgnoreCase)

    /// 名称是否为自定义函数：既非内置，也不是 XLL 前缀形式（此时按普通自定义名处理）。
    let isCustomName (name: string) : bool =
        not (String.IsNullOrEmpty name)
        && (isXllName name || not (isBuiltin name))

    /// 去掉 XLL 前缀，返回裸函数名：`_xll.电机额定功率` -> `电机额定功率`。
    let stripXllPrefix (name: string) : string =
        if isNull name then name
        elif name.StartsWith("_xll.", StringComparison.OrdinalIgnoreCase) then name.Substring("_xll.".Length)
        else name

    /// 公式中是否含有自定义函数（Excel-DNA / XLL / VBA 宏表函数等）。
    ///
    /// 判定口径（先跳过字符串字面量，避免文本里的函数名造成误判）：
    ///   1. 出现 `_xll` 前缀——Excel 保存 XLL 注册函数时的写法，如 `_xll.电机额定功率(D11)`；
    ///   2. 或者调用了一个不在内置函数名清单内的名字。
    /// 内置函数（SUM、IF、SQRT……）一律不算，批量操作不应动它们。
    let hasCustomFunction (formula: string) : bool =
        if String.IsNullOrEmpty formula then
            false
        else
            let n = formula.Length
            let mutable i = 0
            let mutable found = false
            let mutable inString = false
            while i < n && not found do
                let c = formula.[i]
                if inString then
                    if c = '"' then
                        // 字符串字面量内的 "" 是转义的双引号
                        if i + 1 < n && formula.[i + 1] = '"' then i <- i + 1
                        else inString <- false
                elif c = '"' then
                    inString <- true
                elif i + 4 <= n && String.Equals(formula.Substring(i, 4), "_xll", StringComparison.OrdinalIgnoreCase) then
                    found <- true
                elif Char.IsLetter c || c = '_' then
                    // 收集标识符（字母、数字、下划线、点），若后随左括号则视为函数调用
                    let start = i
                    let mutable j = i
                    while j < n && (Char.IsLetterOrDigit formula.[j] || formula.[j] = '_' || formula.[j] = '.') do
                        j <- j + 1
                    let name = formula.Substring(start, j - start)
                    // 容错：函数名与左括号之间可能有空格
                    let p = if j < n && formula.[j] = ' ' then j + 1 else j
                    if p < n && formula.[p] = '(' then
                        if not (isBuiltin (stripXllPrefix name)) then found <- true
                    i <- j - 1
                i <- i + 1
            found
