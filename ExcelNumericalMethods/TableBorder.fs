///「竖框线」：清除当前选择区域的全部框线后，为每个单元格添加左右竖框线，
/// 并给最顶行添加上框线、最底行添加下框线。
module ExcelNumericalMethods.TableBorder

open Microsoft.Office.Interop.Excel

/// 把边框设为细实线（xlContinuous / xlThin）。
let private thin (b: Border) =
    b.LineStyle <- XlLineStyle.xlContinuous
    b.Weight <- XlBorderWeight.xlThin

/// 清除一个连续区域的全部框线（含对角线）。
let private clearBorders (area: Range) =
    let borders = area.Borders
    borders.LineStyle <- XlLineStyle.xlLineStyleNone
    borders.[XlBordersIndex.xlDiagonalDown].LineStyle <- XlLineStyle.xlLineStyleNone
    borders.[XlBordersIndex.xlDiagonalUp].LineStyle <- XlLineStyle.xlLineStyleNone

/// 添加每个单元格的左右竖框线：内部列界线加竖线，区域最左列加左边框、最右列加右边框。
let private addVerticalBorders (area: Range) =
    let borders = area.Borders
    thin borders.[XlBordersIndex.xlInsideVertical]
    thin borders.[XlBordersIndex.xlEdgeLeft]
    thin borders.[XlBordersIndex.xlEdgeRight]

/// 添加最顶上的上框线。
let private addTopBorder (area: Range) =
    thin area.Borders.[XlBordersIndex.xlEdgeTop]

/// 添加最底下的下框线。
let private addBottomBorder (area: Range) =
    thin area.Borders.[XlBordersIndex.xlEdgeBottom]

/// 对当前选择的每个连续子区域依次执行：
/// 1. 清除全部框线；
/// 2. 添加每个单元格的左右竖框线；
/// 3. 添加最顶行的上框线；
/// 4. 添加最底行的下框线。
let apply (sel: Range) =
    for a in 1 .. sel.Areas.Count do
        let area = sel.Areas.[a]
        clearBorders area
        addVerticalBorders area
        addTopBorder area
        addBottomBorder area
