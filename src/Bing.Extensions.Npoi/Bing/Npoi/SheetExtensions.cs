using System;
using NPOI.SS.UserModel;

namespace Bing.Npoi;

/// <summary>
/// 工作表(<see cref="ISheet"/>) 扩展
/// </summary>
public static class SheetExtensions
{
    /// <summary>
    /// 获取或创建指定位置的单元格。
    /// </summary>
    /// <param name="sheet">工作表对象。</param>
    /// <param name="row">行号，从0开始。</param>
    /// <param name="col">列号，从0开始。</param>
    /// <returns>指定位置的单元格。</returns>
    public static ICell GetOrCreateCell(this ISheet sheet, int row, int col)
    {
        if (sheet == null)
            throw new ArgumentNullException(nameof(sheet));
        var targetRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
        return targetRow.GetOrCreateCell(col);
    }

    /// <summary>
    /// 获取指定位置单元格的字符串值。
    /// </summary>
    /// <param name="sheet">要获取值的工作表。</param>
    /// <param name="row">单元格所在的行号，从0开始。</param>
    /// <param name="col">单元格所在的列号，从0开始。</param>
    /// <returns>指定位置单元格的字符串值。如果单元格不存在，将创建一个新的单元格并返回其值。</returns>
    /// <exception cref="ArgumentNullException">如果工作表为null，则抛出此异常。</exception>
    public static string GetStringValue(this ISheet sheet, int row, int col)
    {
        if (sheet == null)
            throw new ArgumentNullException(nameof(sheet));
        var cell = sheet.GetOrCreateCell(row, col);
        return cell.GetStringValue();
    }

    /// <summary>
    /// 获取指定位置单元格的值。
    /// </summary>
    /// <param name="sheet">要获取值的工作表。</param>
    /// <param name="row">单元格所在的行号，从0开始。</param>
    /// <param name="col">单元格所在的列号，从0开始。</param>
    /// <returns>指定位置单元格的值。返回值类型根据单元格的类型而定。如果单元格不存在，将创建一个新的单元格并返回其值。</returns>
    /// <exception cref="ArgumentNullException">如果工作表为null，则抛出此异常。</exception>
    public static dynamic Get(this ISheet sheet, int row, int col)
    {
        if (sheet == null)
            throw new ArgumentNullException(nameof(sheet));
        var cell = sheet.GetOrCreateCell(row, col);
        return cell.CellType switch
        {
            CellType.Numeric => cell.NumericCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Error => cell.ErrorCellValue,
            CellType.Formula => cell.CellFormula,
            _ => "null"
        };
    }

    /// <summary>
    /// 获取公式单元格的计算结果值。
    /// </summary>
    /// <param name="cell">要获取值的单元格。</param>
    /// <returns>公式单元格的计算结果值。返回值类型根据公式的计算结果类型而定。如果单元格不是公式类型，或者公式尚未计算，返回空字符串。</returns>
    public static void Set(this ISheet sheet, int row, int col, object value)
    {
        var cell = sheet.GetOrCreateCell(row, col);
        switch (value)
        {
            case null:
                cell.SetCellValue("");
                break;
            case string s:
                cell.SetCellValue(s);
                break;
            case double:
            case int:
            case float:
            case long:
            case decimal:
                cell.SetCellValue(Convert.ToDouble(value));
                break;
            case DateTime time:
                cell.SetCellValue(time);
                break;
            case bool b:
                cell.SetCellValue(b);
                break;
            default:
                cell.SetCellValue(value.ToString());
                break;
        }
    }
}