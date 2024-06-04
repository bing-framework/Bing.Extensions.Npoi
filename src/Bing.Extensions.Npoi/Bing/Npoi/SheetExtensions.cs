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
    /// <param name="row">行号，从1开始。</param>
    /// <param name="col">列号，从1开始。</param>
    /// <returns>指定位置的单元格。</returns>
    public static ICell Cell(this ISheet sheet, int row, int col)
    {
        var targetRow = sheet.GetRow(row - 1) ?? sheet.CreateRow(row - 1);
        return targetRow.GetCell(col - 1) ?? targetRow.CreateCell(col - 1);
    }

    /// <summary>
    /// 获取指定位置单元格的值。
    /// </summary>
    /// <param name="sheet">工作表对象。</param>
    /// <param name="row">行号，从1开始。</param>
    /// <param name="col">列号，从1开始。</param>
    /// <returns>指定位置单元格的值，返回类型为动态类型。</returns>
    public static dynamic Get(this ISheet sheet, int row, int col)
    {
        var cell = sheet.Cell(row, col);
        switch (cell.CellType)
        {
            case CellType.Numeric:
                return cell.NumericCellValue;
            case CellType.String:
                return cell.StringCellValue;
            case CellType.Boolean:
                return cell.BooleanCellValue;
            case CellType.Error:
                return cell.ErrorCellValue;
            case CellType.Formula:
                return cell.CellFormula;
            default:
                return "null";
        }
    }

    public static void Set(this ISheet sheet, int row, int col, object value)
    {
        var cell = sheet.Cell(row, col);
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