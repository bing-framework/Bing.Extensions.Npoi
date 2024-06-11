using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Bing.Npoi;

/// <summary>
/// 工作表(<see cref="ISheet"/>) 扩展
/// </summary>
public static partial class SheetExtensions
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
    /// 在指定的行索引位置插入一行新行。
    /// </summary>
    /// <param name="sheet">要插入新行的工作表。</param>
    /// <param name="rowIndex">新行的行索引。</param>
    /// <returns>插入的新行。</returns>
    /// <exception cref="ArgumentNullException">如果工作表为null，则抛出此异常。</exception>
    public static IRow InsertRow(this ISheet sheet, int rowIndex) => sheet.InsertRows(rowIndex, 1)[0];

    /// <summary>
    /// 在指定的行索引位置插入指定数量的新行。
    /// </summary>
    /// <param name="sheet">要插入新行的工作表。</param>
    /// <param name="rowIndex">新行的起始行索引。</param>
    /// <param name="rowsCount">要插入的新行的数量。</param>
    /// <returns>插入的新行数组。</returns>
    /// <exception cref="ArgumentNullException">如果工作表为null，则抛出此异常。</exception>
    public static IRow[] InsertRows(this ISheet sheet, int rowIndex, int rowsCount)
    {
        if (sheet == null)
            throw new ArgumentNullException(nameof(sheet));
        if (rowIndex <= sheet.LastRowNum)
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowsCount, true, false);
        var rows = new List<IRow>();
        for (var i = 0; i < rowsCount; i++)
        {
            var row = sheet.CreateRow(rowIndex + i);
            rows.Add(row);
        }
        return rows.ToArray();
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
    /// 检查目标范围是否在指定范围内部或与其相交。
    /// </summary>
    /// <param name="rangeMinRow">指定范围的最小行索引。</param>
    /// <param name="rangeMaxRow">指定范围的最大行索引。</param>
    /// <param name="rangeMinCol">指定范围的最小列索引。</param>
    /// <param name="rangeMaxCol">指定范围的最大列索引。</param>
    /// <param name="targetMinRow">目标范围的最小行索引。</param>
    /// <param name="targetMaxRow">目标范围的最大行索引。</param>
    /// <param name="targetMinCol">目标范围的最小列索引。</param>
    /// <param name="targetMaxCol">目标范围的最大列索引。</param>
    /// <param name="onlyInternal">如果为true，只检查目标范围是否在指定范围内部；如果为false，检查目标范围是否在指定范围内部或与其相交。</param>
    /// <returns>如果目标范围在指定范围内部或与其相交，返回true；否则返回false。</returns>
    private static bool IsInternalOrIntersect(
        int? rangeMinRow, int? rangeMaxRow, 
        int? rangeMinCol, int? rangeMaxCol, 
        int targetMinRow, int targetMaxRow, 
        int targetMinCol, int targetMaxCol, 
        bool onlyInternal)
    {
        var tempMinRow = rangeMinRow ?? targetMinRow;
        var tempMaxRow = rangeMaxRow ?? targetMaxRow;
        var tempMinCol = rangeMinCol ?? targetMinCol;
        var tempMaxCol = rangeMaxCol ?? targetMaxCol;
        if (onlyInternal)
        {
            return tempMinRow <= targetMinRow &&
                   tempMaxRow >= targetMaxRow &&
                   tempMinCol <= targetMinCol &&
                   tempMaxCol >= targetMaxCol;
        }

        return Math.Abs(tempMaxRow - tempMinRow) + Math.Abs(targetMaxRow - targetMinRow) >=
               Math.Abs(tempMaxRow + tempMinRow - targetMaxRow - targetMinRow) &&
               Math.Abs(tempMaxCol - tempMinCol) + Math.Abs(targetMaxCol - targetMinCol) >=
               Math.Abs(tempMaxCol + tempMinCol - targetMaxCol - targetMinCol);
    }
}