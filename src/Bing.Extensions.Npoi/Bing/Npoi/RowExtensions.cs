using System;
using NPOI.SS.UserModel;

namespace Bing.Npoi;

/// <summary>
/// NPOI单元行(<see cref="IRow"/>) 扩展
/// </summary>
public static partial class RowExtensions
{
    /// <summary>
    /// 获取或创建指定位置的单元格。
    /// </summary>
    /// <param name="row">行对象。</param>
    /// <param name="col">列号，从0开始。</param>
    /// <returns>指定位置的单元格。</returns>
    /// <exception cref="ArgumentNullException">如果行对象为null，则抛出此异常。</exception>
    public static ICell GetOrCreateCell(this IRow row, int col)
    {
        if (row is null)
            throw new ArgumentNullException(nameof(row));
        return row.GetCell(col) ?? row.CreateCell(col);
    }

    /// <summary>
    /// 在指定的行中创建单元格，并执行设置操作。
    /// </summary>
    /// <param name="row">要创建单元格的行。</param>
    /// <param name="col">单元格的列号，从0开始。</param>
    /// <param name="setupAct">对新创建的单元格执行的设置操作。</param>
    /// <returns>创建单元格的行。</returns>
    /// <exception cref="ArgumentNullException">如果行对象为null，则抛出此异常。</exception>
    public static IRow CreateCell(this IRow row, int col, Action<ICell> setupAct)
    {
        if (row is null)
            throw new ArgumentNullException(nameof(row));
        var cell = row.GetOrCreateCell(col);
        setupAct?.Invoke(cell);
        return row;
    }

    /// <summary>
    /// 清空内容
    /// </summary>
    /// <param name="row">NPOI单元行</param>
    /// <exception cref="ArgumentNullException"></exception>
    public static IRow ClearContent(this IRow row)
    {
        if (row is null)
            throw new ArgumentNullException(nameof(row));
        foreach (var cell in row.Cells)
            cell.SetCellValue(string.Empty);
        return row;
    }
}