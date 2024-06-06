using System.Globalization;
using System;
using NPOI.SS.UserModel;

namespace Bing.Npoi;

/// <summary>
/// NPOI单元格(<see cref="ICell"/>) 扩展
/// </summary>
public static partial class CellExtensions
{
    /// <summary>
    /// 获取单元格的字符串值。
    /// </summary>
    /// <param name="cell">要获取值的单元格。</param>
    /// <returns>单元格的字符串值。如果单元格为null，返回空字符串。</returns>
    public static string GetStringValue(this ICell cell)
    {
        var result = string.Empty;
        if (cell == null)
            return result;
        try
        {
            result = cell.CellType switch
            {
                CellType.String => cell.StringCellValue,
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Error => cell.ErrorCellValue.ToString(),
                CellType.Formula => cell.CellFormula,
                CellType.Numeric => GetNumericCellValue(cell),
                _ => cell.ToString(),
            };
            return result?.Trim();
        }
        catch
        {
            return result;
        }
    }

    /// <summary>
    /// 获取数值类型单元格的值。如果单元格是日期格式，则返回日期字符串，否则返回数值字符串。
    /// </summary>
    /// <param name="cell">要获取值的单元格。</param>
    /// <returns>单元格的值。如果单元格是日期格式，则返回日期字符串，否则返回数值字符串。</returns>
    private static string GetNumericCellValue(ICell cell)
    {
        if (DateUtil.IsCellDateFormatted(cell))
        {
            // 日期是按1900/1/0作为0起点，相差的天数就是整数部分，小数部分是这样来的：24h*3600s/h=86400s，那么一天有86400秒，用1/86400*现在经过的秒数，就是小数部分
            var date = DateTime.Parse("1900/1/1").AddDays(-2)
                .AddDays((int)cell.NumericCellValue)
                .AddMilliseconds(Math.Ceiling((cell.NumericCellValue - (int)cell.NumericCellValue) * 1000 * 86400));
            return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }
        return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
    }
}