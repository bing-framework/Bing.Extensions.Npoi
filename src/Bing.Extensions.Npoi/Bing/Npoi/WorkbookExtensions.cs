using System;
using System.IO;
using System.Linq;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;

namespace Bing.Npoi;

/// <summary>
/// NPOI工作簿(<see cref="IWorkbook"/>) 扩展
/// </summary>
public static partial class WorkbookExtensions
{
    /// <summary>
    /// 将工作簿转换为字节数组。
    /// </summary>
    /// <param name="workbook">要转换的工作簿。</param>
    /// <param name="leaveOpen">是否保持工作簿打开状态，默认为false。</param>
    /// <returns>表示工作簿的字节数组。</returns>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static byte[] ToBytes(this IWorkbook workbook, bool leaveOpen = false)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        using var stream = new MemoryStream();
        workbook.Write(stream, leaveOpen);
        return stream.ToArray();
    }

    /// <summary>
    /// 将工作簿写入指定的文件中。
    /// </summary>
    /// <param name="workbook">要写入的工作簿。</param>
    /// <param name="fileName">目标文件名。</param>
    /// <param name="leaveOpen">是否保持工作簿打开状态，默认为false。</param>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static void ToFile(this IWorkbook workbook, string fileName, bool leaveOpen = false)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        using var fs = new FileStream(fileName, FileMode.Create);
        workbook.Write(fs, leaveOpen);
    }

    /// <summary>
    /// 计算工作簿中所有公式单元格的值。
    /// </summary>
    /// <param name="workbook">要计算的工作簿。</param>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static void EvaluateAllFormulaCells(this IWorkbook workbook)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        BaseFormulaEvaluator.EvaluateAllFormulaCells(workbook);
    }

    /// <summary>
    /// 计算工作簿中所有公式单元格的值，忽略任何错误。
    /// </summary>
    /// <param name="workbook">要计算的工作簿。</param>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static void EvaluateAllFormulaCellsIgnoreError(this IWorkbook workbook)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        for (var i = 0; i < workbook.NumberOfSheets; i++)
        {
            foreach (IRow item in workbook.GetSheetAt(i))
            {
                foreach (var cell in item)
                {
                    if (cell.CellType == CellType.Formula)
                    {
                        try
                        {
                            evaluator.EvaluateFormulaCell(cell);
                        }
                        catch
                        {
                            // 不做处理，为了后续执行后续单元格
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// 创建一个公式求值器，用于计算工作簿中的公式单元格。
    /// </summary>
    /// <param name="workbook">要创建公式求值器的工作簿。</param>
    /// <returns>创建的公式求值器。</returns>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static IFormulaEvaluator CreateFormulaEvaluator(this IWorkbook workbook)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        return workbook.GetCreationHelper().CreateFormulaEvaluator();
    }

    /// <summary>
    /// 从工作簿中获取第一个存在的工作表。
    /// </summary>
    /// <param name="workbook">要获取工作表的工作簿。</param>
    /// <param name="names">工作表名称列表，按照顺序查找，返回第一个存在的工作表。</param>
    /// <returns>找到的第一个工作表，如果没有找到则返回null。</returns>
    /// <exception cref="ArgumentNullException">如果工作簿为null，则抛出此异常。</exception>
    public static ISheet GetSheet(this IWorkbook workbook, params string[] names)
    {
        if (workbook is null)
            throw new ArgumentNullException(nameof(workbook));
        return names.Select(workbook.GetSheet).FirstOrDefault(sheet => sheet != null);
    }
}