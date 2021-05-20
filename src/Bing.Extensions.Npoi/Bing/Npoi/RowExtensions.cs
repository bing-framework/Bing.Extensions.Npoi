using System;
using NPOI.SS.UserModel;

namespace Bing.Npoi
{
    /// <summary>
    /// NPOI单元行(<see cref="IRow"/>) 扩展
    /// </summary>
    public static partial class RowExtensions
    {
        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        /// <param name="row">NPOI单元行</param>
        /// <param name="cellIndex">单元格索引</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static ICell GetOrCreateCell(this IRow row, int cellIndex)
        {
            if (row is null)
                throw new ArgumentNullException(nameof(row));
            return row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);
        }

        /// <summary>
        /// 创建单元格并进行操作
        /// </summary>
        /// <param name="row">NPOI单元行</param>
        /// <param name="cellIndex">单元格索引</param>
        /// <param name="setupAct">操作函数</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static IRow CreateCell(this IRow row, int cellIndex, Action<ICell> setupAct)
        {
            if (row is null)
                throw new ArgumentNullException(nameof(row));
            var cell = row.GetOrCreateCell(cellIndex);
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
}
