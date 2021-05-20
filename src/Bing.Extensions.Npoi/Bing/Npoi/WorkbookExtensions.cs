using System;
using System.IO;
using NPOI.SS.UserModel;

namespace Bing.Npoi
{
    /// <summary>
    /// NPOI工作簿(<see cref="IWorkbook"/>) 扩展
    /// </summary>
    public static partial class WorkbookExtensions
    {
        /// <summary>
        /// 将工作簿保存到二进制文件流
        /// </summary>
        /// <param name="workbook">NPOI工作簿</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static byte[] SaveToBuffer(this IWorkbook workbook)
        {
            if (workbook is null)
                throw new ArgumentNullException(nameof(workbook));
            using var ms = new MemoryStream();
            workbook.Write(ms);
            return ms.ToArray();
        }
    }
}
