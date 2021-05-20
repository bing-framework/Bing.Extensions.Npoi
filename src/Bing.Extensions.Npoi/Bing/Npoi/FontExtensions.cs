using NPOI.SS.UserModel;

namespace Bing.Npoi
{
    /// <summary>
    /// 字体(<see cref="IFont"/>) 扩展
    /// </summary>
    public static partial class FontExtensions
    {
        /// <summary>
        /// 设置字体大小
        /// </summary>
        /// <param name="font">字体</param>
        /// <param name="fontSize">字体大小</param>
        public static IFont SetFontHeightInPoints(this IFont font, short fontSize)
        {
            font.FontHeightInPoints = fontSize;
            return font;
        }

        /// <summary>
        /// 设置字体颜色
        /// </summary>
        /// <param name="font">字体</param>
        /// <param name="color">颜色</param>
        public static IFont SetColor(this IFont font, short color)
        {
            font.Color = color;
            return font;
        }

        /// <summary>
        /// 设置粗体
        /// </summary>
        /// <param name="font">字体</param>
        /// <param name="isBold">粗体大小</param>
        public static IFont SetBold(this IFont font, bool isBold)
        {
            font.IsBold = isBold;
            return font;
        }

        /// <summary>
        /// 默认字体
        /// </summary>
        /// <param name="font">字体</param>
        public static IFont DefaultFont(this IFont font)
        {
            font.FontName = "宋体";
            font.FontHeightInPoints = 9;
            return font;
        }
    }
}
