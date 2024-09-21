using Microsoft.Office.Interop.Word;

namespace ArticleFormatApp_USST.Models
{
    /// <summary>
    /// 多级列表数字格式
    /// </summary>
    public class HeadingListTemplateModel
    {
        /// <summary>
        /// 多级列表标题
        /// </summary>
        public string Name { get; set; } = "自定义标题多级列表";
        //public string BaseFontName { get; set; } = "宋体";
        //public string BaseFontNameFarEast { get; set; } = "Times New Roman";
        //public int FontBold { get; set; } = 0;
        public WdListNumberStyle NumberStyle { get; set; } = WdListNumberStyle.wdListNumberStyleArabic;
        public float TextPosition { get; set; } = 0f;
        public float NumberPosition { get; set; } = 0f;
        public WdTrailingCharacter TrailingCharacter { get; set; } = WdTrailingCharacter.wdTrailingSpace;

        public string NumberFormat1 { get; set; } = "%1";
        public string NumberFormat2 { get; set; } = "%1.%2";
        public string NumberFormat3 { get; set; } = "%1.%2.%3";
        public string NumberFormat4 { get; set; } = "%1.%2.%3.%4";
        public string NumberFormat5 { get; set; } = "%1.%2.%3.%4.%5";
        public string NumberFormat6 { get; set; } = "%1.%2.%3.%4.%5.%6";
        public string NumberFormat7 { get; set; } = "%1.%2.%3.%4.%5.%6.%7";
        public string NumberFormat8 { get; set; } = "%1.%2.%3.%4.%5.%6.%7.%8";
        public string NumberFormat9 { get; set; } = "%1.%2.%3.%4.%5.%6.%7.%8.%9";

    }
}
