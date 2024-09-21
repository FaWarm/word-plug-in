using Microsoft.Office.Interop.Word;

namespace ArticleFormatApp_USST.Models
{
    public class StyleModel
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string NameLocal { get; set; }

        /// <summary>
        /// 基础样式
        /// </summary>
        public WdBuiltinStyle BaseStyle { get; set; } = WdBuiltinStyle.wdStyleNormal;

        /// <summary>
        /// 下一段样式名称
        /// </summary>
        public string NextParagraphStyle { get; set; } = string.Empty;

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; } = "Times New Roman";

        /// <summary>
        /// 中文字体名称
        /// </summary>
        public string FontNameFarEast { get; set; } = "宋体";

        /// <summary>
        /// 字体大小
        /// </summary>
        public float FontSize { get; set; } = 12f;

        /// <summary>
        /// 字体加粗
        /// </summary>
        public int FontBold { get; set; } = 0;

        /// <summary>
        /// 字体倾斜
        /// </summary>
        public bool FontItalic { get; set; } = false;

        /// <summary>
        /// 字体下划线
        /// </summary>
        public WdUnderline FontUnderline { get; set; } = WdUnderline.wdUnderlineNone;

        /// <summary>
        /// 字体颜色
        /// </summary>
        public WdColor FontColor { get; set; } = WdColor.wdColorBlack;

        /// <summary>
        /// 段落对齐
        /// </summary>
        public WdParagraphAlignment ParaAlignment { get; set; } = WdParagraphAlignment.wdAlignParagraphJustify;

        /// <summary>
        /// 首行或悬挂缩进的值 (以字符为单位)
        /// </summary>
        public float ParaCharacterUnitFirstLineIndent { get; set; } = 0f;

        /// <summary>
        /// 段落的左缩进值 (以字符为单位)
        /// </summary>
        public float ParaCharacterUnitLeftIndent { get; set; } = 0f;

        /// <summary>
        /// 段落的右缩进量（以字符为单位）
        /// </summary>
        public float ParaCharacterUnitRightIndent { get; set; } = 0f;

        /// <summary>
        /// 首行的行或悬挂缩进的值
        /// </summary>
        public float ParaFirstLineIndent { get; set; } = 0f;

        /// <summary>
        /// 段落左缩进值
        /// </summary>
        public float ParaLeftIndent { get; set; } = 0f;

        /// <summary>
        /// 段前间距
        /// </summary>
        public float ParaSpaceBefore { get; set; } = 0f;

        /// <summary>
        /// 段后间距
        /// </summary>
        public float ParaSpaceAfter { get; set; } = 0f;

        /// <summary>
        /// 段落的行距
        /// </summary>
        public WdLineSpacing ParaLineSpacingRule { get; set; } = WdLineSpacing.wdLineSpaceSingle;

        /// <summary>
        /// 段落的行距值 (以磅为单位)
        /// </summary>
        public float ParaLineSpacing { get; set; } = 20f;

        /// <summary>
        /// 段落的大纲级别
        /// </summary>
        public WdOutlineLevel OutlineLevel { get; set; } = WdOutlineLevel.wdOutlineLevelBodyText;

        public string LinkToListTemplate { get; set; } = string.Empty;
    }
}
