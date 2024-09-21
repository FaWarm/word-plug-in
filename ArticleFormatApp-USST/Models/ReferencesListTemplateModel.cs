using Microsoft.Office.Interop.Word;

namespace ArticleFormatApp_USST.Models
{
    public class ReferencesListTemplateModel
    {
        public string Name { get; set; } = "参考文献编号";
        public string NumberFormat { get; set; } = "[%1]";
        public WdTrailingCharacter TrailingCharacter { get; set; } = WdTrailingCharacter.wdTrailingSpace;
    }
}
