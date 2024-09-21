namespace ArticleFormatApp_USST.Models
{
    public class OtherConfigModel
    {
        //多级列表参数
        public HeadingListTemplateModel HeadingListTemplate { get; set; } = new HeadingListTemplateModel();
        //参考文献列表
        public ReferencesListTemplateModel ReferencesListTemplate { get; set; } = new ReferencesListTemplateModel();
    }
}
