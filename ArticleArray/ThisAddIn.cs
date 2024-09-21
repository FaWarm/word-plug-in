using Microsoft.Office.Interop.Word;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace ArticleArray
{
    public partial class ThisAddIn
    {
        public static ThisAddIn ThisAddInInstance { get; private set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ThisAddInInstance = this;
            // 在文档加载时注册样式更改事件
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 在插件停止时取消事件注册
            this.Application.DocumentOpen -= new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
        }

        private void Application_DocumentOpen(Word.Document doc)
        {
            // 加载文档样式
            LoadStyles(doc);
        }

        public void LoadStyles(Word.Document doc)
        {
            // 遍历文档中的所有样式
            foreach (Word.Style style in doc.Styles)
            {
                // 将样式信息输出到控制台
                Console.WriteLine(style.NameLocal + " - " + style.Font.Name);
            }
        }

        // 设置样式
        public void SetStyle(Document doc, string styleName, string fontName, float fontSize)
        {
            // 获取指定名称的样式
            Style style = doc.Styles[styleName];

            // 设置字体名称和字号
            style.Font.Name = fontName;
            style.Font.Size = fontSize;

        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
