using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ArticleArray.Models.FontSizeModel;

namespace ArticleArray
{
    public class FontHelper
    {
        public static float GetFontSize(FontSize fontSize)
        {
            return (float)fontSize;
        }

        public static FontSize GetFontSize(float fontSize)
        {
            return (FontSize)fontSize;
        }
    }
}
