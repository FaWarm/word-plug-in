using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ArticleArray.Models
{
    public class OneKeySettingModel
    {
        public string SettingChose { get; set; }
        public string ChoseTemplate { get; set; }
        public string Heading1 { get; set; }
        public string Heading2 { get; set; }
        public string Heading3 { get; set; }
        public string Heading4 { get; set; }
        public string Heading5 { get; set; }
        public string Abstract { get; set; }
        public string Catalog { get; set; }
        public string Reference { get; set; }

        public bool PicCenter { get; set; }
        public string PicWidth { get; set; }
        public string PicHeight { get; set; }
        public string TableAutoFit { get; set; }
        public string ThreeLineStyle { get; set; }
        public string OddHeader { get; set; }
        public string EvenHeader { get; set; }
        public bool ParityHeader { get; set; }//奇偶页眉
        public bool Pagination { get; set; }//页码
        public bool PaginationCenter { get; set; }//页码居中
        public bool PaginationContinue { get; set; }//页码续前节
        public bool TableCaption { get; set; }//表题注
        public bool PicCaption { get; set; }//图题注
        public bool Seg { get; set; }//一级标题分节
        public bool Normal1 { get; set; }//正文1转正文
        public bool RmSpace { get; set; }//移除正文1空格

        
    }
}
