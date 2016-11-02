using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Word;

namespace ToWord.Model
{
    public class Config
    {
        [XmlAttribute]
        public List<Style> FontStyles;

        public Style this[string Name]
        {
            get
            {
                Style s = null;
                foreach(Style st in FontStyles)
                {
                    if(string.Compare(st.FontName, Name) == 0)
                    {
                        s = st;
                        break;
                    }
                }
                return s;
            }            
        }
    }

    public class Style
    {
        /// <summary>
        /// 样式名称
        /// </summary>
        [XmlAttribute]
        public string FontName { get; set; }

        /// <summary>
        /// 字体
        /// </summary>
        [XmlAttribute]
        public string Font { get; set; }

        /// <summary>
        /// 大小
        /// </summary>
        [XmlAttribute]
        public float Size { get; set; }

        /// <summary>
        /// 字体样式
        /// </summary>
        [XmlAttribute]
        public FontStyle FontStyle { get; set; }

        [XmlAttribute]
        public WdColor FontColor { get; set; }

        /// <summary>
        /// 对齐方式
        /// </summary>
        [XmlAttribute]
        public WdParagraphAlignment Align { get; set; }

        /// <summary>
        /// 行间距        
        /// </summary>
        [XmlAttribute]
        public float LineSpac { get; set; }

        /// <summary>
        /// 缩进
        /// </summary>
        [XmlAttribute]
        public int Indent { get; set; }

        [XmlAttribute]
        /// <summary>
        /// 编号格式
        /// </summary>
        public string NumberFormat { get; set; }

        [XmlAttribute]
        public string Number { get; set; }

        public int Scaling { get; set; }
    }
}
