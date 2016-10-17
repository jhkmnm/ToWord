﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using ToWord.Model;
using MSWord = Microsoft.Office.Interop.Word;

namespace ToWord
{
    public partial class Form1 : Form
    { 
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region
            //            正文
            //1.发文字号[方正仿宋 居中三号，〔〕]  固定
            //2.题目[方正小标宋 二号 加粗 居中]     固定
            //3.主送部门[方正仿宋 三号]
            //4.正文[方正仿宋 三号 首行缩进2字符，对齐方式为两端对齐，行间距为28磅]
            //5.一级标题[方正黑体三号，首行缩进2字符]
            //6.二级标题[方正楷体三号，首行缩进2字符]
            //7.三级标题[方正仿宋三号，首行缩进2字符]
            //8.四级标题[方正仿宋三号，首行缩进2字符]
            //9.附件[与正文空一行，方正防宋三号，首行缩进2字符]
            //10.落款[??]
            //11.日期[右对齐并且空4格，方正防宋三号字体]
            //12.此件发至[方正仿宋三号，首行缩进2字符]

            //文字附件
            //1.附件[方正黑体三号]
            //2.题目[方正小标宋二号，居中]
            //3.正文[方正仿宋三号，首行缩进2字符，对齐方式为两端对齐，行间距为28磅]
            //4.一级标题[方正黑体三号，首行缩进2字符]
            //5.二级标题[方正楷体三号，首行缩进2字符]
            //6.三级标题[方正仿宋三号，首行缩进2字符]
            //7.四级标题[方正仿宋三号，首行缩进2字符]

            //版记
            //1.抄送[方正仿宋，四号]
            //2.结尾[方正仿宋，四号]  固定
            //三号15.75  二号21.75
            #endregion

            Config config = new Config
            {
                FontStyles = new List<Style> {
                    new Style { FontName = "标题", Font = "方正小标宋_GBK", Size = 33, FontStyle = FontStyle.Bold, FontColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify },
                    new Style { FontName = "发文字号", Font = "方正仿宋_GBK", Size = 15.75F, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter },
                    new Style { FontName = "题目", Font = "方正小标宋_GBK", Size = 21.75F, FontStyle = FontStyle.Bold, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter },
                    new Style { FontName = "主送部门", Font = "方正仿宋_GBK", Size = 15.75F },
                    new Style { FontName = "正文", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify, LineSpac = 28F },
                    new Style { FontName = "一级标题", Font = "方正黑体_GBK", Size = 15.75F, Indent = 2, NumberFormat = "{0}、" },
                    new Style { FontName = "二级标题", Font = "方正楷体_GBK", Size = 15.75F, Indent = 2, NumberFormat = "({0})" },
                    new Style { FontName = "三级标题", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2, NumberFormat = "{0}." },
                    new Style { FontName = "四级标题", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2, NumberFormat = "({0})" },
                    new Style { FontName = "正文附件", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2},
                    new Style { FontName = "落款"},
                    new Style { FontName = "日期", Font = "方正仿宋_GBK", Size = 15.75F, Align = MSWord.WdParagraphAlignment.wdAlignParagraphRight, Indent = 4},
                    new Style { FontName = "发送至", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2},
                    new Style { FontName = "附件中附件", Font = "方正黑体_GBK", Size = 15.75F},
                    new Style { FontName = "附件题目", Font = "方正小标宋_GBK", Size = 21.75F, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter},
                    new Style { FontName = "附件正文", Font = "方正仿宋_GBK", Size = 15.75F, Indent = 2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify, LineSpac = 28F},
                    new Style { FontName = "抄送", Font = "方正仿宋_GBK", Size = 14F},
                    new Style { FontName = "结尾", Font = "方正仿宋_GBK", Size = 14F}
                }
            };
            #region
            //object path;//文件路径
            //string strContent;//文件内容
            //MSWord.Application wordApp;//Word应用程序变量
            //MSWord.Document wordDoc;
            //path = "d:\\myWord.doc";
            //wordApp = new MSWord.ApplicationClass();//初始化
            //if (File.Exists((string)path))
            //{
            //    File.Delete((string)path);
            //}
            //Object Nothing = Missing.Value;
            //wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            ////wordApp.Selection.ParagraphFormat.LineSpacing = 35f;//设置文档的行间距
            ////写入普通文本
            //wordApp.Selection.ParagraphFormat.FirstLineIndent = 30;//首行缩进的长度

            //strContent = "c#向Word写入文本   普通文本:\n";
            //wordDoc.Paragraphs.Last.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;

            //wordDoc.Paragraphs.Last.Range.Font.Name = "方正仿宋_GBK";
            //wordDoc.Paragraphs.Last.Range.Font.Size = 15.75F;
            //wordDoc.Paragraphs.Last.Range.Font.Bold = 1;
            //wordDoc.Paragraphs.Last.Range.Text = strContent;
            //wordDoc.Paragraphs.Last.LineSpacing = 28F;

            //object oOutlineNumbered, oName, bContinuousPrev, applyTo, defaultListBehaviour;
            //oOutlineNumbered = oName = bContinuousPrev = applyTo = defaultListBehaviour = null;
            //MSWord.ListTemplate it = wordApp.ActiveDocument.ListTemplates.Add(ref oOutlineNumbered, ref oName);
            //it.ListLevels[1].NumberFormat = "一、";
            //wordApp.ActiveDocument.Paragraphs[1].Range.ListFormat.ApplyListTemplate(it, ref bContinuousPrev, ref applyTo, ref defaultListBehaviour);
            //string[] s = { "一", "一", "1", "1" };

            //for (int m = 4; m <= 7; m++)
            //{
            //    it = wordApp.ActiveDocument.ListTemplates.Add(ref oOutlineNumbered, ref oName);
            //    it.ListLevels[1].NumberFormat = m.ToString();
            //    wordApp.ActiveDocument.Paragraphs[1].Range.ListFormat.ApplyListTemplate(it, ref bContinuousPrev, ref applyTo, ref defaultListBehaviour);
            //    //wordApp.Selection.Range.ListFormat.ListTemplate.ListLevels[1].NumberFormat = m.ToString();//string.Format(config.FontStyles[m].NumberFormat, s[m - 4])
            //    wordDoc.Paragraphs.Last.Range.Text = strContent;
            //    wordApp.Selection.TypeParagraph();
            //}            

            //object format = MSWord.WdSaveFormat.wdFormatDocument;
            //wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            ////Utilities.XmlHelper.XmlSerializeToFile(config, "config.xml", Encoding.UTF8);      
            #endregion

            WordUtil word = new ToWord.WordUtil();
            string[] titleA = { "一", "二", "三", "四" };
            string[] titleB = { "1", "2", "3", "4" };

            List<Content> contents = new List<Content>();
            contents.AddRange(
                new[]
                {
                    new Content { ConType = "标题", ConContent = "国网重庆市电力公司永川供电分公司文件" },
                    new Content { ConType = "发文字号", ConContent = "发文字号"},
                    new Content { ConType = "题目", ConContent = "国网重庆市电力公司永川供电分公司关于XX的通知"},
                    new Content { ConType = "正文", ConContent = "（首行缩进2个字符）x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x〔2013〕x号 x x x x x x x x。（正文 方正仿宋 三号）"},
                    new Content { ConType = "一级标题", ConContent = "一级标题", TitleIndex = 0},
                    new Content { ConType = "二级标题", ConContent = "二级标题", TitleIndex = 0},
                    new Content { ConType = "三级标题", ConContent = "三级标题", TitleIndex = 0},
                    new Content { ConType = "正文", ConContent = "x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x"},
                    new Content { ConType = "四级标题", ConContent = "四级标题", TitleIndex = 0},
                    new Content { ConType = "正文", ConContent = "x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x x"},
                    new Content { ConType = "一级标题", ConContent = "一级标题", TitleIndex = 1}
                }
            );

            foreach(var content in contents)
            {
                var style = config.FontStyles.Find(w => w.FontName == content.ConType);
                if (content.ConType.EndsWith("级标题"))
                {
                    string index = (content.ConType == "一级标题" || content.ConType == "二级标题") ? titleA[content.TitleIndex] : titleB[content.TitleIndex];
                    style.NumberFormat = string.Format(style.NumberFormat, index);
                    word.AddTitle(content.ConContent, style);
                }
                else if(content.ConType == "")//附件
                { }
                else
                {
                    word.AddContent(content.ConContent, style);
                }
            }
            
            word.SaveAndClose("D:\\123.doc");
        }

        public class Content
        {
            public string ConType { get; set; }

            public string ConContent { get; set; }

            public int TitleIndex { get; set; }
        }
    }
}
