using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace ToWord
{
    public class WordUtil
    {
        //Word应用程序
        Word.Application application;
        //Word文档
        Word.Document document;
        object nothing = System.Reflection.Missing.Value;
        public WordUtil()
        {
            //打开Word应用程序
            application = new Word.ApplicationClass();
            //应用程序不可见
            application.Visible = false;
            //创建一个Word文档
            document = application.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
        }

        public void AddContent(string content, Model.Style style)
        {
            Word.Paragraph p;
            p = document.Content.Paragraphs.Add(ref nothing);
            p.Range.Text = content;
            if(style != null)
            {
                p.Range.Font.Name = style.Font;
                p.Range.Font.Size = style.Size == 0 ? p.Range.Font.Size : style.Size;
                if(style.LineSpac > 0)
                {
                    p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;  //行距最小值
                    p.LineSpacing = style.LineSpac;    //行距28磅
                }
                p.Range.Font.Color = style.FontColor;
                p.Range.Font.Bold = style.FontStyle == System.Drawing.FontStyle.Bold ? 1 : 0;
                p.CharacterUnitFirstLineIndent = style.Indent == 0 ? p.CharacterUnitFirstLineIndent : style.Indent;    //缩进2字符
            }
            p.Range.InsertParagraphAfter();
        }

        public void AddTitle(string title, Model.Style style)
        {
            //Word段落
            Word.Paragraph p;
            p = document.Content.Paragraphs.Add(ref nothing);
            //设置段落中的内容文本
            p.Range.Text = title;
            if (style != null)
            {
                object bContinuousPrev = false;//是否为大圆点 true 大圆点
                object a = 1;
                object oName = "";
                object listFormat = Word.WdDefaultListBehavior.wdWord10ListBehavior;

                p.Range.Font.Name = style.Font;
                p.Range.Font.Size = style.Size == 0 ? p.Range.Font.Size : style.Size;
                p.Range.Font.Color = style.FontColor;
                p.Range.Font.Bold = style.FontStyle == System.Drawing.FontStyle.Bold ? 1 : 0;
                if (style.LineSpac > 0)
                {
                    p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;  //行距最小值
                    p.LineSpacing = style.LineSpac;    //行距28磅
                }
                p.CharacterUnitFirstLineIndent = style.Indent == 0 ? p.CharacterUnitFirstLineIndent : style.Indent;    //缩进2字符

                Word.ListTemplate it = application.ActiveDocument.ListTemplates.Add(ref bContinuousPrev, ref oName);
                it.ListLevels[1].NumberFormat = style.NumberFormat;
                p.Range.ListFormat.ApplyListTemplate(it, ref bContinuousPrev, ref a, ref listFormat);
            }

            //添加到末尾
            p.Range.InsertParagraphAfter();            
        }

        public void AddLine()
        {
            application.Selection.TypeParagraph();
        }

        /// <summary>
        /// 添加普通段落
        /// </summary>
        /// <param name="s"></param>
        public void AddParagraph(string s)
        {
            Word.Paragraph p;
            p = document.Content.Paragraphs.Add(ref nothing);
            p.Range.Text = s;
            p.Range.Font.Name = "方正仿宋_GBK";
            p.Range.Font.Size = 15.75F;
            p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;  //行距最小值
            p.LineSpacing = 28F;    //行距28磅
            p.CharacterUnitFirstLineIndent = 2F;    //缩进2字符
            p.Range.InsertParagraphAfter();
        }

        public void AddExcel(object fileName)
        {
            object classType = "Excel.Sheet.12";
            object linkToFile = false;
            object displayAsIcon = false;
            GoToTheEnd();
            application.Selection.InlineShapes.AddOLEObject(ref classType, ref fileName, ref linkToFile, ref displayAsIcon);
            application.Selection.TypeParagraph();
        }

        public void AddImage(string strPicPath)
        {
            string fileName = strPicPath;
            object linkToFile = false;
            object saveWithDocument = true;
            GoToTheEnd();
            application.Selection.InlineShapes.AddPicture(fileName, ref linkToFile, ref saveWithDocument);
            application.Selection.TypeParagraph();
        }

        public void GoToTheEnd()
        {
            object unit;
            unit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            application.Selection.EndKey(ref unit, ref nothing);
        }

        /// <summary>
        /// 添加有序列表
        /// </summary>
        /// <param name="s"></param>
        public void AddOrderList(string s)
        {
            Word.Paragraph p;
            p = document.Content.Paragraphs.Add(ref nothing);
            p.Range.Text = s;
            object style = Word.WdBuiltinStyle.wdStyleListNumber;
            p.set_Style(ref style);
            p.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// 保存和关闭
        /// 此处可能根据Word的版本不同略有变化
        /// 可以考虑使用反射的方式，避免这个问题。
        /// </summary>
        /// <param name="fileName">文件名称</param>
        public void SaveAndClose(object fileName)
        {
            object nothing = System.Reflection.Missing.Value;
            document.SaveAs(ref fileName, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            document.Close(ref nothing, ref nothing, ref nothing);
            application.Quit(ref nothing, ref nothing, ref nothing);
        }

        public void formatContent()
        {
            Word.TableOfContents myContent = document.TablesOfContents[1]; //目录  
            Word.Paragraphs myParagraphs = myContent.Range.Paragraphs; //目录里的所有段，一行一段  
            int[] FirstParaArray = new int[3] { 1, 8, 9 }; //一级标题，直接指定  
            foreach (int i in FirstParaArray)
            {
                myParagraphs[i].Range.Font.Bold = 1;  //加粗  
                myParagraphs[i].Range.Font.Name = "黑体"; //字体  
                myParagraphs[i].Range.Font.Size = 12; //小四  
                myParagraphs[i].Range.ParagraphFormat.SpaceBefore = 6; //段前  
                myParagraphs[i].Range.ParagraphFormat.SpaceAfter = 6; //段后间距  
            }
        }
    }
}
