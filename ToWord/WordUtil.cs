using System;
using System.Collections.Generic;
using System.Text;
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
                p.Format.Alignment = style.Align;
                p.Range.Font.Color = style.FontColor;
                p.Range.Font.Bold = style.FontStyle == System.Drawing.FontStyle.Bold ? 1 : 0;
                p.CharacterUnitFirstLineIndent = style.Indent == 0 ? p.CharacterUnitFirstLineIndent : style.Indent;    //缩进2字符
                p.Range.Font.Scaling = style.Scaling == 0 ? 100 : style.Scaling;
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
                //Word.ListGallery l = application.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];
                //l.ListTemplates[ref a].ListLevels[1].NumberFormat = style.NumberFormat;
                Word.ListTemplate it = application.ActiveDocument.ListTemplates.Add(ref bContinuousPrev, ref oName);
                it.ListLevels[1].NumberFormat = style.NumberFormat;
                //application.ActiveDocument.
                //p.Range.ListFormat.ListTemplate..ApplyListTemplate(l, ref bContinuousPrev, ref a, ref listFormat);
            }

            //添加到末尾
            p.Range.InsertParagraphAfter();            
        }

        public void AddLine()
        {
            Word.Paragraph p;
            p = document.Content.Paragraphs.Add(ref nothing);
            p.Range.Text = "";
            p.Range.InsertParagraphAfter();
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
            var excel = application.Selection.InlineShapes.AddOLEObject(ref classType, ref fileName, ref linkToFile, ref displayAsIcon);
            //excel.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            excel.Fill.Solid();
            excel.Fill.Transparency = 0f;
            excel.Line.Weight = 0.75f;
            excel.Line.Transparency = 0f;
            excel.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            excel.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            excel.Height = 386.95f;
            excel.Width = 479.35f;
            excel.PictureFormat.Brightness = 0.5f;
            excel.PictureFormat.Contrast = 0.5f;
            excel.PictureFormat.ColorType = Microsoft.Office.Core.MsoPictureColorType.msoPictureAutomatic;
            excel.PictureFormat.CropLeft = 0f;
            excel.PictureFormat.CropRight = 0f;
            excel.PictureFormat.CropTop = 0f;
            excel.PictureFormat.CropBottom = 0f;
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

        public void AddTable(string content, Model.Style style)
        {
            object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            object autofitbehavior = Word.WdAutoFitBehavior.wdAutoFitFixed;
            GoToTheEnd();
            Word.Table tablea = null;
            if (!string.IsNullOrWhiteSpace(content))
            {
                tablea = application.ActiveDocument.Tables.Add(application.Selection.Range, 1, 3, ref defaultTableBehavior, ref autofitbehavior);
                application.Selection.ParagraphFormat.SpaceBefore = 0f;
                application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
                application.Selection.ParagraphFormat.FirstLineIndent = application.CentimetersToPoints(0.3f);
                GoToTheEnd();
            }

            var tableb = application.ActiveDocument.Tables.Add(application.Selection.Range, 1, 2, ref defaultTableBehavior, ref autofitbehavior);
            application.Selection.ParagraphFormat.SpaceBefore = 0f;
            application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
            application.Selection.ParagraphFormat.FirstLineIndent = application.CentimetersToPoints(0.3f);

            SetTableStyle(1, style);

            application.Selection.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast;
            application.Selection.Rows.Height = application.CentimetersToPoints(1f);
            application.Selection.Columns.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            application.Selection.Columns.PreferredWidth = 0;
            application.Selection.Cells.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            application.Selection.Cells.PreferredWidth = 0;
            application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            if (tablea != null)
            {
                var cella = tablea.Cell(1, 1);
                cella.Width = 60f;
                cella.Range.Text = "抄送:";

                var celld = tablea.Cell(1, 2);
                celld.Range.Text = content;                
            }

            var cellc = tableb.Cell(1, 2);
            cellc.Range.Text = DateTime.Today.ToString("yyyy年MM月dd日") + "印发";
            //cellc.Range.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignRights, 0);
            cellc.Range.FitTextWidth = application.CentimetersToPoints(4.00f);

            var cellb = tableb.Cell(1, 1);            
            cellb.Range.Text = "国网重庆市电力公司永川供电分公司办公室";
            cellb.Range.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignLefts, 0);
            //cellb.Range.FitTextWidth = application.CentimetersToPoints(1.22f);

            //row = application.Selection.Tables[1].Rows[2];
            //row.Cells[1].Range.Text = "国网重庆市电力公司永川供电分公司办公室";
            //row.Cells[1].Range.Application.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            //row.Cells[2].Range.Text = DateTime.Today.ToString("yyyy年MM月dd日")+"印发";
            //row.Cells[2].Range.Application.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        private void SetTableStyle(int tableindex, Model.Style style)
        {
            object prop = "网格型";
            application.Selection.Tables[tableindex].set_Style(ref prop);
            application.Selection.Tables[tableindex].ApplyStyleHeadingRows = true;
            application.Selection.Tables[tableindex].ApplyStyleLastRow = false;
            application.Selection.Tables[tableindex].ApplyStyleFirstColumn = true;
            application.Selection.Tables[tableindex].ApplyStyleLastColumn = false;
            application.Selection.Tables[tableindex].ApplyStyleRowBands = true;
            application.Selection.Tables[tableindex].ApplyStyleColumnBands = false;
            application.Selection.Tables[tableindex].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            application.Selection.Tables[tableindex].Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            application.Selection.Tables[tableindex].Borders[Word.WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            //application.Selection.Tables[tableindex].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            application.Selection.Tables[tableindex].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            application.Selection.Tables[tableindex].PreferredWidth = application.CentimetersToPoints(15.6f);
            application.Selection.Tables[tableindex].Rows.WrapAroundText = 1;

            application.Selection.Tables[tableindex].Range.Font.Name = style.Font;
            application.Selection.Tables[tableindex].Range.Font.Name = style.Font;
            application.Selection.Tables[tableindex].Range.Font.Size = style.Size;
            application.Selection.Tables[tableindex].Range.Font.Color = style.FontColor;
            application.Selection.Tables[tableindex].Range.Font.Bold = style.FontStyle == System.Drawing.FontStyle.Bold ? 1 : 0;
        }

        public void GoToTheEnd()
        {
            object unit;
            unit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            application.Selection.EndKey(ref unit, ref nothing);
        }

        public void NewPage()
        {
            GoToTheEnd();
            object type = Word.WdBreakType.wdPageBreak;
            application.Selection.InsertBreak(ref type);
            //Selection.InsertBreak Type:= wdPageBreak
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

        public void InsertPageNumber(string strType, bool bHeader)
        {
            object oAlignment = Word.WdPageNumberAlignment.wdAlignPageNumberOutside;
            object oFirstPage = bHeader;
            Word.WdHeaderFooterIndex WdFooterIndex = Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;
            var page = application.Selection.Sections[1].Footers[WdFooterIndex].PageNumbers;
            page.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleNumberInDash;
            page.HeadingLevelForChapter = 0;
            page.ChapterPageSeparator = Word.WdSeparatorType.wdSeparatorHyphen;
            page.RestartNumberingAtSection = false;
            page.StartingNumber = 0;
            page.Add(ref oAlignment, ref oFirstPage);
            //application.ActiveWindow.ActivePane.View.Zoom.PageFit = Word.WdPageFit.wdPageFitFullPage;
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
    }
}
