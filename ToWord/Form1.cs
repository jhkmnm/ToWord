﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ToWord.Model;
using MSWord = Microsoft.Office.Interop.Word;

namespace ToWord
{
    public partial class Form1 : Form
    {
        Config config;
        Notice n;
        WordUtil word;
        string[] titleA = { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };
        string[] titleB = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" };

        public Form1()
        {
            InitializeComponent();

            DDLSource source = new DDLSource { Index = -1, Text = "新增" };
            dDLSourceBindingSource.DataSource = source;
            
            BindDDL();
        }

        private void BindDDL()
        {
            ddlWord.DataSource = null;
            ddlWord.DataSource = dDLSourceBindingSource;
            ddlWord.ValueMember = "Index";
            ddlWord.DisplayMember = "Text";
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
            //三号16  二号21.75
            #endregion
            float font2 = 22f;
            float font3 = 16f;
            float font4 = 14f;

            config = new Config
            {                
                FontStyles = new List<Style> {
                    new Style { FontName = "标题", Font = "方正小标宋_GBK", Size = 33, FontStyle = FontStyle.Bold, FontColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify, Scaling = 68 },
                    new Style { FontName = "发文字号", Font = "方正仿宋_GBK", Size = font3, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter },
                    new Style { FontName = "题目", Font = "方正小标宋_GBK", Size = font2, FontStyle = FontStyle.Bold, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter },
                    new Style { FontName = "主送部门", Font = "方正仿宋_GBK", Size = font3 },
                    new Style { FontName = "正文", Font = "方正仿宋_GBK", Size = font3, Indent = 2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify, LineSpac = 28F },
                    new Style { FontName = "一级标题", Font = "方正黑体_GBK", Size = font3, Indent = 2, NumberFormat = "{0}、" },
                    new Style { FontName = "二级标题", Font = "方正楷体_GBK", Size = font3, Indent = 2, NumberFormat = "（{0}） " },
                    new Style { FontName = "三级标题", Font = "方正仿宋_GBK", Size = font3, Indent = 2, NumberFormat = "{0}." },
                    new Style { FontName = "四级标题", Font = "方正仿宋_GBK", Size = font3, Indent = 2, NumberFormat = "（{0}）" },
                    new Style { FontName = "正文表格", Font = "方正小标宋_GBK", Size = font2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter },
                    new Style { FontName = "附件列表", Font = "方正仿宋_GBK", Size = font3, Indent = 2},
                    new Style { FontName = "正文附件", Font = "方正黑体_GBK", Size = font3, Indent = 2},
                    new Style { FontName = "落款", Font = "方正仿宋_GBK", Size = font3, Align = MSWord.WdParagraphAlignment.wdAlignParagraphRight},
                    new Style { FontName = "日期", Font = "方正仿宋_GBK", Size = font3, Align = MSWord.WdParagraphAlignment.wdAlignParagraphRight, Indent = 4},
                    new Style { FontName = "发送至", Font = "方正仿宋_GBK", Size = font3, Indent = 2},
                    new Style { FontName = "附件中附件", Font = "方正黑体_GBK", Size = font3},
                    new Style { FontName = "附件题目", Font = "方正小标宋_GBK", Size = font2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphCenter},
                    new Style { FontName = "附件正文", Font = "方正仿宋_GBK", Size = font3, Indent = 2, Align = MSWord.WdParagraphAlignment.wdAlignParagraphJustify, LineSpac = 28F},
                    new Style { FontName = "抄送", Font = "方正仿宋_GBK", Size = font4},
                    new Style { FontName = "结尾", Font = "方正仿宋_GBK", Size = font4}
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
        }

        private void btnNotice_Click(object sender, EventArgs e)
        {            
            FormNotice f = new FormNotice(n);
            if(f.ShowDialog() == DialogResult.OK)
            {
                n = f.Notice;                
            }
        }

        private void ExportWord(Notice n)
        {
            word = new ToWord.WordUtil();

            var biaoti = new Content { ConType = "标题", ConContent = "国网重庆市电力公司永川供电分公司文件" };
            var style = config.FontStyles.Find(w => w.FontName == biaoti.ConType);
            word.AddContent(biaoti.ConContent, style);

            //空一行
            word.AddLine();

            style = config.FontStyles.Find(w => w.FontName == n.Fawen.ConType);
            word.AddContent(n.Fawen.ConContent, style);

            word.AddImage(System.Environment.CurrentDirectory+"\\a.png");

            style = config.FontStyles.Find(w => w.FontName == n.Timu.ConType);
            word.AddContent(n.Timu.ConContent, style);

            //空一行
            word.AddLine();

            style = config.FontStyles.Find(w => w.FontName == n.Bumen.ConType);
            word.AddContent(n.Bumen.ConContent, style);

            style = config.FontStyles.Find(w => w.FontName == n.BumenZhenwen.ConType);
            word.AddContent(n.BumenZhenwen.ConContent, style);

            int a, b, c, d = 0;
            a = b = c = d;

            if (n.Zhenwen != null)
            {
                foreach (var v in n.Zhenwen)
                {
                    style = config.FontStyles.Find(w => w.FontName == v.TContent.ConType);
                    if (v.TContent.ConType == "一级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleA[a++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "二级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleA[b++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "三级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleB[c++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "四级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleB[d++]) + v.TContent.ConContent;
                    }
                    if (v.TContent.ConType == "正文表格")
                    {
                        style = config.FontStyles.Find(w => w.FontName == "正文表格");
                        word.AddContent(v.TContent.ConContent, style);
                        //word.AddExcel(((FileAppendix)v.Appendixs[0]).FilePath);
                        var path = ((FileAppendix)v.Appendixs[0]).FilePath;
                        if (path.EndsWith(".xls") || path.EndsWith(".xlsx"))
                        {
                            word.AddExcel(path);
                        }
                        else
                        {
                            word.AddImage(path);
                        }
                    }
                    else
                    {
                        word.AddContent(v.TContent.ConContent, style);
                    }
                }
            }

            //空一行
            word.AddLine();

            //附件列表
            int index = 1;
            if (n.Appendixs != null)
            {
                style = config.FontStyles.Find(w => w.FontName == "附件列表");
                n.Appendixs.ForEach(item =>
                {
                    if (index == 1)
                        word.AddContent(string.Format("附件： {0}. {1}", index, item.Title), style);
                    else
                        word.AddContent(string.Format("　　　 {0}. {1}", index, item.Title), style);
                    index++;
                });
            }

            style = config.FontStyles.Find(w => w.FontName == "落款");
            word.AddContent("国网重庆市电力公司永川供电分公司", style);

            style = config.FontStyles.Find(w => w.FontName == "日期");
            word.AddContent(DateTime.Today.ToString("yyyy年MM月dd日")+ "　　　 ", style);

            //空一行
            word.AddLine();

            //style = config.FontStyles.Find(w => w.FontName == "发送至");
            //word.AddContent("国网重庆市电力公司永川供电分公司", style);

            AddAppendixs(n.Appendixs, 0);
            #region
            //index = 1;
            //if (n.Appendixs != null)
            //{
            //    foreach (var appendix in n.Appendixs)
            //    {
            //        style = config.FontStyles.Find(w => w.FontName == "正文附件");
            //        word.AddContent("附件" + index.ToString(), style);

            //        if (appendix.Type == 0)
            //        {
            //            #region 正文中的文字附件
            //            var wa = (WordsAppendix)appendix;

            //            style = config.FontStyles.Find(w => w.FontName == wa.Timu.ConType);
            //            word.AddContent(wa.Timu.ConContent, style);

            //            style = config.FontStyles.Find(w => w.FontName == wa.TZhenwen.ConType);
            //            word.AddContent(wa.TZhenwen.ConContent, style);

            //            a = b = c = d = 0;

            //            foreach (var v in wa.Zhenwen)
            //            {
            //                style = config.FontStyles.Find(w => w.FontName == v.TContent.ConType);
            //                if (v.TContent.ConType == "一级标题")
            //                {
            //                    v.TContent.ConContent = string.Format(style.NumberFormat, titleA[a++]) + v.TContent.ConContent;
            //                }
            //                else if (v.TContent.ConType == "二级标题")
            //                {
            //                    v.TContent.ConContent = string.Format(style.NumberFormat, titleA[b++]) + v.TContent.ConContent;
            //                }
            //                else if (v.TContent.ConType == "三级标题")
            //                {
            //                    v.TContent.ConContent = string.Format(style.NumberFormat, titleB[c++]) + v.TContent.ConContent;
            //                }
            //                else if (v.TContent.ConType == "四级标题")
            //                {
            //                    v.TContent.ConContent = string.Format(style.NumberFormat, titleB[d++]) + v.TContent.ConContent;
            //                }
            //                word.AddContent(v.TContent.ConContent, style);
            //            }

            //            #region 附件中的附件
            //            foreach (var waf in wa.Appendixs)
            //            {
            //                style = config.FontStyles.Find(w => w.FontName == "附件中附件");
            //                word.AddContent("附", style);

            //                if (waf.Type == 0)
            //                {
            //                    var waa = (WordsAppendix)waf;

            //                    style = config.FontStyles.Find(w => w.FontName == waa.Timu.ConType);
            //                    word.AddContent(waa.Timu.ConContent, style);

            //                    style = config.FontStyles.Find(w => w.FontName == waa.TZhenwen.ConType);
            //                    word.AddContent(waa.TZhenwen.ConContent, style);

            //                    a = b = c = d = 0;

            //                    foreach (var waav in waa.Zhenwen)
            //                    {
            //                        style = config.FontStyles.Find(w => w.FontName == waav.TContent.ConType);
            //                        if (waav.TContent.ConType == "一级标题")
            //                        {
            //                            waav.TContent.ConContent = string.Format(style.NumberFormat, titleA[a++]) + waav.TContent.ConContent;
            //                        }
            //                        else if (waav.TContent.ConType == "二级标题")
            //                        {
            //                            waav.TContent.ConContent = string.Format(style.NumberFormat, titleA[b++]) + waav.TContent.ConContent;
            //                        }
            //                        else if (waav.TContent.ConType == "三级标题")
            //                        {
            //                            waav.TContent.ConContent = string.Format(style.NumberFormat, titleB[c++]) + waav.TContent.ConContent;
            //                        }
            //                        else if (waav.TContent.ConType == "四级标题")
            //                        {
            //                            waav.TContent.ConContent = string.Format(style.NumberFormat, titleB[d++]) + waav.TContent.ConContent;
            //                        }
            //                        word.AddContent(waav.TContent.ConContent, style);
            //                    }

            //                    foreach (var waff in wa.Appendixs)
            //                    {
            //                        style = config.FontStyles.Find(w => w.FontName == "附件中附件");
            //                        word.AddContent("附", style);

            //                        if (appendix is FileAppendix)
            //                        {
            //                            var fa = (FileAppendix)appendix;
            //                            style = config.FontStyles.Find(w => w.FontName == "附件题目");
            //                            word.AddContent(fa.Title, style);

            //                            if (fa.FilePath.EndsWith(".xls") || fa.FilePath.EndsWith(".xlsx"))
            //                            {
            //                                word.AddExcel(fa.FilePath);
            //                            }
            //                            else
            //                            {
            //                                word.AddImage(fa.FilePath);
            //                            }
            //                        }
            //                        else
            //                        {
            //                            var waa3 = (WordsAppendix)waf;

            //                            style = config.FontStyles.Find(w => w.FontName == waa3.Timu.ConType);
            //                            word.AddContent(waa3.Timu.ConContent, style);

            //                            style = config.FontStyles.Find(w => w.FontName == waa3.TZhenwen.ConType);
            //                            word.AddContent(waa3.TZhenwen.ConContent, style);

            //                            a = b = c = d = 0;

            //                            foreach (var waav3 in waa3.Zhenwen)
            //                            {
            //                                style = config.FontStyles.Find(w => w.FontName == waav3.TContent.ConType);
            //                                if (waav3.TContent.ConType == "一级标题")
            //                                {
            //                                    waav3.TContent.ConContent = string.Format(style.NumberFormat, titleA[a++]) + waav3.TContent.ConContent;
            //                                }
            //                                else if (waav3.TContent.ConType == "二级标题")
            //                                {
            //                                    waav3.TContent.ConContent = string.Format(style.NumberFormat, titleA[b++]) + waav3.TContent.ConContent;
            //                                }
            //                                else if (waav3.TContent.ConType == "三级标题")
            //                                {
            //                                    waav3.TContent.ConContent = string.Format(style.NumberFormat, titleB[c++]) + waav3.TContent.ConContent;
            //                                }
            //                                else if (waav3.TContent.ConType == "四级标题")
            //                                {
            //                                    waav3.TContent.ConContent = string.Format(style.NumberFormat, titleB[d++]) + waav3.TContent.ConContent;
            //                                }
            //                                word.AddContent(waav3.TContent.ConContent, style);
            //                            }

            //                            if (waa3.Appendixs != null)
            //                            {
            //                                foreach (var waff3 in waa3.Appendixs)
            //                                {
            //                                    style = config.FontStyles.Find(w => w.FontName == "附件中附件");
            //                                    word.AddContent("附", style);

            //                                    if (appendix is FileAppendix)
            //                                    {
            //                                        var fa = (FileAppendix)appendix;
            //                                        style = config.FontStyles.Find(w => w.FontName == "附件题目");
            //                                        word.AddContent(fa.Title, style);

            //                                        if (fa.FilePath.EndsWith(".xls") || fa.FilePath.EndsWith(".xlsx"))
            //                                        {
            //                                            word.AddExcel(fa.FilePath);
            //                                        }
            //                                        else
            //                                        {
            //                                            word.AddImage(fa.FilePath);
            //                                        }
            //                                    }
            //                                }
            //                            }
            //                        }
            //                    }
            //                }
            //                else
            //                {
            //                    var fa = (FileAppendix)appendix;
            //                    style = config.FontStyles.Find(w => w.FontName == "附件题目");
            //                    word.AddContent(fa.Title, style);

            //                    if (fa.FilePath.EndsWith(".xls") || fa.FilePath.EndsWith(".xlsx"))
            //                    {
            //                        word.AddExcel(fa.FilePath);
            //                    }
            //                    else
            //                    {
            //                        word.AddImage(fa.FilePath);
            //                    }
            //                }
            //            }
            //            #endregion

            //            #endregion
            //        }
            //        else
            //        {
            //            var fa = (FileAppendix)appendix;
            //            style = config.FontStyles.Find(w => w.FontName == "附件题目");
            //            word.AddContent(fa.Title, style);

            //            if (fa.FilePath.EndsWith(".xls") || fa.FilePath.EndsWith(".xlsx"))
            //            {
            //                word.AddExcel(fa.FilePath);
            //            }
            //            else
            //            {
            //                word.AddImage(fa.FilePath);
            //            }
            //        }
            //    }
            //}
            #endregion

            //结尾
            style = config.FontStyles.Find(w => w.FontName == n.Chaosong.ConType);
            word.AddTable(n.Chaosong.ConContent, style);
            
            word.InsertPageNumber("Rfight", true);
            using(SaveFileDialog file = new SaveFileDialog())
            {
                file.FileName = "公告.doc";
                file.Filter = @"Word文件|*.doc";
                if(file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    word.SaveAndClose(file.FileName);
                }
            }
            MessageBox.Show("保存成功");
            new Msg().ShowDialog();
        }

        private void AddAppendixs(List<Appendix> appendixs, int leve)
        {
            if (appendixs == null) return;

            word.NewPage();

            string stylename = leve == 0 ? "正文附件" : "附件中附件";
            string appendixname = leve == 0 ? "附件{0}" : "附";
            Style style;

            int index = 1;
            foreach (var appendix in appendixs)
            {
                style = config.FontStyles.Find(w => w.FontName == stylename);
                word.AddContent(string.Format(appendixname, (index++).ToString()), style);

                switch (appendix.Type)
                {
                    case 0:
                        AddAppendixTitles((WordsAppendix)appendix);
                        break;
                    case 1:
                        AddAppendixFiles((FileAppendix)appendix);
                        break;
                }
            }
        }        

        private void AddAppendixTitles(WordsAppendix wa)
        {
            Style style = config.FontStyles.Find(w => w.FontName == wa.Timu.ConType);
            word.AddContent(wa.Timu.ConContent, style);

            style = config.FontStyles.Find(w => w.FontName == wa.TZhenwen.ConType);
            word.AddContent(wa.TZhenwen.ConContent, style);

            int a, b, c, d;
            a = b = c = d = 0;

            if (wa.Zhenwen != null)
            {
                foreach (var v in wa.Zhenwen)
                {
                    style = config.FontStyles.Find(w => w.FontName == v.TContent.ConType);
                    if (v.TContent.ConType == "一级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleA[a++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "二级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleA[b++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "三级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleB[c++]) + v.TContent.ConContent;
                    }
                    else if (v.TContent.ConType == "四级标题")
                    {
                        v.TContent.ConContent = string.Format(style.NumberFormat, titleB[d++]) + v.TContent.ConContent;
                    }
                    word.AddContent(v.TContent.ConContent, style);
                }
            }

            if (wa.Appendixs != null)
                AddAppendixs(wa.Appendixs, 1);
        }

        private void AddAppendixFiles(FileAppendix fa)
        {
            var style = config.FontStyles.Find(w => w.FontName == "附件题目");
            word.AddContent(fa.Title, style);

            if (fa.FilePath.EndsWith(".xls") || fa.FilePath.EndsWith(".xlsx"))
            {
                word.AddExcel(fa.FilePath);
            }
            else
            {
                word.AddImage(fa.FilePath);
            }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            DataVerifier dv = new DataVerifier();
            dv.Check(n == null, "请先添加通知模板再添加附件");

            if (dv.Pass)
            {                
                var index = (int)ddlWord.SelectedValue;

                WordsAppendix words = null;
                FormWordsAppendix f;

                if (index >= 0)
                {
                    dv.Check(n.Appendixs[index].Type != 0, "修改附件时请选择对应的附件类型");
                    if (dv.Pass)
                        words = (WordsAppendix)n.Appendixs[index];
                }

                if (dv.Pass)
                {
                    f = new FormWordsAppendix(words);

                    if (f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        if (f.Appendix.Timu == null) return;

                        f.Appendix.Type = 0;
                        f.Appendix.Title = f.Appendix.Timu.ConContent;

                        if (index == -1)
                        {
                            if (n.Appendixs == null)
                                n.Appendixs = new List<Appendix>();

                            n.Appendixs.Add(f.Appendix);
                            dDLSourceBindingSource.List.Add(new DDLSource() { Index = n.Appendixs.Count - 1, Text = "【文字】" + f.Appendix.Title });
                        }
                        else
                        {
                            n.Appendixs[index] = f.Appendix;
                            ((DDLSource)dDLSourceBindingSource.Current).Text = "【文字】" + f.Appendix.Title;
                            BindDDL();
                        }
                    }
                }
            }

            dv.ShowMsgIfFailed();
        }

        private void btnAppendix_Click(object sender, EventArgs e)
        {
            DataVerifier dv = new DataVerifier();
            dv.Check(n == null, "请先添加通知模板再添加附件");

            if (dv.Pass)
            {
                var index = (int)ddlWord.SelectedValue;

                FileAppendix f = null;

                if (index == -1)
                {
                    f = new FileAppendix();
                }
                else
                {
                    dv.Check(n.Appendixs[index].Type != 1, "修改附件时请选择对应的附件类型");
                    if(dv.Pass)
                        f = (FileAppendix)n.Appendixs[index];
                }

                if (dv.Pass)
                {
                    using (OpenFileDialog file = new OpenFileDialog())
                    {
                        if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            if (index == -1)
                            {
                                f.Type = 1;
                                f.Title = file.SafeFileName.Substring(0, file.SafeFileName.LastIndexOf("."));
                                f.FileName = new Content() { ConType = "附件", ConContent = file.FileName };
                                f.FilePath = file.FileName;
                                if (n.Appendixs == null) n.Appendixs = new List<Appendix>();
                                n.Appendixs.Add(f);
                                DDLSource source = new DDLSource() { Index = n.Appendixs.Count - 1, Text = "【文件】" + f.Title };
                                dDLSourceBindingSource.List.Add(source);
                            }
                            else
                            {
                                n.Appendixs[index] = f;
                                ((DDLSource)dDLSourceBindingSource.Current).Text = "【文件】" + f.Title;
                                BindDDL();
                            }
                        }
                    }
                }
            }

            dv.ShowMsgIfFailed();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(ddlWord.SelectedIndex > 0)
            {
                var index = (int)ddlWord.SelectedValue;
                n.Appendixs.RemoveAt(index);
                dDLSourceBindingSource.RemoveCurrent();                
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(n !=null)
            {
                ExportWord(n);
            }            
        }
    }

    public class Content
    {
        /*
         内容类型分为正文、标题、文件
         * 1.正文只需要处理ConContent
         * 2.标题需要记录当前的索引，用于显示标题前的序号
         * 3.文件记录FileName,ConContent为文件路径
         */

        public string ConType { get; set; }

        public string ConContent { get; set; }

        ///// <summary>
        ///// 文件标题
        ///// </summary>
        //public string FileName { get; set; }

        /// <summary>
        /// 当前标题索引
        /// </summary>
        public int TitleIndex { get; set; }
    }

    public enum ConType
    {
        Content,
        Title
    }

    public class Title
    {
        public int ID { get; set; }
        public Content TContent { get; set; }

        public int Parent { get; set; }

        public List<Appendix> Appendixs { get; set; }
    }

    public class Notice
    {
        public Content Fawen { get; set; }

        public Content Timu { get; set; }

        public Content Bumen { get; set; }

        public Content BumenZhenwen { get; set; }

        public List<Title> Zhenwen { get; set; }

        public Content Luokuan { get; set; }

        public Content Riqi { get; set; }

        public Content Chaosong { get; set; }

        public Content Fasong { get; set; }

        public List<Appendix> Appendixs { get; set; }
    }

    public class WordsAppendix : Appendix
    {
        public Content Timu { get; set; }

        public Content TZhenwen { get; set; }

        public List<Title> Zhenwen { get; set; }

        public List<Appendix> Appendixs { get; set; }

        //public List<WordsAppendix> WAppendixs { get; set; }

        //public List<FileAppendix> FAppendixs { get; set; }
    }

    public class FileAppendix : Appendix
    {
        public Content FileName { get; set; }
        public string FilePath { get; set; }
    }

    public abstract class Appendix
    {
        /// <summary>
        /// 0:文字 1:文件
        /// </summary>
        public int Type;
        public string Title;
    }

    public class DDLSource
    {
        public string Text { get; set; }
        public int Index { get; set; }
    }
}