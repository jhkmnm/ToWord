using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ToWord.Model;

namespace ToWord
{
    public partial class Form2 : Form
    {
        private Notice _notice;

        public Form2()
        {
            InitializeComponent();

            DDLSource source = new DDLSource() { Text = "新增", Index = 0 };
            //comboBox1.ValueMember = "Index";
            //comboBox1.DisplayMember = "Text";
            //comboBox1.Items.Add(source);
            dDLSourceBindingSource.DataSource = source;

            load1();
            load2();

            return;

            WordUtil word = new WordUtil();
            var style = new Model.Style { FontName = "抄送", Font = "方正仿宋_GBK", Size = 14f};
            word.AddTable("x x x x x x x x，x x x x x x x x。", style);
            word.SaveAndClose("D:\\1234.doc");

            _notice = new Notice();

            Title t = new Title() { ID = 1, Parent = 0, TContent = new Content(){ ConContent = "一级标题", ConType = "一级标题", TitleIndex = 0}};
            Title t1 = new Title() { ID = 2, Parent = 0, TContent = new Content() { ConContent = "一级标题", ConType = "一级标题", TitleIndex = 1 } };

            Title tt = new Title() { ID = 3, Parent = 1, TContent = new Content() { ConContent = "二级标题", ConType = "二级标题", TitleIndex = 0 } };
            Title tt1 = new Title() { ID = 4, Parent = 1, TContent = new Content() { ConContent = "二级标题", ConType = "二级标题", TitleIndex = 1 } };

            Title ttt = new Title() { ID = 5, Parent = 3, TContent = new Content() { ConContent = "三级标题", ConType = "三级标题", TitleIndex = 0 } };
            Title ttt1 = new Title() { ID = 6, Parent = 4, TContent = new Content() { ConContent = "三级标题", ConType = "三级标题", TitleIndex = 1 } };

            _notice.Zhenwen = new List<Title>();
            _notice.Zhenwen.Add(t);
            _notice.Zhenwen.Add(t1);
            _notice.Zhenwen.Add(tt);
            _notice.Zhenwen.Add(tt1);
            _notice.Zhenwen.Add(ttt);
            _notice.Zhenwen.Add(ttt1);

            InitTree(null);
        }

        private void InitTree(TreeNode node)
        {
            List<Title> nodeSource;

            if (node == null)
            {
                nodeSource = _notice.Zhenwen.FindAll(a => a.Parent == 0);
            }
            else
            {
                nodeSource = _notice.Zhenwen.FindAll(a => a.Parent == (int)node.Tag);
            }

            foreach (Title t in nodeSource)
            {
                TreeNode no = new TreeNode(t.TContent.ConContent) { Tag = t.ID };                

                if (node == null)
                {
                    treeView1.Nodes.Add(no);
                }
                else
                {
                    node.Nodes.Add(no);
                }

                InitTree(no);
            }
        }

        Config config = null;
        WordUtil word = null;        

        private void button1_Click(object sender, EventArgs e)
        {
            word = new WordUtil();
            var style = new Style { FontName = "抄送", Font = "方正仿宋_GBK", Size = 14f };
            word.AddTable("抄送抄送抄送抄送抄送", style);
            //word.AddExcel(@"D:\桌面文件\list(20151101--20151130).xls");
            //word.AddExcel(@"D:\桌面文件\重新实名.xlsx");

            //using (SaveFileDialog file = new SaveFileDialog())
            //{
            //    file.FileName = "公告.doc";
            //    file.Filter = @"Word文件|*.doc";
            //    if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //    {
            //        word.SaveAndClose(file.FileName);
            //    }
            //}                        
            //dDLSourceBindingSource.List.Add(new DDLSource() { Text = "基本面", Index = 1 });
            //((DDLSource)dDLSourceBindingSource.DataSource).Text = "基本面";
            //comboBox1.DataSource = null;
            //comboBox1.DataSource = dDLSourceBindingSource;
            //comboBox1.ValueMember = "Index";
            //comboBox1.DisplayMember = "Text";                          
        }

        private void load1()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://www.gongyejf.com/Common/verify/");
            request.Timeout = 20000;
            request.ServicePoint.ConnectionLimit = 100;
            request.ReadWriteTimeout = 30000;
            request.Method = "GET";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode != HttpStatusCode.OK)
                return;
            Stream resStream = response.GetResponseStream();
            this.pictureBox1.Image = new Bitmap(resStream);
        }

        private void load2()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://www.gongyejf.com/Common/verify/");
            request.Timeout = 20000;
            request.ServicePoint.ConnectionLimit = 100;
            request.ReadWriteTimeout = 30000;
            request.Method = "GET";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode != HttpStatusCode.OK)
                return;
            StreamReader reader = new StreamReader(response.GetResponseStream());
            this.label1.Text = reader.ReadLine();
        }
    }
}
