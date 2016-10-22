using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToWord
{
    public partial class Form2 : Form
    {
        private Notice _notice;

        public Form2()
        {
            InitializeComponent();

            DDLSource source = new DDLSource() { Text = "新增", Index = 0 };
            comboBox1.ValueMember = "Index";
            comboBox1.DisplayMember = "Text";
            comboBox1.Items.Add(source);

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

        private void button1_Click(object sender, EventArgs e)
        {
            ((DDLSource)comboBox1.SelectedItem).Text = "基本面";
            ((DDLSource)comboBox1.SelectedItem).Index = 3;                             
        }
    }
}
