using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

namespace ToWord
{
    public partial class FormWordsAppendix : Form
    {
        private WordsAppendix _appendix;
        private Title _title;
        private string[] stitle = { "正文", "一级标题", "二级标题", "三级标题", "四级标题" };
        private int index;
        private bool isadd;
        private int id;
        private TreeNode selectedNode;
        public WordsAppendix Appendix { get { return _appendix; } }

        private int _formIndex = 0; 

        public FormWordsAppendix(WordsAppendix appendix, int formindex = 0)
        {
            InitializeComponent();

            _appendix = appendix;
            _formIndex = formindex;

            if (_formIndex == 1)
                btnWord.Visible = false;

            DDLSource source = new DDLSource { Index = -1, Text = "新增" };
            dDLSourceBindingSource.DataSource = source;

            BindDDL();

            Init();
        }

        private void Init()
        {
            if (_appendix == null)
            {
                _appendix = new WordsAppendix();
                return;
            }

            txtTimu.Text = _appendix.Timu.ConContent;
            txtZhenwen.Text = _appendix.TZhenwen.ConContent;

            if (_appendix.Zhenwen != null)
            {
                InitTree(null);
                _appendix.Zhenwen.ForEach(f => id = id == 0 ? f.ID : (f.ID > id ? f.ID : id));
                treeView1.ExpandAll();                
            }

            if(_appendix.Appendixs != null)
            {
                int i = 0;
                _appendix.Appendixs.ForEach(item => {
                    dDLSourceBindingSource.Add(new DDLSource() { Index = i++, Text = item.Title });
                });
            }
        }

        private void InitTree(TreeNode node)
        {
            List<Title> nodeSource;

            if (node == null)
            {
                nodeSource = _appendix.Zhenwen.FindAll(a => a.Parent == 0);
            }
            else
            {
                nodeSource = _appendix.Zhenwen.FindAll(a => a.Parent == ((Title)node.Tag).ID);
            }

            foreach (Title t in nodeSource)
            {
                TreeNode no = new TreeNode(t.TContent.ConContent) { Tag = t };

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

        private void GetTree(TreeNodeCollection node)
        {
            foreach (TreeNode n in node)
            {
                if (_appendix.Zhenwen == null)
                {
                    _appendix.Zhenwen = new List<Title>();
                }

                _appendix.Zhenwen.Add((Title)n.Tag);
                GetTree(n.Nodes);
            }
        }

        private void SaveAppendix()
        {            

            if (_appendix.Timu == null)
                _appendix.Timu = new Content() { ConType = "题目" };

            if (_appendix.TZhenwen == null)
                _appendix.TZhenwen = new Content() { ConType = "正文" };
                        
            _appendix.Timu.ConContent = txtTimu.Text;
            _appendix.TZhenwen.ConContent = txtZhenwen.Text;
            GetTree(treeView1.Nodes);
        }

        private void AddTreeNode()
        {
            TreeNode n = new TreeNode() { Tag = _title, Text = _title.TContent.ConContent };
            if (treeView1.Nodes == null)
            {
                treeView1.Nodes.Add(n);
            }
            else if (treeView1.SelectedNode != null)
            {
                treeView1.SelectedNode.Nodes.Add(n);
            }
            else
            {
                treeView1.Nodes.Add(n);
            }
            treeView1.ExpandAll();
        }

        private void DeleteTreeNode()
        {
            var oldN = treeView1.SelectedNode;
            foreach (TreeNode cn in oldN.Nodes)
            {
                ((Title)cn.Tag).Parent = ((Title)cn.Tag).Parent;
            }
            treeView1.Nodes.Remove(oldN);
        }

        private void UpTreeNode()
        {
            treeView1.SelectedNode.Tag = _title;
            treeView1.SelectedNode.Text = _title.TContent.ConContent;
        }

        private bool AddCheck()
        {
            DataVerifier dv = new DataVerifier();
            dv.Check(selectedNode == null && index > 1, "请按照 一级标题->二级标题->三级标题->四级标题 的顺序添加!");
            if (selectedNode != null)
            {
                int pindex = stitle.ToList().FindIndex(a => ((Title)selectedNode.Tag).TContent.ConType == a);
                dv.CheckIfBeforePass(pindex == 0 && index != 0, "请非正文后面添加下级标题!");
                dv.CheckIfBeforePass(index != 0 && pindex != 0 && index != pindex + 1, "请按照 一级标题->二级标题->三级标题->四级标题 的顺序添加!");
            }

            dv.ShowMsgIfFailed();
            return dv.Pass;
        }

        private bool DelCheck()
        {
            DataVerifier dv = new DataVerifier();

            dv.Check(selectedNode == null, "请选择要删除的节点");
            if (selectedNode != null)
            {
                if (selectedNode.Nodes.Count > 0)
                {
                    var child = selectedNode.Nodes[0];
                    int sindex = stitle.ToList().FindIndex(a => ((Title)selectedNode.Tag).TContent.ConType == a);
                    if (child != null)
                    {
                        int cindex = stitle.ToList().FindIndex(a => ((Title)child.Tag).TContent.ConType == a);
                        dv.CheckIfBeforePass(sindex > 0 && cindex > 0, "请按照 一级标题->二级标题->三级标题->四级标题 的顺序");
                    }
                }
            }

            dv.ShowMsgIfFailed();
            return dv.Pass;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            _title = (Title)e.Node.Tag;
            selectedNode = e.Node;
            isadd = false;
            txtContent.Text = _title.TContent.ConContent;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (DelCheck())
            {
                DeleteTreeNode();
                selectedNode = selectedNode.PrevNode;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrWhiteSpace(txtContent.Text))
                _title.TContent.ConContent = txtContent.Text;

            if (isadd)
            {
                if (AddCheck())
                {
                    AddTreeNode();
                }
            }
            else
            {
                UpTreeNode();
            }
        }

        private void btnBiaoti1_Click(object sender, EventArgs e)
        {
            isadd = true;
            //int _id = 1;
            //if (_title != null)
            //{
            //    _id = ++id;
            //}
            index = 1;
            SetTitle(++id);
        }

        private void SetTitle(int id)
        {
            if (index > 1 && selectedNode == null)
                return;

            _title = new Title() { ID = id, TContent = new Content() { ConType = stitle[index] }, Parent = index == 1 ? 0 : ((Title)selectedNode.Tag).ID };
            txtContent.Text = "";
            txtContent.Focus();
        }

        private void btnBiaoti2_Click(object sender, EventArgs e)
        {
            isadd = true;
            index = 2;
            SetTitle(++id);
        }

        private void btnBiaoti3_Click(object sender, EventArgs e)
        {
            isadd = true;
            index = 3;
            SetTitle(++id);
        }

        private void btnBiaoti4_Click(object sender, EventArgs e)
        {
            isadd = true;
            index = 4;
            SetTitle(++id);
        }

        private void btnTzhenwen_Click(object sender, EventArgs e)
        {
            isadd = true;
            index = 0;
            SetTitle(++id);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SaveAppendix();
        }

        private void FormNotice_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void btnAppendix_Click(object sender, EventArgs e)
        {
            var index = (int)ddlWord.SelectedValue;

            FileAppendix f;

            if (index == -1)
            {
                f = new FileAppendix();
            }
            else
            {
                f = (FileAppendix)_appendix.Appendixs[index];
            }

            using (OpenFileDialog file = new OpenFileDialog())
            {
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (index == -1)
                    {
                        if (_appendix.Appendixs == null)
                            _appendix.Appendixs = new List<ToWord.Appendix>();

                        f.Type = 1;
                        f.Title = file.SafeFileName.Substring(0, file.SafeFileName.LastIndexOf("."));
                        f.FileName = new Content() { ConType = "附件中附件", ConContent = file.FileName };
                        f.FilePath = file.FileName;
                        _appendix.Appendixs.Add(f);
                        DDLSource source = new DDLSource() { Index = _appendix.Appendixs.Count - 1, Text = "【文件】" + f.Title };
                        dDLSourceBindingSource.List.Add(source);
                    }
                    else
                    {
                        _appendix.Appendixs[index] = f;
                        ((DDLSource)dDLSourceBindingSource.Current).Text = "【文件】" + f.Title;
                        BindDDL();
                    }
                }
            }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            var index = (int)ddlWord.SelectedValue;

            WordsAppendix words = null;
            FormWordsAppendix f;

            if (index >= 0)
            {
                words = (WordsAppendix)_appendix.Appendixs[index];
            }
            f = new FormWordsAppendix(words, _formIndex + 1);

            if (f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (f.Appendix.Timu == null) return;

                f.Appendix.Type = 0;
                f.Appendix.Title = f.Appendix.Timu.ConContent;

                if (index == -1)
                {
                    if (_appendix.Appendixs == null)
                        _appendix.Appendixs = new List<ToWord.Appendix>();

                    dDLSourceBindingSource.List.Add(new DDLSource() { Index = _appendix.Appendixs.Count - 1, Text = "【文字】" + f.Appendix.Title });
                    _appendix.Appendixs.Add(f.Appendix);
                }
                else
                {
                    _appendix.Appendixs[index] = f.Appendix;
                    ((DDLSource)dDLSourceBindingSource.Current).Text = "【文字】" + f.Appendix.Title;
                    BindDDL();
                }
            }
        }

        private void BindDDL()
        {
            ddlWord.DataSource = null;
            ddlWord.DataSource = dDLSourceBindingSource;
            ddlWord.ValueMember = "Index";
            ddlWord.DisplayMember = "Text";
        }

        private void FormWordsAppendix_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
