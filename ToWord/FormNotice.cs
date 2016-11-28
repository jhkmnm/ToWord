using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ToWord
{
    public partial class FormNotice : Form
    {
        private Notice _notice;
        private Title _title;
        private string[] stitle = {"正文", "一级标题", "二级标题", "三级标题", "四级标题" };
        private int index;
        private bool isadd;
        private int id;
        private TreeNode selectedNode;
        /*
         * 1. 添加: 选中节点->点击标题或正文按钮->输入框中输入文字->确定，不选节点将会在根节点下添加
         * 2. 修改: 选中节点->输入框中修改文字->确定
         * 3. 删除: 选中节点->删除按钮->确定
         */

        public Notice Notice{get{return _notice;}}

        public FormNotice(Notice notice)
        {
            InitializeComponent();

            _notice = notice;
            Init();
        }

        private void Init()
        {
            if (_notice == null || _notice.Fawen == null)
            {
                _notice = new Notice();
                txtFawen.Text = "〔"+ DateTime.Now.Year.ToString() +"〕";
                return;
            }

            if (_notice.Fawen != null)
                txtFawen.Text = _notice.Fawen.ConContent;

            if (_notice.Timu != null)
                txtTimu.Text = _notice.Timu.ConContent;

            if (_notice.Bumen != null)
                txtBumen.Text = _notice.Bumen.ConContent;

            if (_notice.BumenZhenwen != null)
                txtZhenwen.Text = _notice.BumenZhenwen.ConContent;

            if (_notice.Chaosong != null)
                txtChaosong.Text = _notice.Chaosong.ConContent;            

            if (_notice.Zhenwen != null)
            {
                InitTree(null);
                _notice.Zhenwen.ForEach(f => id = id == 0 ? f.ID : (f.ID > id ? f.ID : id));
                treeView1.ExpandAll();
            }
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
                nodeSource = _notice.Zhenwen.FindAll(a => a.Parent == ((Title)node.Tag).ID);
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
            foreach(TreeNode n in node)
            {
                if(_notice.Zhenwen == null)
                {
                    _notice.Zhenwen = new List<Title>();
                }

                _notice.Zhenwen.Add((Title)n.Tag);
                GetTree(n.Nodes);
            }
        }

        private void SaveNotice()
        {
            if (_notice.Fawen == null)
                _notice.Fawen = new Content() { ConType = "发文字号" };

            if (_notice.Timu == null)
                _notice.Timu = new Content() { ConType = "题目" };

            if (_notice.Bumen == null)
                _notice.Bumen = new Content() { ConType = "主送部门" };

            if (_notice.BumenZhenwen == null)
                _notice.BumenZhenwen = new Content() { ConType = "正文" };

            if (_notice.Chaosong == null)
                _notice.Chaosong = new Content { ConType = "抄送" };

            _notice.Fawen.ConContent = txtFawen.Text;
            _notice.Timu.ConContent = txtTimu.Text;
            _notice.Bumen.ConContent = txtBumen.Text;
            _notice.BumenZhenwen.ConContent = txtZhenwen.Text.Replace('[', '〔').Replace(']', '〕').Replace('【', '〔').Replace('】', '〕');
            _notice.Chaosong.ConContent = txtChaosong.Text;
            GetTree(treeView1.Nodes);
        }

        private void AddTreeNode()
        {
            TreeNode n = new TreeNode(){ Tag = _title, Text = _title.TContent.ConContent };
            if (treeView1.Nodes == null)
            {
                treeView1.Nodes.Add(n);
            }
            else if(treeView1.SelectedNode != null)
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
            foreach(TreeNode cn in oldN.Nodes)
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
            if(selectedNode != null)
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
            if(DelCheck())
            {
                DeleteTreeNode();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DataVerifier dv = new DataVerifier();
            dv.Check(_title == null, "请选择要添加标题还是正文!");

            if (dv.Pass)
            {
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

            dv.ShowMsgIfFailed();
        }

        private void btnBiaoti1_Click(object sender, EventArgs e)
        {
            isadd = true;            
            index = 1;
            SetTitle(++id);
        }

        private void SetTitle(int id)
        {
            DataVerifier dv = new DataVerifier();
            dv.Check(index != 1 && selectedNode == null, "添加队一级标题以外的内容时请先选择它的上级内容");

            if (dv.Pass)
            {
                _title = new Title() { ID = id, TContent = new Content() { ConType = stitle[index] }, Parent = index == 1 ? 0 : ((Title)selectedNode.Tag).ID };
                txtContent.Text = "";
                txtContent.Focus();
            }

            dv.ShowMsgIfFailed();
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
            SaveNotice();
        }

        private void FormNotice_FormClosed(object sender, FormClosedEventArgs e)
        {            
            this.DialogResult = DialogResult.OK;
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            index = 0;
            DataVerifier dv = new DataVerifier();
            dv.Check(selectedNode == null, "请选择上级节点来添加表格");
            if (dv.Pass)
            {
                using (OpenFileDialog file = new OpenFileDialog())
                {
                    if (file.ShowDialog() == DialogResult.OK)
                    {
                        var Appendixs = new List<Appendix>();
                        Appendixs.Add(
                                new FileAppendix()
                                {
                                    Type = 1,
                                    Title = file.SafeFileName.Substring(0, file.SafeFileName.LastIndexOf(".")),
                                    FileName = new Content() { ConType = "附件", ConContent = file.FileName },
                                    FilePath = file.FileName
                                }
                                );

                        Title title = new Title()
                        {
                            ID = ++id,
                            Parent = ((Title)selectedNode.Tag).ID,
                            TContent = new Content() { ConType = "正文表格", ConContent = file.SafeFileName.Substring(0, file.SafeFileName.LastIndexOf(".")) },
                            Appendixs = Appendixs
                        };

                        TreeNode n = new TreeNode() { Tag = title, Text = title.TContent.ConContent };
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
                }
            }
            dv.ShowMsgIfFailed();
        }
    }
}
