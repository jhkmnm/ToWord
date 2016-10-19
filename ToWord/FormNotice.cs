﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SRDFAA.Fram.Utilities;

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
        }

        private void Init()
        {
            if (_notice == null) return;

            txtFawen.Text = _notice.Fawen.ConContent;
            txtTimu.Text = _notice.Timu.ConContent;
            txtBumen.Text = _notice.Bumen.ConContent;
            txtZhenwen.Text = _notice.BumenZhenwen.ConContent;

            if (_notice.Zhenwen != null)
            {
                InitTree(null);
                _notice.Zhenwen.ForEach(f => id = id == 0 ? f.ID : (f.ID > id ? f.ID : id));
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

            _notice.Fawen.ConContent = txtFawen.Text;
            _notice.Timu.ConContent = txtTimu.Text;
            _notice.Bumen.ConContent = txtBumen.Text;
            _notice.BumenZhenwen.ConContent = txtZhenwen.Text;
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
                selectedNode = null;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _title.TContent.ConContent = txtContent.Text;

            if(isadd)
            {
                if(AddCheck())
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
            int _id = 1;
            if(_title != null)
            {
                _id = ++id;
            }
            index = 1;
            SetTitle(_id);
        }

        private void SetTitle(int id)
        {
            _title = new Title() { ID = id, TContent = new Content() { ConType = stitle[index] }, Parent = 0 };
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
            SaveNotice();
        }

        private void FormNotice_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}