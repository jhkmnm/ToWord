namespace ToWord
{
    partial class FormWordsAppendix
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnDel = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtContent = new System.Windows.Forms.TextBox();
            this.button7 = new System.Windows.Forms.Button();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.btnTzhenwen = new System.Windows.Forms.Button();
            this.btnBiaoti4 = new System.Windows.Forms.Button();
            this.btnBiaoti3 = new System.Windows.Forms.Button();
            this.btnBiaoti2 = new System.Windows.Forms.Button();
            this.btnBiaoti1 = new System.Windows.Forms.Button();
            this.txtZhenwen = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTimu = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnAppendix = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label8
            // 
            this.label8.AutoEllipsis = true;
            this.label8.ForeColor = System.Drawing.Color.Red;
            this.label8.Location = new System.Drawing.Point(14, 636);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(210, 21);
            this.label8.TabIndex = 47;
            this.label8.Text = "3. 删除: 选中节点->删除按钮->确定";
            // 
            // label7
            // 
            this.label7.AutoEllipsis = true;
            this.label7.ForeColor = System.Drawing.Color.Red;
            this.label7.Location = new System.Drawing.Point(14, 604);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(210, 32);
            this.label7.TabIndex = 46;
            this.label7.Text = "2. 修改: 选中节点->输入框中修改文字->确定";
            // 
            // label6
            // 
            this.label6.AutoEllipsis = true;
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(14, 562);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(210, 43);
            this.label6.TabIndex = 45;
            this.label6.Text = "1. 添加: 选中节点->点击标题或正文按钮->输入框中输入文字->确定，不选节点将会在根节点下添加";
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(137, 191);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 44;
            this.btnDel.Text = "删除";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // label5
            // 
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(570, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(240, 28);
            this.label5.TabIndex = 43;
            this.label5.Text = "1. 【保存】按钮会保存当前整个界面的内容2. 【确定】按钮保存左边的内容";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(32, 191);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 42;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(12, 234);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.Size = new System.Drawing.Size(212, 317);
            this.txtContent.TabIndex = 41;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(784, 120);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 23);
            this.button7.TabIndex = 40;
            this.button7.Text = "保存";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.treeView1.HideSelection = false;
            this.treeView1.Location = new System.Drawing.Point(235, 164);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(660, 584);
            this.treeView1.TabIndex = 39;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // btnTzhenwen
            // 
            this.btnTzhenwen.Location = new System.Drawing.Point(513, 120);
            this.btnTzhenwen.Name = "btnTzhenwen";
            this.btnTzhenwen.Size = new System.Drawing.Size(75, 23);
            this.btnTzhenwen.TabIndex = 38;
            this.btnTzhenwen.Text = "正文";
            this.btnTzhenwen.UseVisualStyleBackColor = true;
            this.btnTzhenwen.Click += new System.EventHandler(this.btnTzhenwen_Click);
            // 
            // btnBiaoti4
            // 
            this.btnBiaoti4.Location = new System.Drawing.Point(402, 120);
            this.btnBiaoti4.Name = "btnBiaoti4";
            this.btnBiaoti4.Size = new System.Drawing.Size(75, 23);
            this.btnBiaoti4.TabIndex = 37;
            this.btnBiaoti4.Text = "四级标题";
            this.btnBiaoti4.UseVisualStyleBackColor = true;
            this.btnBiaoti4.Click += new System.EventHandler(this.btnBiaoti4_Click);
            // 
            // btnBiaoti3
            // 
            this.btnBiaoti3.Location = new System.Drawing.Point(288, 120);
            this.btnBiaoti3.Name = "btnBiaoti3";
            this.btnBiaoti3.Size = new System.Drawing.Size(75, 23);
            this.btnBiaoti3.TabIndex = 36;
            this.btnBiaoti3.Text = "三级标题";
            this.btnBiaoti3.UseVisualStyleBackColor = true;
            this.btnBiaoti3.Click += new System.EventHandler(this.btnBiaoti3_Click);
            // 
            // btnBiaoti2
            // 
            this.btnBiaoti2.Location = new System.Drawing.Point(181, 120);
            this.btnBiaoti2.Name = "btnBiaoti2";
            this.btnBiaoti2.Size = new System.Drawing.Size(75, 23);
            this.btnBiaoti2.TabIndex = 35;
            this.btnBiaoti2.Text = "二级标题";
            this.btnBiaoti2.UseVisualStyleBackColor = true;
            this.btnBiaoti2.Click += new System.EventHandler(this.btnBiaoti2_Click);
            // 
            // btnBiaoti1
            // 
            this.btnBiaoti1.Location = new System.Drawing.Point(74, 120);
            this.btnBiaoti1.Name = "btnBiaoti1";
            this.btnBiaoti1.Size = new System.Drawing.Size(75, 23);
            this.btnBiaoti1.TabIndex = 34;
            this.btnBiaoti1.Text = "一级标题";
            this.btnBiaoti1.UseVisualStyleBackColor = true;
            this.btnBiaoti1.Click += new System.EventHandler(this.btnBiaoti1_Click);
            // 
            // txtZhenwen
            // 
            this.txtZhenwen.Location = new System.Drawing.Point(49, 60);
            this.txtZhenwen.Multiline = true;
            this.txtZhenwen.Name = "txtZhenwen";
            this.txtZhenwen.Size = new System.Drawing.Size(846, 42);
            this.txtZhenwen.TabIndex = 33;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 63);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 12);
            this.label4.TabIndex = 32;
            this.label4.Text = "正文:";
            // 
            // txtTimu
            // 
            this.txtTimu.Location = new System.Drawing.Point(49, 12);
            this.txtTimu.Multiline = true;
            this.txtTimu.Name = "txtTimu";
            this.txtTimu.Size = new System.Drawing.Size(443, 42);
            this.txtTimu.TabIndex = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 29;
            this.label3.Text = "题目:";
            // 
            // btnAppendix
            // 
            this.btnAppendix.Location = new System.Drawing.Point(628, 120);
            this.btnAppendix.Name = "btnAppendix";
            this.btnAppendix.Size = new System.Drawing.Size(75, 23);
            this.btnAppendix.TabIndex = 48;
            this.btnAppendix.Text = "附件";
            this.btnAppendix.UseVisualStyleBackColor = true;
            this.btnAppendix.Click += new System.EventHandler(this.btnAppendix_Click);
            // 
            // FormWordsAppendix
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(910, 754);
            this.Controls.Add(this.btnAppendix);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtContent);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.btnTzhenwen);
            this.Controls.Add(this.btnBiaoti4);
            this.Controls.Add(this.btnBiaoti3);
            this.Controls.Add(this.btnBiaoti2);
            this.Controls.Add(this.btnBiaoti1);
            this.Controls.Add(this.txtZhenwen);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtTimu);
            this.Controls.Add(this.label3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormWordsAppendix";
            this.Text = "文字附件";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button btnTzhenwen;
        private System.Windows.Forms.Button btnBiaoti4;
        private System.Windows.Forms.Button btnBiaoti3;
        private System.Windows.Forms.Button btnBiaoti2;
        private System.Windows.Forms.Button btnBiaoti1;
        private System.Windows.Forms.TextBox txtZhenwen;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTimu;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnAppendix;
    }
}