namespace ToWord
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnNotice = new System.Windows.Forms.Button();
            this.btnWord = new System.Windows.Forms.Button();
            this.btnAppendix = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.ddlWord = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnNotice
            // 
            this.btnNotice.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnNotice.Location = new System.Drawing.Point(235, 25);
            this.btnNotice.Name = "btnNotice";
            this.btnNotice.Size = new System.Drawing.Size(157, 84);
            this.btnNotice.TabIndex = 1;
            this.btnNotice.Text = "通知模板";
            this.btnNotice.UseVisualStyleBackColor = true;
            this.btnNotice.Click += new System.EventHandler(this.btnNotice_Click);
            // 
            // btnWord
            // 
            this.btnWord.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnWord.Location = new System.Drawing.Point(132, 150);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(157, 60);
            this.btnWord.TabIndex = 2;
            this.btnWord.Text = "文字附件";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // btnAppendix
            // 
            this.btnAppendix.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnAppendix.Location = new System.Drawing.Point(333, 150);
            this.btnAppendix.Name = "btnAppendix";
            this.btnAppendix.Size = new System.Drawing.Size(157, 60);
            this.btnAppendix.TabIndex = 3;
            this.btnAppendix.Text = "附件";
            this.btnAppendix.UseVisualStyleBackColor = true;
            this.btnAppendix.Click += new System.EventHandler(this.btnAppendix_Click);
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExport.Location = new System.Drawing.Point(235, 386);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(157, 84);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            // 
            // ddlWord
            // 
            this.ddlWord.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddlWord.FormattingEnabled = true;
            this.ddlWord.Location = new System.Drawing.Point(132, 124);
            this.ddlWord.Name = "ddlWord";
            this.ddlWord.Size = new System.Drawing.Size(358, 20);
            this.ddlWord.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(235, 229);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(157, 60);
            this.button2.TabIndex = 8;
            this.button2.Text = "删除附件";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 510);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.ddlWord);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnAppendix);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.btnNotice);
            this.Name = "Form1";
            this.Text = "公文";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnNotice;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Button btnAppendix;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ComboBox ddlWord;
        private System.Windows.Forms.Button button2;
    }
}

