namespace ArticleArray.MyForms
{
    partial class VerifyForm
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
            this.TbxLicence = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnApply = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.LblResult = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // TbxLicence
            // 
            this.TbxLicence.Location = new System.Drawing.Point(72, 99);
            this.TbxLicence.Margin = new System.Windows.Forms.Padding(5);
            this.TbxLicence.Name = "TbxLicence";
            this.TbxLicence.Size = new System.Drawing.Size(1108, 39);
            this.TbxLicence.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(77, 57);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(180, 28);
            this.label1.TabIndex = 1;
            this.label1.Text = "请输入许可码";
            // 
            // BtnApply
            // 
            this.BtnApply.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.BtnApply.Location = new System.Drawing.Point(254, 180);
            this.BtnApply.Margin = new System.Windows.Forms.Padding(5);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(213, 75);
            this.BtnApply.TabIndex = 2;
            this.BtnApply.Text = "确定";
            this.BtnApply.UseVisualStyleBackColor = false;
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(613, 180);
            this.BtnCancel.Margin = new System.Windows.Forms.Padding(5);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(219, 75);
            this.BtnCancel.TabIndex = 3;
            this.BtnCancel.Text = "取消";
            this.BtnCancel.UseVisualStyleBackColor = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(897, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 28);
            this.label2.TabIndex = 5;
            this.label2.Text = "验证结果：";
            // 
            // LblResult
            // 
            this.LblResult.AutoSize = true;
            this.LblResult.Location = new System.Drawing.Point(1045, 57);
            this.LblResult.Name = "LblResult";
            this.LblResult.Size = new System.Drawing.Size(0, 28);
            this.LblResult.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(77, 318);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(446, 28);
            this.label3.TabIndex = 7;
            this.label3.Text = "添加wx获取插件许可:yu2272896219";
            // 
            // VerifyForm
            // 
            this.AcceptButton = this.BtnApply;
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 28F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(1244, 405);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.LblResult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnApply);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TbxLicence);
            this.Font = new System.Drawing.Font("黑体", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "VerifyForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "输入许可码";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.VerifyForm_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbxLicence;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnApply;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label LblResult;
        private System.Windows.Forms.Label label3;
    }
}