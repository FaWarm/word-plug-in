namespace ArticleArray.MyForms
{
    partial class WidthAndHeightSetForm
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.CbxCenter = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnOk = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.TbxHeight = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TbxWidth = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.CbxCenter);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.BtnCancel);
            this.panel1.Controls.Add(this.BtnOk);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.TbxHeight);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.TbxWidth);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(369, 303);
            this.panel1.TabIndex = 5;
            // 
            // CbxCenter
            // 
            this.CbxCenter.AutoSize = true;
            this.CbxCenter.ForeColor = System.Drawing.Color.Black;
            this.CbxCenter.Location = new System.Drawing.Point(125, 158);
            this.CbxCenter.Name = "CbxCenter";
            this.CbxCenter.Size = new System.Drawing.Size(106, 22);
            this.CbxCenter.TabIndex = 5;
            this.CbxCenter.Text = "水平居中";
            this.CbxCenter.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(65, 112);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(224, 18);
            this.label5.TabIndex = 4;
            this.label5.Text = "只设置宽或高则按比例缩放";
            // 
            // BtnCancel
            // 
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(203, 203);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 44);
            this.BtnCancel.TabIndex = 3;
            this.BtnCancel.Text = "取消";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOk
            // 
            this.BtnOk.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnOk.Location = new System.Drawing.Point(68, 204);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 43);
            this.BtnOk.TabIndex = 2;
            this.BtnOk.Text = "确定";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(273, 73);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(26, 18);
            this.label4.TabIndex = 0;
            this.label4.Text = "cm";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(273, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 18);
            this.label3.TabIndex = 0;
            this.label3.Text = "cm";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "图片宽";
            // 
            // TbxHeight
            // 
            this.TbxHeight.Location = new System.Drawing.Point(154, 70);
            this.TbxHeight.Name = "TbxHeight";
            this.TbxHeight.Size = new System.Drawing.Size(100, 28);
            this.TbxHeight.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 18);
            this.label2.TabIndex = 0;
            this.label2.Text = "图片高";
            // 
            // TbxWidth
            // 
            this.TbxWidth.Location = new System.Drawing.Point(154, 21);
            this.TbxWidth.Name = "TbxWidth";
            this.TbxWidth.Size = new System.Drawing.Size(100, 28);
            this.TbxWidth.TabIndex = 1;
            // 
            // WidthAndHeightSetForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(403, 327);
            this.Controls.Add(this.panel1);
            this.Name = "WidthAndHeightSetForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "图片尺寸设置";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TbxHeight;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TbxWidth;
        private System.Windows.Forms.CheckBox CbxCenter;
    }
}