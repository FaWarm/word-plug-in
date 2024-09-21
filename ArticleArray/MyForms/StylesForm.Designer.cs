namespace ArticleArray.MyForms
{
    partial class StylesForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BtnApply = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.CbxEnFontName = new System.Windows.Forms.ComboBox();
            this.BtnSetDefaultStyles = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.BtnSaveStyles = new System.Windows.Forms.Button();
            this.BtnReadStyles = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnApply
            // 
            this.BtnApply.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.BtnApply.Font = new System.Drawing.Font("微软雅黑", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnApply.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.BtnApply.Location = new System.Drawing.Point(1078, 28);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(265, 51);
            this.BtnApply.TabIndex = 0;
            this.BtnApply.Text = "应用样式";
            this.BtnApply.UseVisualStyleBackColor = false;
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.CbxEnFontName);
            this.panel1.Controls.Add(this.BtnSetDefaultStyles);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.BtnSaveStyles);
            this.panel1.Controls.Add(this.BtnReadStyles);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.BtnApply);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 524);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1394, 185);
            this.panel1.TabIndex = 3;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(549, 12);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 18);
            this.label8.TabIndex = 7;
            this.label8.Text = "英文字体";
            // 
            // CbxEnFontName
            // 
            this.CbxEnFontName.FormattingEnabled = true;
            this.CbxEnFontName.Location = new System.Drawing.Point(639, 6);
            this.CbxEnFontName.Name = "CbxEnFontName";
            this.CbxEnFontName.Size = new System.Drawing.Size(201, 26);
            this.CbxEnFontName.TabIndex = 6;
            // 
            // BtnSetDefaultStyles
            // 
            this.BtnSetDefaultStyles.Location = new System.Drawing.Point(871, 111);
            this.BtnSetDefaultStyles.Name = "BtnSetDefaultStyles";
            this.BtnSetDefaultStyles.Size = new System.Drawing.Size(183, 49);
            this.BtnSetDefaultStyles.TabIndex = 5;
            this.BtnSetDefaultStyles.Text = "设为默认模板";
            this.BtnSetDefaultStyles.UseVisualStyleBackColor = true;
            this.BtnSetDefaultStyles.Click += new System.EventHandler(this.BtnSetDefaultStyles_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label7.Font = new System.Drawing.Font("宋体", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label7.Location = new System.Drawing.Point(11, 11);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(450, 22);
            this.label7.TabIndex = 4;
            this.label7.Text = "双击字体/字体颜色/字体尺寸单元格可修改值";
            // 
            // BtnSaveStyles
            // 
            this.BtnSaveStyles.Location = new System.Drawing.Point(1220, 111);
            this.BtnSaveStyles.Name = "BtnSaveStyles";
            this.BtnSaveStyles.Size = new System.Drawing.Size(123, 49);
            this.BtnSaveStyles.TabIndex = 3;
            this.BtnSaveStyles.Text = "保存模板";
            this.BtnSaveStyles.UseVisualStyleBackColor = true;
            this.BtnSaveStyles.Click += new System.EventHandler(this.BtnSaveStyles_Click);
            // 
            // BtnReadStyles
            // 
            this.BtnReadStyles.Location = new System.Drawing.Point(1078, 111);
            this.BtnReadStyles.Name = "BtnReadStyles";
            this.BtnReadStyles.Size = new System.Drawing.Size(123, 49);
            this.BtnReadStyles.TabIndex = 3;
            this.BtnReadStyles.TabStop = false;
            this.BtnReadStyles.Text = "读取模板";
            this.BtnReadStyles.UseVisualStyleBackColor = true;
            this.BtnReadStyles.Click += new System.EventHandler(this.BtnReadStyles_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 158);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(116, 18);
            this.label5.TabIndex = 2;
            this.label5.Text = "段前段后行距";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 18);
            this.label3.TabIndex = 2;
            this.label3.Text = "行间距说明";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 18);
            this.label2.TabIndex = 2;
            this.label2.Text = "字体大小对照";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label6.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(134, 156);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(479, 20);
            this.label6.TabIndex = 1;
            this.label6.Text = "输入数字为行数，如0.5为0.5行（1行固定值为12磅）\r\n";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label4.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(134, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(569, 40);
            this.label4.TabIndex = 1;
            this.label4.Text = "数字 1——11 为行倍数，如1表示单倍行距，1.5表示1.5倍行距\r\n数字 >12 为固定值";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(134, 61);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(899, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "一号:26，小一:24，二号:22，小二:18，三号:16，小三:15，四号:14，小四:12，五号:10.5，小五:9";
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 30;
            this.dataGridView1.Size = new System.Drawing.Size(1394, 709);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.BindingContextChanged += new System.EventHandler(this.dataGridView1_BindingContextChanged);
            // 
            // StylesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1394, 709);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "StylesForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "快速样式设置";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.StylesForm_FormClosed);
            this.Load += new System.EventHandler(this.StylesForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnApply;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button BtnReadStyles;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button BtnSaveStyles;
        private System.Windows.Forms.Button BtnSetDefaultStyles;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox CbxEnFontName;
    }
}