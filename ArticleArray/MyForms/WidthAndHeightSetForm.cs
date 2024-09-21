using System;
using System.Windows.Forms;

namespace ArticleArray.MyForms
{
    public partial class WidthAndHeightSetForm : Form
    {
        public float width;
        public float height;
        public bool isCenter;
        public WidthAndHeightSetForm()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (TbxWidth.Text != "") width = float.Parse(TbxWidth.Text);
            if (TbxHeight.Text != "") height = float.Parse(TbxHeight.Text);
            isCenter = CbxCenter.Checked;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}