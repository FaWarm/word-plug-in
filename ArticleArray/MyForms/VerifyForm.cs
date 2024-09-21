using ArticleArray.Verify;
using System;
using System.IO;
using System.Windows.Forms;

namespace ArticleArray.MyForms
{
    public partial class VerifyForm : Form
    {
        private static string _folderPath;
        public VerifyForm(string folderPath)
        {
            InitializeComponent();
            _folderPath = folderPath;

            _folderPath = folderPath;
            var resourceName = _folderPath + "Licence";

            if (!Directory.Exists(_folderPath)) Directory.CreateDirectory(_folderPath);

            if (!File.Exists(resourceName))
            {
                var file = new FileStream(resourceName, FileMode.Create);
                file.Close();
            }
            File.SetAttributes(resourceName, File.GetAttributes(resourceName) | FileAttributes.Hidden);
        }

        private void BtnApply_Click(object sender, EventArgs e)//验证许可
        {
            string txt = TbxLicence.Text.ToString();
            var verClass = new VerifyLicence(_folderPath);
            int res = verClass.InputVerifyLicence(txt);
            if (res > 0)
            {
                verClass.WriteToLicence(txt);
                LblResult.Text = "验证成功";
                this.Close();
            }
            else { LblResult.Text = "验证失败"; }
        }

        private void VerifyForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }
    }
}
