using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ArticleArray.Verify
{
    public class VerifyLicence
    {
        private string _folderPath;
        private string resourceName;
        public VerifyLicence(string folderPath)
        {
            _folderPath = folderPath;
            resourceName = _folderPath + "Licence";

            if (!Directory.Exists(_folderPath))
            {
                Directory.CreateDirectory(_folderPath);
            }

            if (!File.Exists(resourceName))
            {
                FileStream fs = new FileStream(resourceName, FileMode.Create);
                fs.Close();
                //string nowTime = DateTime.Today.ToString("yyyyM M d d");
                //string[] nums = nowTime.Split(' ');
                //string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghizklmnopqrstuvwxyz";
                //string lic = "ALMPN"+{}+"HAUUL"+0+"tQFJWT"+2+"CCUQR"+3+"VWSCQ"+0+"HALPU"+5+"HVYNU"+1+"SCESV"+1+"ZIKTM";
                //string lic = $"ALMPN{nums[0]}HAUUL{nums[1]}QFJWT{nums[2]}CCUQR{nums[3]}VWSCQ{nums[4]}HALPU{nums[5]}HVYNU{nums[6]}SCESV{nums[7]}ZIKTM";
                //File.WriteAllText(resourceName, lic);
            }
        }
        public string ReadResourceFile()
        {
            string content = File.ReadAllText(resourceName);
            return content;
        }
        public int UserVerifyLicence()
        {
            try
            {
                string licence = ReadResourceFile();
                string pattern = @"\d+";
                MatchCollection matches = Regex.Matches(licence, pattern);

                string concatenatedNumbers = "";
                foreach (Match match in matches)
                {
                    concatenatedNumbers += match.Value;
                }
                int lic = int.Parse(concatenatedNumbers);

                int today = int.Parse(DateTime.Today.ToString("yyyyMMdd"));//当前时间

                return lic - today;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        public int InputVerifyLicence(string licence)//输入验证码验证 >0通过验证
        {
            if (licence.Length != 53) return -1;
            string pattern = @"\d+";
            MatchCollection matches = Regex.Matches(licence, pattern);

            string concatenatedNumbers = "";
            foreach (Match match in matches)
            {
                concatenatedNumbers += match.Value;
            }
            int lic = int.Parse(concatenatedNumbers);

            int today = int.Parse(DateTime.Today.ToString("yyyyMMdd"));//当前时间
            return lic - today;
        }

        public void WriteToLicence(string licenseText)//写入文件
        {
            try
            {
                File.Delete(resourceName);
            }
            catch (Exception)
            {
            }
            finally
            {
                File.WriteAllText(resourceName, licenseText);
                File.SetAttributes(resourceName, File.GetAttributes(resourceName) | FileAttributes.Hidden);//设置文件隐藏
            }

        }

    }
}