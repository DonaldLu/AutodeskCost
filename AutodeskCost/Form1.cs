using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AutodeskCost
{
    public partial class Form1 : Form
    {
        public string filePath = string.Empty; // Autodesk費用統整路徑
        public Form1()
        {
            InitializeComponent();
            CenterToParent();
        }
        // 選擇檔案路徑
        private void excelBtn_Click(object sender, EventArgs e)
        {
            ReadFile readFile = new ReadFile();
            readFile.ChooseFiles();
            this.filePath = readFile.filePath;
            if (readFile.trueOrFalse) { label2.Text = Path.GetFileName(filePath); }            
        }
        // 確定
        private void sureBtn_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(filePath)) { MessageBox.Show("尚未選擇路徑。"); }
            else
            {
                try
                {
                    ReadFile readFile = new ReadFile();
                    int leaderId = Convert.ToInt32(leaderIdTB.Text);
                    bool success = readFile.ReadExcel(filePath, prjNumberTB.Text, leaderId);
                    if (!success) { MessageBox.Show(readFile.errorInfo); }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                Close();
            }
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
