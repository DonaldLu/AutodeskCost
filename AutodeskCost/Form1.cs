using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static AutodeskCost.DataObject;
using Excel = Microsoft.Office.Interop.Excel;

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
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                try
                {
                    ReadFile readFile = new ReadFile();
                    int leaderId = Convert.ToInt32(leaderIdTB.Text);
                    bool success = readFile.ReadExcel(workbook, prjNumberTB.Text, leaderId);
                    if (!success) { MessageBox.Show(readFile.errorInfo); }
                    else
                    {
                        List<UserData> userDatas = readFile.userDatas; // 部門電腦使用費月報
                        List<PrjData> prjInfos = readFile.prjInfos; // 計畫資訊
                        if (readFile.PrjCosts(userDatas, prjInfos)) // 計算使用者各計畫所花費占比
                        {

                        }
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                // 關閉與釋放
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
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
