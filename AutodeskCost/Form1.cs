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
                    DateTime timeStart = DateTime.Now; // 計時開始 取得目前時間
                    bool success = readFile.ReadExcel(workbook, prjNumberTB.Text, leaderId);
                    if (!success) { MessageBox.Show(readFile.errorInfo); }
                    else
                    {
                        List<UserData> userDatas = readFile.userDatas; // 部門電腦使用費月報
                        List<PrjData> prjInfos = readFile.prjDatas; // 計畫資訊
                        double shareCost = readFile.shareCost; // 分攤金額
                        List<UserData> shareCosts = userDatas.Where(x => x.drawing > 0 || x.hardware > 0 || x.software > 0 || x.network > 0).ToList();
                        if (shareCosts.Count > 0) { readFile.SharePrjCost(userDatas, prjInfos, prjNumberTB.Text); } // 各計劃分攤(耗材), 分配剩餘金額
                        List<PrjData> shareCostPrjs = prjInfos.Where(x => x.percent.Equals(1)).ToList();
                        foreach(PrjData shareCostPrj in shareCostPrjs) { shareCostPrj.share = shareCost / shareCostPrjs.Count; }
                        readFile.WriteExcel(prjInfos, prjNumberTB.Text); // 將整合費用寫入Excel檔中

                        DateTime timeEnd = DateTime.Now; // 計時結束 取得目前時間
                        TimeSpan totalTime = timeEnd - timeStart;
                        MessageBox.Show("已完成整合費用。" + "\n\n完成，耗時：" + totalTime.Minutes + " 分 " + totalTime.Seconds + " 秒。\n\n");
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
        private void cancelBtn_Click(object sender, EventArgs e) { Close(); }
    }
}
