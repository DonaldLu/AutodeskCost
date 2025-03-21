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
                        List<SharePrjCost> sharePrjCosts = readFile.SharePrjCost(userDatas); // 各計劃分攤(耗材), 分配剩餘金額
                        List<PrjData> prjInfos = readFile.prjDatas; // 計畫資訊
                        double shareCost = readFile.shareCost;
                        foreach (SharePrjCost sharePrjCost in sharePrjCosts)
                        {
                            PrjData prjData = prjInfos.Where(x => x.id.Equals(sharePrjCost.prjId)).FirstOrDefault();
                            if(prjData != null)
                            {
                                if (prjData.id.Equals(prjNumberTB.Text)) { shareCost += sharePrjCost.cost; }
                                else { prjData.consumables += sharePrjCost.cost; }
                            }
                        }
                        List<PrjData> shareCostPrjs = prjInfos.Where(x => x.percent.Equals(1)).ToList();
                        foreach(PrjData shareCostPrj in shareCostPrjs) { shareCostPrj.share = shareCost / shareCostPrjs.Count; }
                        readFile.WriteExcel(prjInfos, prjNumberTB.Text); // 將整合費用寫入Excel檔中
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
