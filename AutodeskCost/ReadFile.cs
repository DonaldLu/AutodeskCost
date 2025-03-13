using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static AutodeskCost.DataObject;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutodeskCost
{
    public class ReadFile
    {
        public string filePath = string.Empty; // 選擇檔案路徑
        public bool trueOrFalse = false; // 預設未選取檔案
        public List<string> sheetNames = new List<string>() { "計畫資訊", "部門電腦使用費月報", "磁區", "auto cad", "BDSP", "sap", "Rhino", "Lumion", "Autodesk軟體使用計畫" };

        /// <summary>
        /// 選擇來源檔案
        /// </summary>
        /// <returns></returns>
        public void ChooseFiles()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "請選擇檔案";
            ofd.InitialDirectory = ".\\";
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls|" +
                         "Word Files (*.docx)|*.docx|Word Files (*.doc)|*.doc|" +
                         "All Files (*.*)|*.*";
            ofd.Multiselect = false; // 多選檔案
            this.filePath = string.Empty; // 選擇檔案路徑
            this.trueOrFalse = false; // 預設未選取檔案
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.filePath = ofd.FileName;
                this.trueOrFalse = true;
            }
            else { this.trueOrFalse = false; }
        }
        /// <summary>
        /// 選取Excel比對數量差異
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="charsToRemove"></param>
        /// <returns></returns>
        public List<ExcelContent> ReadExcel(string filePath, string prjNumber)
        {
            List<ExcelContent> excelContentList = new List<ExcelContent>();
            string errorSheetName = string.Empty;
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                foreach (string sheetName in sheetNames)
                {
                    errorSheetName = sheetName;
                    Excel._Worksheet workSheet = workbook.Sheets[sheetName];
                    Excel.Range Range = workSheet.UsedRange;

                    int rows = Range.Rows.Count;
                    int cols = Range.Columns.Count;

                    if (sheetName.Equals("計畫資訊"))
                    {

                    }
                    else if (sheetName.Equals("部門電腦使用費月報"))
                    {
                        (List<PrjData>, List<UserData>) prjAndUserDatas = GetPrjAndUserDatas(workSheet, rows, cols, prjNumber);
                        List<PrjData> prjDatas = prjAndUserDatas.Item1;
                        List<UserData> userDatas = prjAndUserDatas.Item2;
                    }

                    // 清理記憶體
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    // 釋放COM對象的經驗法則, 單獨引用與釋放COM對象, 不要使用多"."釋放
                    Marshal.ReleaseComObject(Range);
                    Marshal.ReleaseComObject(workSheet);
                }
                // 關閉與釋放
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch(Exception ex) { MessageBox.Show("【" + errorSheetName + "】資料讀取錯誤, 請檢查。\n" + ex.Message + "\n" + ex.ToString()); }

            return excelContentList;
        }
        /// <summary>
        /// 讀取標頭順序
        /// </summary>
        /// <param name="colCount"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private TitalNames SaveTitleNames(int colCount, Excel._Worksheet workSheet)
        {
            TitalNames titleNames = new TitalNames();
            for (int i = 1; i <= colCount; i++)
            {
                string titleName = string.Empty;
                try
                {
                    titleName = workSheet.Cells[1, i].Value;
                }
                catch { continue; }
                switch (titleName)
                {
                    case "標誌類型":
                        titleNames.type = i;
                        break;
                    case "標誌分類":
                        titleNames.category = i;
                        break;
                    case "標籤形狀":
                        titleNames.shape = i;
                        break;
                    case "參考圖號":
                        titleNames.detailDrawing = i; // 細部圖
                        titleNames.layout = i + 1; // 版面(設計準則)
                        break;
                    case "放指標":
                        titleNames.target = i;
                        break;
                    case "放標籤":
                        titleNames.tag = i;
                        break;
                    case "放說明\n(設計人員自行檢核)":
                        titleNames.explain = i;
                        break;
                    case "數量":
                        titleNames.count = i;
                        break;
                }
            }
            return titleNames;
        }
        /// <summary>
        /// 儲存Excel資料
        /// </summary>
        /// <param name="excelCompare"></param>
        /// <param name="titleNames"></param>
        /// <param name="workSheet"></param>
        /// <param name="charsToRemove"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private ExcelContent SaveExcelValue(ExcelContent excelCompare, TitalNames titleNames, Excel._Worksheet workSheet, List<string> charsToRemove, int i)
        {
            try
            {
                //excelCompare.code = workSheet.Cells[i, titleNames.code].Value; // 代碼
                //if (excelCompare.code == null)
                //{
                //    excelCompare.code = "";
                //}
                //excelCompare.classification = workSheet.Cells[i, titleNames.classification].Value; // 區域
                //excelCompare.level = workSheet.Cells[i, titleNames.level].Value; // 樓層
                //// 名稱(設定)
                //string editName = workSheet.Cells[i, titleNames.name].Value;
                //foreach (string c in charsToRemove)
                //{
                //    try
                //    {
                //        editName = editName.Replace(c, string.Empty); // 空間名稱(中文)
                //    }
                //    catch (Exception ex)
                //    {
                //        string error = ex.Message + "\n" + ex.ToString();
                //    }
                //}
                //excelCompare.name = editName;
                //try
                //{
                //    excelCompare.engName = workSheet.Cells[i, titleNames.engName].Value; // 空間名稱(英文)
                //}
                //catch (Exception)
                //{
                //    excelCompare.engName = "";
                //}
            }
            catch (Exception)
            {

            }

            return excelCompare;
        }
        // 部門電腦使用費月報
        private (List<PrjData>, List<UserData>) GetPrjAndUserDatas(Excel._Worksheet workSheet, int rows, int cols, string prjNumber)
        {
            List<PrjData> prjDatas = new List<PrjData>();
            List<UserData> userDatas = new List<UserData>();
            string lastPrjId = string.Empty;
            for (int i = 2; i <= rows; i++)
            {
                if(i == 108) { }
                string prjId = workSheet.Cells[i, 1].Value?.ToString() ?? "";
                if (!prjId.Contains("-"))
                {
                    if (!String.IsNullOrEmpty(prjId)) { lastPrjId = prjId; }
                    if (!prjId.Equals(prjNumber) && !lastPrjId.Equals(prjNumber))
                    {
                        if (!String.IsNullOrEmpty(prjId)) { lastPrjId = prjId; }
                        string value = workSheet.Cells[i, cols].Value?.ToString() ?? "";
                        if (!String.IsNullOrEmpty(value))
                        {
                            double total = 0;
                            double.TryParse(workSheet.Cells[i, cols].Value2.ToString(), out total);
                            if (total > 0)
                            {
                                PrjData prjData = new PrjData();
                                prjData.id = lastPrjId;
                                prjData.consumables = total;
                                prjDatas.Add(prjData);
                            }
                        }
                    }
                    else 
                    {
                        if (!String.IsNullOrEmpty(prjId) && !lastPrjId.Equals(prjNumber)) { lastPrjId = prjId; }
                        else
                        {
                            string value = workSheet.Cells[i, 4].Value?.ToString() ?? "";
                            if (!String.IsNullOrEmpty(value))
                            {
                                try
                                {
                                    string value1 = value.Substring(0, 4);
                                    int id = Convert.ToInt32(value1);
                                    string name = value.Substring(4, value.Length - 4);
                                    UserData userData = new UserData();
                                    userData.id = id;
                                    userData.name = name;
                                    userDatas.Add(userData);
                                }
                                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                            }
                        }
                    }
                }
            }
            return (prjDatas, userDatas);
        }
    }
}