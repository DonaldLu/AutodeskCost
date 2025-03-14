using System;
using System.Collections.Generic;
using System.Linq;
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
        public string errorInfo = string.Empty; // 錯誤訊息

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
        public bool ReadExcel(string filePath, string prjNumber)
        {
            string errorSheetName = string.Empty;
            List<PrjData> prjInfos = new List<PrjData>(); // 計畫資訊
            List<PrjData> prjDatas = new List<PrjData>(); // 部門電腦使用費月報
            List<UserData> userDatas = new List<UserData>(); // 部門電腦使用費月報
            List<UserData> userInfos = new List<UserData>(); // Autodesk軟體使用計畫

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
                        prjInfos = GetPrjInfos(workSheet, rows);
                        List<PrjData> nullPrjDatas = prjInfos.Where(x => String.IsNullOrEmpty(x.id) || String.IsNullOrEmpty(x.name) || String.IsNullOrEmpty(x.managerName) || x.managerId.Equals(0) || x.departmentId.Equals(0)).ToList();
                        if(nullPrjDatas.Count > 0) { errorInfo = "【計畫資訊】資料有缺漏。"; return false; }
                    }
                    else if (sheetName.Equals("部門電腦使用費月報"))
                    {
                        (List<PrjData>, List<UserData>) prjAndUserDatas = GetPrjAndUserDatas(workSheet, rows, cols, prjNumber);
                        prjDatas = prjAndUserDatas.Item1;
                        userDatas = prjAndUserDatas.Item2;
                        // 檢查是否有計畫編號缺漏
                        List<string> userNames = new List<string>();
                        foreach(PrjData prjData in prjDatas)
                        {
                            PrjData prj = prjInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault();
                            if (prj == null) { userNames.Add(prjData.id); }
                        }
                        if (userNames.Count > 0)
                        {
                            errorInfo = "【計畫資訊】缺少計畫編號：";
                            int i = 1;
                            foreach (string userName in userNames) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                    }
                    else if (sheetName.Equals("Autodesk軟體使用計畫"))
                    {
                        userInfos = GetUserInfos(workSheet, rows);
                        List<UserData> nullUserInfos = userInfos.Where(x => String.IsNullOrEmpty(x.project1) && String.IsNullOrEmpty(x.project2) && String.IsNullOrEmpty(x.project3)).ToList();
                        if (nullUserInfos.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者未填寫計畫編號：";
                            int i = 1;
                            foreach(string userName in nullUserInfos.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                        nullUserInfos = userInfos.Where(x => !Math.Round((x.percent1 + x.percent2 + x.percent3), 0, MidpointRounding.AwayFromZero).Equals(1.0)).ToList();
                        UserData user = userInfos.Where(x => x.name.Equals("郭天慈")).FirstOrDefault();
                        double percent = Math.Round(user.percent1 + user.percent2 + user.percent3, 0, MidpointRounding.AwayFromZero);
                        if (nullUserInfos.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者計畫比例未達100%：";
                            int i = 1;
                            foreach (string userName in nullUserInfos.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
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

            return true;
        }
        // 計畫資訊
        private List<PrjData> GetPrjInfos(Excel._Worksheet workSheet, int rows)
        {
            List<PrjData> prjInfos = new List<PrjData>();
            for (int i = 2; i <= rows; i++)
            {
                try
                {
                    PrjData prjData = new PrjData();
                    prjData.id = workSheet.Cells[i, 1].Value?.ToString() ?? "";
                    prjData.name = workSheet.Cells[i, 2].Value?.ToString() ?? "";
                    prjData.managerName = workSheet.Cells[i, 3].Value?.ToString() ?? "";
                    string value = workSheet.Cells[i, 4].Value?.ToString() ?? "";
                    if (!String.IsNullOrEmpty(value))
                    {
                        int managerId = 0;
                        int.TryParse(value, out managerId);
                        prjData.managerId = managerId;
                    }
                    value = workSheet.Cells[i, 5].Value?.ToString() ?? "";
                    if (!String.IsNullOrEmpty(value))
                    {
                        int departmentId = 0;
                        int.TryParse(value, out departmentId);
                        prjData.departmentId = departmentId;
                    }
                    value = workSheet.Cells[i, 6].Value?.ToString() ?? "";
                    if (!String.IsNullOrEmpty(value))
                    {
                        double percent = 0;
                        double.TryParse(value, out percent);
                        prjData.percent = percent;
                    }
                    prjInfos.Add(prjData);
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }            
            return prjInfos;
        }
        // 部門電腦使用費月報
        private (List<PrjData>, List<UserData>) GetPrjAndUserDatas(Excel._Worksheet workSheet, int rows, int cols, string prjNumber)
        {
            List<PrjData> prjDatas = new List<PrjData>();
            List<UserData> userDatas = new List<UserData>();
            string lastPrjId = string.Empty;
            for (int i = 2; i <= rows; i++)
            {
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
        // Autodesk軟體使用計畫
        private List<UserData> GetUserInfos(Excel._Worksheet workSheet, int rows)
        {
            List<UserData> userInfos = new List<UserData>();
            for (int i = 3; i <= rows; i++)
            {
                try
                {
                    string value = workSheet.Cells[i, 8].Value?.ToString() ?? "";
                    if (!String.IsNullOrEmpty(value))
                    {
                        string value1 = value.Substring(0, 4);
                        int id = Convert.ToInt32(value1);
                        string name = value.Substring(4, value.Length - 4);
                        UserData userData = new UserData();
                        userData.id = id;
                        userData.name = name.Trim();
                        string project = workSheet.Cells[i, 2].Value?.ToString() ?? "";
                        if (!String.IsNullOrEmpty(project))
                        {
                            userData.project1 = project;
                            string isNullOrEmpty = workSheet.Cells[i, 3].Value?.ToString() ?? "";
                            if (!String.IsNullOrEmpty(isNullOrEmpty))
                            {
                                double percent = 0;
                                double.TryParse(isNullOrEmpty, out percent);
                                userData.percent1 = percent;
                            }
                        }
                        project = workSheet.Cells[i, 4].Value?.ToString() ?? "";
                        if (!String.IsNullOrEmpty(project))
                        {
                            userData.project2 = project;
                            string isNullOrEmpty = workSheet.Cells[i, 5].Value?.ToString() ?? "";
                            if (!String.IsNullOrEmpty(isNullOrEmpty))
                            {
                                double percent = 0;
                                double.TryParse(isNullOrEmpty, out percent);
                                userData.percent2 = percent;
                            }
                        }
                        project = workSheet.Cells[i, 6].Value?.ToString() ?? "";
                        if (!String.IsNullOrEmpty(project))
                        {
                            userData.project3 = project;
                            string isNullOrEmpty = workSheet.Cells[i, 7].Value?.ToString() ?? "";
                            if (!String.IsNullOrEmpty(isNullOrEmpty))
                            {
                                double percent = 0;
                                double.TryParse(isNullOrEmpty, out percent);
                                userData.percent3 = percent;
                            }
                        }
                        userInfos.Add(userData);
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
            return userInfos;
        }
    }
}