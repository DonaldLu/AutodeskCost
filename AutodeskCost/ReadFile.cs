using Microsoft.Office.Interop.Excel;
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

        public List<PrjData> prjDatas = new List<PrjData>(); // 計畫資訊
        public List<UserData> userDatas = new List<UserData>(); // 部門電腦使用費月報
        public List<PrjData> getDiskInfos = new List<PrjData>(); // 磁區
        public List<UserData> useSoftInfos = new List<UserData>(); // Autodesk軟體使用計畫
        public double shareCost { get; set; } // 要Share的費用

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
        public bool ReadExcel(Excel.Workbook workbook, string prjNumber, int leaderId)
        {
            string errorSheetName = string.Empty;
            this.prjDatas = new List<PrjData>(); // 計畫資訊
            this.userDatas = new List<UserData>(); // 部門電腦使用費月報
            this.getDiskInfos = new List<PrjData>(); // 磁區
            this.useSoftInfos = new List<UserData>(); // Autodesk軟體使用計畫

            try
            {
                foreach (string sheetName in sheetNames)
                {
                    errorSheetName = sheetName;
                    Excel._Worksheet workSheet = workbook.Sheets[sheetName];
                    Excel.Range Range = workSheet.UsedRange;

                    int rows = Range.Rows.Count;
                    int cols = Range.Columns.Count;

                    if (sheetName.Equals("計畫資訊"))
                    {
                        prjDatas = GetPrjInfos(workSheet, rows);
                        List<PrjData> nullPrjDatas = prjDatas.Where(x => String.IsNullOrEmpty(x.id) || String.IsNullOrEmpty(x.name) || String.IsNullOrEmpty(x.managerName) || x.managerId.Equals(0) || x.departmentId.Equals(0)).ToList();
                        if(nullPrjDatas.Count > 0) { errorInfo = "【計畫資訊】資料有缺漏。"; return false; }
                    }
                    else if (sheetName.Equals("部門電腦使用費月報"))
                    {
                        (List<string>, List<UserData>) prjAndUserDatas = GetPrjAndUserDatas(workSheet, rows, cols, prjNumber, leaderId, prjDatas);
                        List<string> losePrjIds = prjAndUserDatas.Item1;
                        userDatas = prjAndUserDatas.Item2;
                        // 檢查是否有計畫編號缺漏
                        if (losePrjIds.Count > 0)
                        {
                            errorInfo = "【計畫資訊】缺少計畫編號：";
                            int i = 1;
                            foreach (string losePrjId in losePrjIds) { errorInfo += "\n" + i + ". " + losePrjId; i++; }
                            return false;
                        }
                    }
                    else if (sheetName.Equals("磁區"))
                    {
                        getDiskInfos = GetDiskInfo(workSheet, rows);
                        foreach (PrjData prjData in prjDatas)
                        {
                            if (getDiskInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault() != null)
                            {
                                if (!prjData.id.Equals(prjNumber)) // 部門費用扣除, 要平均
                                {
                                    prjData.diskCost = getDiskInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault().diskCost;
                                    prjData.consumables = prjData.total - prjData.diskCost;
                                }
                                else
                                {
                                    this.shareCost = getDiskInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault().diskCost; // 要Share的費用
                                }
                            }
                            else { prjData.consumables = prjData.total; }
                        }
                    }
                    // 取得使用者各軟體使用費用
                    else if (sheetName.Equals("auto cad") || sheetName.Equals("BDSP") || sheetName.Equals("sap") || sheetName.Equals("Rhino") || sheetName.Equals("Lumion"))
                    {
                        if (!GetUserSoftCost(workSheet, rows, sheetName, userDatas))
                        {
                            errorInfo = "【" + sheetName + "】尚有員工編號資訊錯誤：";
                            return false;
                        }
                    }
                    else if (sheetName.Equals("Autodesk軟體使用計畫"))
                    {
                        useSoftInfos = GetUseSoftInfos(workSheet, rows);
                        List<UserData> nullUserInfos = useSoftInfos.Where(x => String.IsNullOrEmpty(x.project1) && String.IsNullOrEmpty(x.project2) && String.IsNullOrEmpty(x.project3)).ToList();
                        if (nullUserInfos.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者未填寫計畫編號：";
                            int i = 1;
                            foreach (string userName in nullUserInfos.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                        nullUserInfos = useSoftInfos.Where(x => !Math.Round((x.percent1 + x.percent2 + x.percent3), 0, MidpointRounding.AwayFromZero).Equals(1.0)).ToList();
                        if (nullUserInfos.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者計畫比例未達100%：";
                            int i = 1;
                            foreach (string userName in nullUserInfos.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                        // 比對使用者使用軟體的狀態
                        foreach (UserData userData in userDatas)
                        {
                            UserData userSoftInfo = useSoftInfos.Where(x => x.id.Equals(userData.id)).FirstOrDefault();
                            if (userSoftInfo != null)
                            {
                                userData.project1 = userSoftInfo.project1;
                                userData.percent1 = userSoftInfo.percent1;
                                userData.project2 = userSoftInfo.project2;
                                userData.percent2 = userSoftInfo.percent2;
                                userData.project3 = userSoftInfo.project3;
                                userData.percent3 = userSoftInfo.percent3;
                            }
                        }
                        List<UserData> userPrjsAllNulls = userDatas.Where(x => String.IsNullOrEmpty(x.project1) && String.IsNullOrEmpty(x.project2) && String.IsNullOrEmpty(x.project3)).ToList();
                        if (userPrjsAllNulls.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者無對應到軟體使用計畫編號：";
                            int i = 1;
                            foreach (string userName in userPrjsAllNulls.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                        else
                        {
                            List<string> nullPrjIds = PrjCosts(userDatas, prjDatas); // 計算使用者各計畫所花費占比
                            if (nullPrjIds.Count > 0)
                            {
                                errorInfo = "【計畫資訊】缺少對應Autodesk軟體使用計畫編號：";
                                int i = 1;
                                foreach (string nullPrjId in nullPrjIds) { errorInfo += i + ". " + nullPrjId + "\n"; i++; }
                                return false;
                            }
                            else
                            {

                            }
                        }
                    }

                    // 清理記憶體
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    // 釋放COM對象的經驗法則, 單獨引用與釋放COM對象, 不要使用多"."釋放
                    Marshal.ReleaseComObject(Range);
                    Marshal.ReleaseComObject(workSheet);
                }
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
                string value = workSheet.Cells[i, 1].Value?.ToString() ?? "";
                if (!String.IsNullOrEmpty(value))
                {
                    try
                    {
                        PrjData prjData = new PrjData();
                        prjData.id = workSheet.Cells[i, 1].Value?.ToString() ?? "";
                        prjData.name = workSheet.Cells[i, 2].Value?.ToString() ?? "";
                        prjData.managerName = workSheet.Cells[i, 3].Value?.ToString() ?? "";
                        value = workSheet.Cells[i, 4].Value?.ToString() ?? "";
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
            }            
            return prjInfos;
        }
        // 部門電腦使用費月報
        private (List<string>, List<UserData>) GetPrjAndUserDatas(Excel._Worksheet workSheet, int rows, int cols, string prjNumber, int leaderId, List<PrjData> prjInfos)
        {
            List<string> losePrjIds = new List<string>();
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
                            double.TryParse(value, out total);
                            if (total > 0)
                            {
                                PrjData prjInfo = prjInfos.Where(x => x.id.Equals(lastPrjId)).FirstOrDefault();
                                if(prjInfo != null)
                                {
                                    for (int col = 5; col <= 9; col++)
                                    {
                                        value = workSheet.Cells[i, col].Value?.ToString() ?? "";
                                        if (!String.IsNullOrEmpty(value))
                                        {
                                            double cost = 0;
                                            double.TryParse(value, out cost);
                                            if (col.Equals(5)) { prjInfo.drawing += cost; } // 繪圖
                                            else if (col.Equals(7)) { prjInfo.hardware += cost; } // 硬體
                                            else if (col.Equals(8)) { prjInfo.software += cost; } // 軟體
                                            else if (col.Equals(9)) { prjInfo.network += cost; } // 網路維護
                                        }
                                    }
                                    prjInfo.total = total; // 計畫合計
                                }
                                else { losePrjIds.Add(lastPrjId); }
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
                                    for(int col = 5; col <= 9; col++)
                                    {
                                        value = workSheet.Cells[i, col].Value?.ToString() ?? "";
                                        if (!String.IsNullOrEmpty(value))
                                        {
                                            double cost = 0;
                                            double.TryParse(value, out cost);
                                            if (col.Equals(5)) { userData.drawing = cost; } // 繪圖
                                            else if (col.Equals(7)) { userData.hardware = cost; } // 硬體
                                            else if (col.Equals(8)) { userData.software = cost; } // 軟體
                                            else if (col.Equals(9)) { userData.network = cost; } // 網路維護
                                        }
                                    }
                                    if (id.Equals(leaderId))
                                    {
                                        userData.leader = true;
                                        userData.project1 = prjNumber;
                                    }
                                    userDatas.Add(userData);
                                }
                                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                            }
                        }
                    }
                }
            }
            return (losePrjIds, userDatas);
        }
        // Autodesk軟體使用計畫
        private List<UserData> GetUseSoftInfos(Excel._Worksheet workSheet, int rows)
        {
            List<UserData> useSoftInfos = new List<UserData>();
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
                        useSoftInfos.Add(userData);
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
            return useSoftInfos;
        }
        // 磁區
        private List<PrjData> GetDiskInfo(Excel._Worksheet workSheet, int rows)
        {
            List<PrjData> getDiskInfos = new List<PrjData>();
            for (int i = 2; i <= rows; i++)
            {
                string value = workSheet.Cells[i, 2].Value?.ToString() ?? "";
                if (!String.IsNullOrEmpty(value))
                {
                    try
                    {
                        PrjData prjData = new PrjData();
                        prjData.id = value; // 計畫編號
                        value = workSheet.Cells[i, 3].Value?.ToString() ?? ""; // 主管
                        if (!String.IsNullOrEmpty(value))
                        {
                            string value1 = value.Substring(0, 4);
                            int managerId = Convert.ToInt32(value1);
                            string managerName = value.Substring(4, value.Length - 4);
                            prjData.managerId = managerId;
                            prjData.managerName = managerName;
                        }
                        value = workSheet.Cells[i, 6].Value?.ToString() ?? ""; // 歸屬部門
                        if (!String.IsNullOrEmpty(value))
                        {
                            prjData.department = value;
                        }
                        value = workSheet.Cells[i, 7].Value?.ToString() ?? ""; // 費用
                        if (!String.IsNullOrEmpty(value))
                        {
                            double cost = 0;
                            double.TryParse(value, out cost);
                            prjData.diskCost = cost;
                        }
                        getDiskInfos.Add(prjData);
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
            }
            return getDiskInfos;
        }
        // 取得使用者各軟體使用費用
        private bool GetUserSoftCost(Excel._Worksheet workSheet, int rows, string sheetName, List<UserData> userDatas)
        {
            for (int i = 2; i <= rows; i++)
            {
                string value = workSheet.Cells[i, 1].Value?.ToString() ?? "";
                if (!String.IsNullOrEmpty(value))
                {
                    int id = Convert.ToInt32(value);
                    UserData userData = userDatas.Where(x => x.id.Equals(id)).FirstOrDefault();
                    value = workSheet.Cells[i, 12].Value?.ToString() ?? "";
                    if (userData != null && !String.IsNullOrEmpty(value))
                    {
                        try
                        {
                            double cost = Convert.ToDouble(value);
                            if (sheetName.Equals("auto cad")) { userData.cadCost += cost; userData.total += cost; }
                            else if (sheetName.Equals("BDSP")) { userData.bdspCost += cost; userData.total += cost; }
                            else if (sheetName.Equals("sap")) { userData.sapCost += cost; userData.total += cost; }
                            else if (sheetName.Equals("Rhino")) { userData.rhinoCost += cost; userData.total += cost; }
                            else if (sheetName.Equals("Lumion")) { userData.lumionCost += cost; userData.total += cost; }
                        }
                        catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); return false; }
                    }
                }
                else { return false; }
            }
            return true;
        }
        // 計算使用者各計畫所花費占比
        private List<string> PrjCosts(List<UserData> userDatas, List<PrjData> prjInfos)
        {
            List<string> nullPrjIds = new List<string>();
            // 計算使用者各計畫所花費占比
            foreach (UserData userData in userDatas)
            {
                userData.cost1 = userData.total * userData.percent1;
                userData.cost2 = userData.total * userData.percent2;
                userData.cost3 = userData.total * userData.percent3;
            }
            // 各專案使用的月租/時數費用
            List<string> prjIds = new List<string>();
            prjIds = userDatas.Where(x => !String.IsNullOrEmpty(x.project1)).Select(x => x.project1.ToUpper()).Distinct().ToList().Concat
                    (userDatas.Where(x => !String.IsNullOrEmpty(x.project2)).Select(x => x.project2.ToUpper()).Distinct().ToList()).Concat
                    (userDatas.Where(x => !String.IsNullOrEmpty(x.project3)).Select(x => x.project3.ToUpper()).Distinct().ToList()).Distinct().OrderBy(x => x).ToList();
            foreach (string prjId in prjIds)
            {
                PrjData prjInfo = prjInfos.Where(x => x.id.Equals(prjId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault(); // 忽略大小寫
                if (prjInfo == null) { nullPrjIds.Add(prjId); }
                else
                {
                    List<UserData> samePrj1 = userDatas.Where(x => !String.IsNullOrEmpty(x.project1)).Where(x => x.project1.Equals(prjId, StringComparison.OrdinalIgnoreCase)).ToList();
                    List<UserData> samePrj2 = userDatas.Where(x => !String.IsNullOrEmpty(x.project2)).Where(x => x.project2.Equals(prjId, StringComparison.OrdinalIgnoreCase)).ToList();
                    List<UserData> samePrj3 = userDatas.Where(x => !String.IsNullOrEmpty(x.project3)).Where(x => x.project3.Equals(prjId, StringComparison.OrdinalIgnoreCase)).ToList();
                    double rent = samePrj1.Sum(x => x.cost1) + samePrj2.Sum(x => x.cost2) + samePrj3.Sum(x => x.cost3);
                    prjInfo.rent = rent;
                }
            }
            return nullPrjIds;
        }
        // 各計劃分攤(耗材), 分配剩餘金額
        public List<SharePrjCost> SharePrjCost(List<UserData> userDatas)
        {
            List<UserData> sharePrjCosts = userDatas.Where(x => x.drawing > 0 || x.hardware > 0 || x.software > 0 || x.network > 0).ToList();
            ShareCostForm shareCostForm = new ShareCostForm(sharePrjCosts);
            shareCostForm.ShowDialog();
            return shareCostForm.sharePrjCosts;
        }
        // 將整合費用寫入Excel檔中
        public void WriteExcel(List<PrjData> prjInfos, string prjNumber)
        {
            List<string> colNames = new List<string>() { "計畫編號", "計畫簡稱", "消耗品/其他", "月租/時數", "各計劃分攤(耗材)", "小計", "負責人", "員工編號", "磁區費用/月" };
            string excelPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Test";
            Excel.Application excelApp = new Excel.Application(); // 創建Excel
            //excelApp.Visible = true; // 開啟Excel可見
            Workbook workbook = excelApp.Workbooks.Add(); // 創建一個空的workbook
            Sheets sheets = workbook.Sheets; // 獲取當前工作簿的數量
            int sheetCount = 1;
            string fileName = "整合費用";
            List<string> existingNames = workbook.Worksheets.Cast<Worksheet>().Select(x => x.Name).ToList();
            Worksheet worksheet = sheets[1];
            try
            {
                if (sheetCount == 1) { if (!existingNames.Contains(fileName)) { worksheet.Name = fileName; } }
                else
                {
                    worksheet = sheets.Add(After: sheets[sheets.Count]); // 新增一個工作表
                    try
                    {
                        if (!existingNames.Contains(fileName)) { worksheet.Name = fileName; }
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
                sheetCount++;

                worksheet.Cells.Font.Name = "微軟正黑體"; // 設定Excel資料字體字型
                worksheet.Cells.Font.Size = 10; // 設定Excel資料字體大小
                for (int col = 0; col < colNames.Count; col++)
                {
                    excelApp.Cells[1, col + 1] = colNames[col];
                }
                prjInfos = prjInfos.Where(x => x.id.Equals(prjNumber) || x.consumables > 0 || x.rent > 0 || x.share > 0 || x.total > 0 || x.diskCost > 0).ToList();
                for (int i = 0; i < prjInfos.Count; i++)
                {
                    excelApp.Cells[i + 2, 1] = prjInfos[i].id; // 計畫編號
                    excelApp.Cells[i + 2, 2] = prjInfos[i].name; // 計畫簡稱
                    excelApp.Cells[i + 2, 3] = Math.Round(prjInfos[i].consumables, 2, MidpointRounding.AwayFromZero); // 消耗品/其他
                    excelApp.Cells[i + 2, 4] = Math.Round(prjInfos[i].rent, 2, MidpointRounding.AwayFromZero); // 月租/時數
                    excelApp.Cells[i + 2, 5] = Math.Round(prjInfos[i].share, 2, MidpointRounding.AwayFromZero); // 各計劃分攤(耗材)
                    double total = prjInfos[i].consumables + prjInfos[i].rent + prjInfos[i].share;
                    excelApp.Cells[i + 2, 6] = Math.Round(total, 2, MidpointRounding.AwayFromZero); // 小計
                    excelApp.Cells[i + 2, 7] = prjInfos[i].managerName; // 負責人
                    excelApp.Cells[i + 2, 8] = prjInfos[i].managerId; // 員工編號
                    excelApp.Cells[i + 2, 9] = Math.Round(prjInfos[i].diskCost, 2, MidpointRounding.AwayFromZero); // 磁區費用/月
                }
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            ReleaseObject(worksheet);

            workbook.SaveAs(excelPath);

            // 關閉工作簿和ExcelApp
            workbook.Close();
            excelApp.Quit();

            // 釋放COM
            ReleaseObject(sheets);
            ReleaseObject(workbook);
            ReleaseObject(excelApp);
        }
        // 釋放COM
        static void ReleaseObject(object obj)
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null; }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); obj = null; }
            finally { GC.Collect(); }
        }
    }
}