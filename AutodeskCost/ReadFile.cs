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
        public List<string> sheetNames = new List<string>() { "計畫資訊", "部門電腦使用費月報", "Autodesk軟體使用計畫", "磁區", "auto cad", "BDSP", "sap", "Rhino", "Lumion" };
        public string errorInfo = string.Empty; // 錯誤訊息

        public List<PrjData> prjInfos = new List<PrjData>(); // 計畫資訊
        public List<PrjData> prjDatas = new List<PrjData>(); // 部門電腦使用費月報
        public List<UserData> userDatas = new List<UserData>(); // 部門電腦使用費月報
        public List<UserData> useSoftInfos = new List<UserData>(); // Autodesk軟體使用計畫
        public List<PrjData> getDiskInfos = new List<PrjData>(); // 磁區

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
            this.prjInfos = new List<PrjData>(); // 計畫資訊
            this.prjDatas = new List<PrjData>(); // 部門電腦使用費月報
            this.userDatas = new List<UserData>(); // 部門電腦使用費月報
            this.useSoftInfos = new List<UserData>(); // Autodesk軟體使用計畫
            this.getDiskInfos = new List<PrjData>(); // 磁區

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
                        prjInfos = GetPrjInfos(workSheet, rows);
                        List<PrjData> nullPrjDatas = prjInfos.Where(x => String.IsNullOrEmpty(x.id) || String.IsNullOrEmpty(x.name) || String.IsNullOrEmpty(x.managerName) || x.managerId.Equals(0) || x.departmentId.Equals(0)).ToList();
                        if(nullPrjDatas.Count > 0) { errorInfo = "【計畫資訊】資料有缺漏。"; return false; }
                    }
                    else if (sheetName.Equals("部門電腦使用費月報"))
                    {
                        (List<PrjData>, List<UserData>) prjAndUserDatas = GetPrjAndUserDatas(workSheet, rows, cols, prjNumber, leaderId);
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
                        useSoftInfos = GetUseSoftInfos(workSheet, rows);
                        List<UserData> nullUserInfos = useSoftInfos.Where(x => String.IsNullOrEmpty(x.project1) && String.IsNullOrEmpty(x.project2) && String.IsNullOrEmpty(x.project3)).ToList();
                        if (nullUserInfos.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者未填寫計畫編號：";
                            int i = 1;
                            foreach(string userName in nullUserInfos.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
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
                        SaveUserUseSoftInfos(userDatas, useSoftInfos); // 比對使用者使用軟體的狀態
                        List<UserData> userPrjsAllNulls = userDatas.Where(x => String.IsNullOrEmpty(x.project1) && String.IsNullOrEmpty(x.project2) && String.IsNullOrEmpty(x.project3)).ToList();
                        if(userPrjsAllNulls.Count > 0)
                        {
                            errorInfo = "【Autodesk軟體使用計畫】使用者無對應到軟體使用計畫編號：";
                            int i = 1;
                            foreach (string userName in userPrjsAllNulls.Select(x => x.name)) { errorInfo += "\n" + i + ". " + userName; i++; }
                            return false;
                        }
                    }
                    else if (sheetName.Equals("磁區"))
                    {
                        getDiskInfos = GetDiskInfo(workSheet, rows);
                        foreach(PrjData prjData in prjDatas)
                        {
                            if(getDiskInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault() != null)
                            {
                                prjData.diskCost = getDiskInfos.Where(x => x.id.Equals(prjData.id)).FirstOrDefault().diskCost;
                                prjData.consumables = prjData.total - prjData.diskCost;
                            }
                            else { prjData.consumables = prjData.total; }
                        }
                    }
                    // 取得使用者各軟體使用費用
                    else if (sheetName.Equals("auto cad") || sheetName.Equals("BDSP") || sheetName.Equals("sap") || sheetName.Equals("Rhino") || sheetName.Equals("Lumion"))
                    {
                        if(!GetUserSoftCost(workSheet, rows, sheetName, userDatas)) 
                        {
                            errorInfo = "【" + sheetName + "】尚有員工編號資訊錯誤：";
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
        private (List<PrjData>, List<UserData>) GetPrjAndUserDatas(Excel._Worksheet workSheet, int rows, int cols, string prjNumber, int leaderId)
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
                                prjData.total = total;
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
            return (prjDatas, userDatas);
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
        // 比對使用者使用軟體的狀態
        private void SaveUserUseSoftInfos(List<UserData> userDatas, List<UserData> useSoftInfos)
        {
            foreach(UserData userData in userDatas)
            {
                UserData userSoftInfo = useSoftInfos.Where(x => x.id.Equals(userData.id)).FirstOrDefault();
                if(userSoftInfo != null)
                {
                    userData.project1 = userSoftInfo.project1;
                    userData.percent1 = userSoftInfo.percent1;
                    userData.project2 = userSoftInfo.project2;
                    userData.percent2 = userSoftInfo.percent2;
                    userData.project3 = userSoftInfo.project3;
                    userData.percent3 = userSoftInfo.percent3;
                }
            }
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
        public bool PrjCosts(List<UserData> userDatas, List<PrjData> prjInfos)
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

                }
            }
            if (nullPrjIds.Count > 0)
            {
                string errorInfo = string.Empty;
                int i = 1;
                foreach (string nullPrjId in nullPrjIds) { errorInfo += i + ". " + nullPrjId + "\n"; i++; }
                MessageBox.Show("【計畫資訊】缺少對應Autodesk軟體使用計畫編號：\n" + errorInfo);
                return false;
            }
            return true;
        }
    }
}