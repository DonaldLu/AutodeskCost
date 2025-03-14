using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutodeskCost
{
    public class DataObject
    {
        public class TitalNames
        {
            public int type { get; set; } // 標誌類型
            public int category { get; set; } // 標誌分類
            public int shape { get; set; } // 標誌形狀
            public int refNumber { get; set; } // 參考圖號
            public int detailDrawing { get; set; } // 細部圖
            public int layout { get; set; } // 版面(設計準則)
            public int target { get; set; } // 放指標
            public int tag { get; set; } // 放標籤
            public int explain { get; set; } // 放說明(設計人員自行檢核)
            public int count { get; set; } // 數量
        }
        public class ExcelContent
        {
            public string type { get; set; } // 標誌類型
            public string category { get; set; } // 標誌分類
            public string shape { get; set; } // 標誌形狀
            public string refNumber { get; set; } // 參考圖號
            public string detailDrawing { get; set; } // 細部圖
            public string layout { get; set; } // 版面(設計準則)
            public string target { get; set; } // 放指標
            public string tag { get; set; } // 放標籤
            public string explain { get; set; } // 放說明(設計人員自行檢核)
            public string count { get; set; } // 數量
        }
        public class UserData
        {
            public int id { get; set; }
            public string name { get; set; }
            public double cadCost { get; set; }
            public double bdspCost { get; set; }
            public double sapCost { get; set; }
            public double rhinoCost { get; set; }
            public double lumionCost { get; set; }
            public string project1 { get; set; }
            public double percent1 { get; set; }
            public double cost1 { get; set; }
            public string project2 { get; set; }
            public double percent2 { get; set; }
            public double cost2 { get; set; }
            public string project3 { get; set; }
            public double percent3 { get; set; }
            public double cost3 { get; set; }
        }
        public class PrjData
        {
            public string id { get; set; }
            public string name { get; set; }
            public string managerName { get; set; }
            public int managerId { get; set; }
            public int departmentId { get; set; }
            public string department { get; set; }
            public double percent { get; set; }
            public double drawing { get; set; }
            public double rent { get; set; }
            public double consumables { get; set; }
        }
    }
}