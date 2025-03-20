using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutodeskCost
{
    public class DataObject
    {
        public class UserData
        {
            public int id { get; set; }
            public string name { get; set; }
            public bool leader { get; set; }
            public double cadCost { get; set; }
            public double bdspCost { get; set; }
            public double sapCost { get; set; }
            public double rhinoCost { get; set; }
            public double lumionCost { get; set; }
            public double total { get; set; }
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
            public string id { get; set; } // 計畫編號
            public string name { get; set; } // 計畫簡稱
            public string managerName { get; set; } // 計畫主管
            public int managerId { get; set; } // 主管員編
            public int departmentId { get; set; } // 主辦部ID
            public string department { get; set; } // 主辦部
            public double percent { get; set; }
            public double drawing { get; set; } // 繪圖
            public double rent { get; set; } // 月租/時數
            public double consumables { get; set; } // 消耗品/其他
            public double diskCost { get; set; } // 磁區費用
            public double total { get; set; } // 計畫合計
        }
    }
}