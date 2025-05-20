namespace AutodeskCost
{
    public class DataObject
    {
        public class PrjData
        {
            public string id { get; set; } // 計畫編號
            public string name { get; set; } // 計畫簡稱
            public string managerName { get; set; } // 計畫主管
            public int managerId { get; set; } // 主管員編
            public int departmentId { get; set; } // 主辦部ID
            public string department { get; set; } // 主辦部
            public double percent { get; set; } // 人月
            public double drawing { get; set; } // 繪圖
            public double hardware { get; set; } // 硬體
            public double software { get; set; } // 軟體
            public double network { get; set; } // 網路維護
            public double consumables { get; set; } // 消耗品/其他
            public double rent { get; set; } // 月租/時數
            public double share { get; set; } // 各計畫分攤(耗材)
            public double total { get; set; } // 小計
            public double diskCost { get; set; } // 磁區費用
        }
        public class UserData
        {
            public int id { get; set; }
            public string name { get; set; }
            public bool leader { get; set; }
            public double drawing { get; set; } // 繪圖
            public double rent { get; set; } // 月租/時數
            public double hardware { get; set; } // 硬體
            public double software { get; set; } // 軟體
            public double network { get; set; } // 網路維護
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
        public class SharePrjCost
        {
            public string prjId { get; set; }
            public double cost { get; set; }
        }
    }
}