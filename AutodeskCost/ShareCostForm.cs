using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using static AutodeskCost.DataObject;

namespace AutodeskCost
{
    public partial class ShareCostForm : Form
    {
        public List<SharePrjCost> sharePrjCosts = new List<SharePrjCost>();

        public ShareCostForm(List<UserData> sharePrjCosts, string prjNumber)
        {
            InitializeComponent();
            DataObject dataObject = new DataObject();
            CreateSharePrjItems(sharePrjCosts, prjNumber);
            CenterToParent();
        }
        // 新增
        public void CreateSharePrjItems(List<UserData> sharePrjCosts, string prjNumber)
        {
            List<string> items = new List<string>();
            foreach(UserData userData in sharePrjCosts)
            {
                if (userData.drawing > 0) { items.Add(userData.id + "_" + userData.name + "_繪圖_" + userData.drawing + "元："); }
                if (userData.hardware > 0) { items.Add(userData.id + "_" + userData.name + "_硬體_" + userData.hardware + "元："); }
                if (userData.software > 0) { items.Add(userData.id + "_" + userData.name + "_軟體_" + userData.software + "元："); }
                if (userData.network > 0) { items.Add(userData.id + "_" + userData.name + "_網路維修_" + userData.network + "元："); }
            }
            Label[] labels = new Label[items.Count];
            ComboBox[] comboBoxs = new ComboBox[items.Count];
            for (int i = 0; i < items.Count; i++)
            {
                try
                {
                    int id = Convert.ToInt32(items[i].Split('_')[0]);
                    UserData userData = sharePrjCosts.Where(x => x.id.Equals(id)).FirstOrDefault();
                    List<string> prjNames = new List<string>() { userData.project1, userData.project2, userData.project3, prjNumber };
                    AddControl(labels, comboBoxs, items[i], i, prjNames);
                }
                catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        private void AddControl(Label[] labels, ComboBox[] comboBoxs, string content, int i, List<string> prjNames)
        {
            labels[i] = new Label();
            labels[i].Font = new Font("微軟正黑體", 10, FontStyle.Regular);
            labels[i].Text = content;
            labels[i].AutoSize = true;
            labels[i].Location = new System.Drawing.Point(5, 7 + i * 25);

            comboBoxs[i] = new ComboBox();
            comboBoxs[i].Font = new Font("微軟正黑體", 10, FontStyle.Regular);
            foreach (string prjName in prjNames)
            {
                if (!String.IsNullOrEmpty(prjName)){ comboBoxs[i].Items.Add(prjName); }
            }
            //comboBoxs[i].Items.Add("");
            comboBoxs[i].AutoSize = true;
            comboBoxs[i].Location = new System.Drawing.Point(200, 5 + i * 25);
            comboBoxs[i].SelectedIndex = comboBoxs[i].Items.Count - 1; // 預設選最後一個(9510Q)
            comboBoxs[i].DropDownStyle = ComboBoxStyle.DropDownList; // 只能選取列表內容

            Panel.Controls.Add(labels[i]);
            Panel.Controls.Add(comboBoxs[i]);
        }
        // 確定
        private void sureBtn_Click(object sender, EventArgs e)
        {
            sharePrjCosts = new List<SharePrjCost>();
            foreach (Control control in Panel.Controls)
            {
                if(control is ComboBox)
                {
                    ComboBox comboBox = control as ComboBox;
                    SharePrjCost sharePrjCost = new SharePrjCost();                    
                    string[] spends = control.AccessibilityObject.Name.Split('_');
                    string spend = spends[spends.Length - 1].Replace("元：", "");
                    int id = Convert.ToInt32(spends[0]);
                    double cost = Convert.ToDouble(spend);

                    sharePrjCost.id = id;
                    sharePrjCost.name = spends[1];
                    sharePrjCost.type = spends[2];
                    sharePrjCost.prjId = comboBox.Text;
                    sharePrjCost.cost = cost;
                    this.sharePrjCosts.Add(sharePrjCost);
                }
            }
            Close();
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
