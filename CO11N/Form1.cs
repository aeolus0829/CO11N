using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using SAP.Middleware.Connector;

using System.Data.SqlClient;
using System.Collections;

namespace CO11N
{
    public partial class Form1 : Form
    {
        string D_connIP, D_connUser, D_connPwd, D_rptNm, D_status, D_connClient, D_connLanguage, D_RFCgetOrderDetail, D_RFCconfirmCommit,D_connNum, D_connSID;
      
      
        public Form1()
        {
            sapReportPrms sapReportPrms = new sapReportPrms();
            string[] ALL = sapReportPrms.SQL();

            // 連線字串
            D_connIP = "192.168.0.15";
            D_connUser = "DDIC";
            D_connPwd = "Ubn3dx";
            // D_rptNm = ALL[3];
            D_status = ALL[4];
            D_connClient = "620";
            D_connLanguage = "ZF";
            D_RFCgetOrderDetail = "ZPPRFC006"; //讀取工單資料
            D_RFCconfirmCommit = "ZPPRFC005"; //送出報工結果
            D_connSID = "DEV";

            if (D_status == "False")
            {
                MessageBox.Show("目前程式停用中，請連絡資訊組");            
            }
            else {
                InitializeComponent();
            }
        }
     
       
       
          public class cboDataList
        {
            public string cbo_Name { get; set; }
            public string cbo_Value { get; set; }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
                      
            List<cboDataList> lis_DataList = new List<cboDataList>()
            {
                new cboDataList
                {
                    cbo_Name = "",
                    cbo_Value = ""                    
                },
                new cboDataList
                {
                    cbo_Name = "機器製損",
                    cbo_Value = "0001"                    
                },
                new cboDataList
                {
                    cbo_Name = "人為損耗",
                    cbo_Value = "0002"                    
                },
                  new cboDataList
                {
                    cbo_Name = "庫存差異",
                    cbo_Value = "0003"                    
                }
            };
            comboBox1.DataSource = lis_DataList;
            comboBox1.DisplayMember = "cbo_Name";
            comboBox1.ValueMember = "cbo_Value";
                        
            txtActivity4.ReadOnly = true;
            txtActiunit5.ReadOnly = true;
            txtActivity5.ReadOnly = true;
            txtActiunit4.ReadOnly = true;
            txtActivity6.ReadOnly = true;
            txtActiunit6.ReadOnly = true;

            //過帳日期
            dtpPostgdate.Format = DateTimePickerFormat.Custom;
            dtpPostgdate.CustomFormat = "yyyy/MM/dd";

        }

        private void btnSubmin_Click(object sender, EventArgs e)
        {
            string windowsAccount = Environment.UserName;      

            RfcConfigParameters rfcPar = new RfcConfigParameters();
            rfcPar.Add(RfcConfigParameters.Name, D_connSID);
            rfcPar.Add(RfcConfigParameters.AppServerHost,D_connIP);
            rfcPar.Add(RfcConfigParameters.Client,  D_connClient);
            rfcPar.Add(RfcConfigParameters.User,D_connUser);
            rfcPar.Add(RfcConfigParameters.Password, D_connPwd);
            rfcPar.Add(RfcConfigParameters.SystemNumber, "00");
            rfcPar.Add(RfcConfigParameters.Language,D_connLanguage);
            RfcDestination dest = RfcDestinationManager.GetDestination(rfcPar);
            RfcRepository rfcrep = dest.Repository;

            IRfcFunction myfun = null;
            //函數名稱
            myfun = rfcrep.CreateFunction(D_RFCconfirmCommit);
            //設置輸入參數
            //工單號碼
            myfun.SetValue("AUFNR", txtAufnr.Text);
            //作業
            myfun.SetValue("OPERATION", txtOperation.Text);
            //確認良品率
            myfun.SetValue("YIELD", txtYield.Text);
            //廢品
            myfun.SetValue("SCRAP", txtScrap.Text);
            //重工
            myfun.SetValue("REWORK", txtRework.Text);
            //差異原因
            myfun.SetValue("REASON", comboBox1.SelectedValue);
            //數量單位
            myfun.SetValue("QUANUNIT", txtQuanunit.Text);
            //整備
            myfun.SetValue("ACTIVITY1", txtActivity1.Text);
            //整備單位
            myfun.SetValue("ACTIUNIT1", txtActiunit1.Text);
            //機器
            myfun.SetValue("ACTIVITY2", txtActivity2.Text);
            //機器單位
            myfun.SetValue("ACTIUNIT2", txtActiunit2.Text);
            //人工
            myfun.SetValue("ACTIVITY3", txtActivity3.Text);
            //人工單位
            myfun.SetValue("ACTIUNIT3", txtActiunit3.Text);
            //製造費用-其他
            myfun.SetValue("ACTIVITY4", txtActivity4.Text);
            //製造費用-其他單位
            myfun.SetValue("ACTIUNIT4", txtActiunit4.Text);
            //製造費用-間接人工
            myfun.SetValue("ACTIVITY5", txtActivity5.Text);
            //製造費用-間接人工單位
            myfun.SetValue("ACTIUNIT5", txtActiunit5.Text);
            //製造費用-折舊
            myfun.SetValue("ACTIVITY6", txtActivity6.Text);
            //製造費用-折舊單位
            myfun.SetValue("ACTIUNIT6", txtActiunit6.Text);
            //過帳日期
            myfun.SetValue("POSTG_DATE", Convert.ToDateTime(dtpPostgdate.Value.Date).ToString("yyyyMMdd"));
            //開時執行日期
            myfun.SetValue("START_DATE", txtStart_Date.Text);
            //開始執行時間
            if(txtStart_Time.Text != "") 
            myfun.SetValue("START_TIME", txtStart_Time.Text + "00");
            //完成執行日期
            myfun.SetValue("FIN_DATE", txtEnd_Date.Text);
            //完成執行時間
            if (txtFin_Time.Text != "")
            myfun.SetValue("FIN_TIME", txtFin_Time.Text + "00");
            //休息時間
            myfun.SetValue("BREAK_TIME", txtBreak_Time.Text);
            //休息時間單位
            myfun.SetValue("BREAK_UNIT", txtBreak_Unit.Text);
            //確認內文
            myfun.SetValue("CONF_TEXT", txtConf_Text.Text);
            //外部確認者
            myfun.SetValue("EX_CREATED_BY", windowsAccount);


            // Call function.
            myfun.Invoke(dest);

            //回傳參數
            string type = myfun.GetValue("STYPE").ToString();
            string status = myfun.GetValue("STATUS").ToString();

            // Declare message title. 
            string title = "";
            switch (type)
            {
                //訊息類型︰S 成功，E 錯誤， W 警告﹐I 資訊﹐A 取消
                case "S": title = "成功"; break;
                case "E": title = "錯誤"; break;
                case "W": title = "警告"; break;
                case "I": title = "資訊"; break;
                case "A": title = "取消"; break;
            }

            //MessageBox.Show(status, title);
            if (MessageBox.Show(status, title) == DialogResult.OK)
            {
                btnClear.PerformClick();
            }
        }


        private void btnClear_Click(object sender, EventArgs e)
        {   
            this.Controls.Clear();
            this.InitializeComponent();
            tableLayoutPanel1.Visible = true;
            dataGridView1.Visible = true;
            lblQty.Visible = true;
            lblStatus.Visible = true;
            lblSoitme.Visible = true;
            lblEnddate.Visible = true;
            Form1_Load(null,null);
        }

        private void btnPO_Click(object sender, EventArgs e)
        {
            RfcConfigParameters rfcPar = new RfcConfigParameters();
            rfcPar.Add(RfcConfigParameters.Name, D_connSID);
            rfcPar.Add(RfcConfigParameters.AppServerHost, D_connIP);
            rfcPar.Add(RfcConfigParameters.Client,  D_connClient);
            rfcPar.Add(RfcConfigParameters.User, D_connUser);
            rfcPar.Add(RfcConfigParameters.Password, D_connPwd);
            rfcPar.Add(RfcConfigParameters.SystemNumber, D_connNum);
            rfcPar.Add(RfcConfigParameters.Language, D_connLanguage);
            RfcDestination dest = RfcDestinationManager.GetDestination(rfcPar);
            RfcRepository rfcrep = dest.Repository;
            IRfcFunction rfcFunc = null;

            //函數名稱
            rfcFunc = rfcrep.CreateFunction(D_RFCgetOrderDetail);
            //輸入參數：工單號碼
            rfcFunc.SetValue("P_AUFNR", txtAufnr.Text.ToString().Trim());
            // Call function.
            rfcFunc.Invoke(dest);
            //回傳內表
            IRfcTable ITAB = rfcFunc.GetTable("ITAB");
            DataTable dt = new DataTable();
            dt.Columns.Add("作業號碼");
            dt.Columns.Add("作業短文");
            dt.Columns.Add("報工數量");
            
            for (int i = 0; i <= ITAB.RowCount-1 ; i++)
            {
                DataRow dr = dt.NewRow();
                ITAB.CurrentIndex = i;
                dr["作業號碼"] = ITAB.GetString("VORNR").ToString();
                dr["作業短文"] = ITAB.GetString("LTXA1").ToString();
                dr["報工數量"] = ITAB.GetString("GMNGA").ToString();
                dt.Rows.Add(dr);
            }
            
            //GridData資料來源
            dataGridView1.DataSource = dt.DefaultView;
            dataGridView1.ReadOnly = true;

            //回傳參數
            string KDAUF = rfcFunc.GetValue("KDAUF").ToString().TrimStart('0');
            string KDPOS = rfcFunc.GetValue("KDPOS").ToString().TrimStart('0');
            string PSMNG = rfcFunc.GetValue("PSMNG").ToString().TrimEnd('0').TrimEnd('.');
            string DGLTS = rfcFunc.GetValue("DGLTS").ToString();
            string USER_LINE = rfcFunc.GetValue("USER_LINE").ToString();
            //lblQty
            lblQty.Visible = true;
            lblQty.Text = "工單數量：" + PSMNG;
            //lblStatus
            lblStatus.Visible = true;
            lblStatus.Text = "使用者自定狀態：" + USER_LINE;
            //lblSoitem
            lblSoitme.Visible = true;
            lblSoitme.Text = "銷售訂單/項次：" + KDAUF + " / " + KDPOS;
            //lblEndate
            lblEnddate.Visible = true;
            lblEnddate.Text = "工單排程結束日期：" + DGLTS;
        }

        private void txtActivity2_TextChanged(object sender, EventArgs e)
        {
            //作業6 = 機器
            txtActivity6.Text = txtActivity2.Text;
            txtActiunit6.Text = txtActiunit2.Text;
        }      
           
        private void txtActivity3_TextChanged(object sender, EventArgs e)
        {
            //作業4 = 人工
            txtActivity4.Text = txtActivity3.Text;
            txtActiunit5.Text = txtActiunit3.Text;
            //作業5 = 人工
            txtActivity5.Text = txtActivity3.Text;
            txtActiunit4.Text = txtActiunit3.Text;
        }      

        private void txtStart_Date_KeyPress(object sender, KeyPressEventArgs e)
        {
            // e.KeyChar == (Char)48 ~ 57 -----> 0~9
            // e.KeyChar == (Char)8 -----------> Backpace
            // e.KeyChar == (Char)13-----------> Enter
            if (e.KeyChar == (Char)48 || e.KeyChar == (Char)49 ||
               e.KeyChar == (Char)50 || e.KeyChar == (Char)51 ||
               e.KeyChar == (Char)52 || e.KeyChar == (Char)53 ||
               e.KeyChar == (Char)54 || e.KeyChar == (Char)55 ||
               e.KeyChar == (Char)56 || e.KeyChar == (Char)57 ||
               e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void txtFin_Date_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (Char)48 || e.KeyChar == (Char)49 ||
               e.KeyChar == (Char)50 || e.KeyChar == (Char)51 ||
               e.KeyChar == (Char)52 || e.KeyChar == (Char)53 ||
               e.KeyChar == (Char)54 || e.KeyChar == (Char)55 ||
               e.KeyChar == (Char)56 || e.KeyChar == (Char)57 ||
               e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void txtStart_Time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (Char)48 || e.KeyChar == (Char)49 ||
            e.KeyChar == (Char)50 || e.KeyChar == (Char)51 ||
            e.KeyChar == (Char)52 || e.KeyChar == (Char)53 ||
            e.KeyChar == (Char)54 || e.KeyChar == (Char)55 ||
            e.KeyChar == (Char)56 || e.KeyChar == (Char)57 ||
            e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void txtFin_Time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (Char)48 || e.KeyChar == (Char)49 ||
                e.KeyChar == (Char)50 || e.KeyChar == (Char)51 ||
                e.KeyChar == (Char)52 || e.KeyChar == (Char)53 ||
                e.KeyChar == (Char)54 || e.KeyChar == (Char)55 ||
                e.KeyChar == (Char)56 || e.KeyChar == (Char)57 ||
                e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
        private void txtAufnr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    btnPO_Click(sender, e);
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                }
                catch (Exception)
                {
                    MessageBox.Show("工單" + txtAufnr.Text + "不存在", "錯誤");
                }

            }
        }
        private void txtAufnr_Leave(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                btnPO_Click(sender, e);
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
            catch (Exception)
            {
                MessageBox.Show("工單" + txtAufnr.Text + "不存在", "錯誤");
            }
        }

        private void txtStrat_Date_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtStrat_Date_TextChanged(object sender, EventArgs e)
        {
            //txtFin_Date.Text = txtStrat_Date.Text;
        }

        //時間計算參數
        int start_year, fin_year;
        int start_date1, start_date2;
        int fin_date1, fin_date2;

        private void txtMachine_Calc(object sender, EventArgs e)
        {
            // 計算機器加工時間，單位秒數
            // for 加工組
            try {
                double machineTimeInSec, orderQty, machineTimeInMin;
                orderQty = Convert.ToInt32(txtYield.Text);
                machineTimeInSec = Convert.ToInt32(txtMachineTime.Text);
                machineTimeInMin = Math.Round((orderQty*machineTimeInSec)/60,0);
                txtActivity2.Text = machineTimeInMin.ToString();
            } catch {
                MessageBox.Show("只能輸入秒數，格式為整數", "錯誤");
            }
        }

        private void txtMachineTime_TextChanged(object sender, EventArgs e)
        {
            txtMachineTime.Text = "";
        }

        private void machineTime_reclac(object sender, EventArgs e)
        {
            // 機器工時不是0或空的
            // 人工工時就要扣除機器工時

            if (txtActivity2.Text != "0" || string.IsNullOrEmpty(txtActivity2.Text))
            {
                try
                {
                    double menTimeInMin, machineTimeInMin, subMachineTimeFromMenTime;
                    menTimeInMin = Convert.ToInt32(txtActivity3.Text);
                    machineTimeInMin = Convert.ToInt32(txtActivity2.Text);
                    subMachineTimeFromMenTime = menTimeInMin - machineTimeInMin;
                    if (subMachineTimeFromMenTime < 0) {
                        txtActivity3.Text = "1";
                    } else { 
                        txtActivity3.Text = subMachineTimeFromMenTime.ToString();
                    }
                }
                catch
                {
                    MessageBox.Show("時間不能為負數", "錯誤");
                }
            }
}

        int start_time1, start_time2;
        int fin_time1, fin_time2;
        int sec;
        int countday,counthours,countminutes;
        int count;
        int final;

        //時間計算
        private void btnCalcTime_Click(object sender, EventArgs e)
        {   
            if(txtStart_Date.Text.Length==0|txtEnd_Date.Text.Length==0)
            { 
              txtStart_Date.Text=DateTime.Now.ToString("yyyyMMdd");          
              txtEnd_Date.Text=DateTime.Now.ToString("yyyyMMdd");   
            }
           if (txtStart_Date.Text.Length == 4 | txtEnd_Date.Text.Length == 4)
           {   
               txtStart_Date.Text = DateTime.Now.ToString("yyyy") + txtStart_Date.Text;
               txtEnd_Date.Text = DateTime.Now.ToString("yyyy") + txtEnd_Date.Text;
           }
           if (txtStart_Date.Text.Length == 3 | txtEnd_Date.Text.Length == 3)
           {
               txtStart_Date.Text = DateTime.Now.ToString("yyyy") + "0" + txtStart_Date.Text;
               txtEnd_Date.Text = DateTime.Now.ToString("yyyy") + "0" + txtEnd_Date.Text;
           }
            if (txtStart_Date.Text.Length != 8 | txtEnd_Date.Text.Length != 8 |
                txtStart_Time.Text.Length != 4 | txtFin_Time.Text.Length != 4 )
            {
                MessageBox.Show("日期或時間請輸入完整格式！ 例: 日期20150105 ; 時間0800","錯誤");
            }
            else {
               
                start_year = Convert.ToUInt16(txtStart_Date.Text.Substring(0, 4));
                start_date1 = Convert.ToInt16(txtStart_Date.Text.Substring(4, 2));
                start_date2 = Convert.ToInt16(txtStart_Date.Text.Substring(6 ,2));

                fin_year = Convert.ToInt16(txtEnd_Date.Text.Substring(0, 4));
                fin_date1 = Convert.ToInt16(txtEnd_Date.Text.Substring(4, 2));
                fin_date2 = Convert.ToInt16(txtEnd_Date.Text.Substring(6, 2));

                start_time1 = Convert.ToInt16(txtStart_Time.Text.Substring(0, 2));
                start_time2 = Convert.ToInt16(txtStart_Time.Text.Substring(2));

                fin_time1 = Convert.ToInt16(txtFin_Time.Text.Substring(0, 2));
                fin_time2 = Convert.ToInt16(txtFin_Time.Text.Substring(2));

                //       起始時間              年           月           日            時           分      秒預設為0
                DateTime d1 = new DateTime(start_year, start_date1, start_date2, start_time1, start_time2, sec);
                //       結束時間
                DateTime d2 = new DateTime(fin_year, fin_date1, fin_date2, fin_time1, fin_time2, sec);
                // d3= 結束時間─起始時間
                TimeSpan d3 = d2.Subtract(d1);

                //d3算出來的結果會是{00年:00月:00天:00時:00分:00秒}
                //因此要個別取值出來再運算!
                countday = Convert.ToUInt16(d3.Days.ToString());
                counthours = Convert.ToUInt16(d3.Hours.ToString());
                countminutes = Convert.ToUInt16(d3.Minutes.ToString());

               //分鐘數運算
                count = ((countday * 24 ) + counthours)*60+ countminutes;

                //label22.Text =" 相差分鐘數:"+count +"分";
                
                //投入人工預設為1
                final = Convert.ToInt16(textBox1.Text) * count;
                txtActivity3.Text = Convert.ToString(final);      
            }

        }
        //按enter 執行計算
        private void  txtFin_Time_KeyDown(object sender, KeyEventArgs e)
        {

        }
        //按enter 執行計算
        private void txtFin_Date_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnCalcTime_Click(sender, e);
            }
        }
        //按enter 執行計算
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnCalcTime_Click(sender, e);
            }
        }

        

        private void buttime_Click_1(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                btnCalcTime_Click(sender, e);
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
            catch (Exception)
            {
                MessageBox.Show("請檢查輸入時間是否正確", "錯誤");
            }
        }
        //離開 txtBox1 後自動觸發：時間計算
        private void textBox1_Leave(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                btnCalcTime_Click(sender, e);
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
            catch (Exception)
            {
                MessageBox.Show("請檢查輸入時間是否正確", "錯誤");
            }
        }
     }
    }

