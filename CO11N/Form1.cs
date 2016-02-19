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
        string D_connIP, D_connUser, D_connPwd, D_status, D_connClient, D_connLanguage, D_RFCgetOrderDetail, D_RFCconfirmCommit,D_connNum, D_connSID;
        bool keyIsAccept;

        public Form1()
        {
            sapReportPrms sapReportPrms = new sapReportPrms();
            string[] ALL = sapReportPrms.SQL();

            // 連線字串
            D_connIP = "192.168.0.15";
            D_connUser = "DDIC";
            D_connPwd = "Ubn3dx";
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

            RfcConfigParameters rfcPara = new RfcConfigParameters();
            rfcPara.Add(RfcConfigParameters.Name, D_connSID);
            rfcPara.Add(RfcConfigParameters.AppServerHost,D_connIP);
            rfcPara.Add(RfcConfigParameters.Client,  D_connClient);
            rfcPara.Add(RfcConfigParameters.User,D_connUser);
            rfcPara.Add(RfcConfigParameters.Password, D_connPwd);
            rfcPara.Add(RfcConfigParameters.SystemNumber, "00");
            rfcPara.Add(RfcConfigParameters.Language,D_connLanguage);
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcPara);
            RfcRepository rfcrep = rfcDest.Repository;

            IRfcFunction rfcFunc = null;
            //函數名稱
            rfcFunc = rfcrep.CreateFunction(D_RFCconfirmCommit);
            //設置輸入參數
            //工單號碼
            rfcFunc.SetValue("AUFNR", txtAufnr.Text);
            //作業
            rfcFunc.SetValue("OPERATION", txtOperation.Text);
            //確認良品率
            rfcFunc.SetValue("YIELD", txtYield.Text);
            //廢品
            rfcFunc.SetValue("SCRAP", txtScrap.Text);
            //重工
            rfcFunc.SetValue("REWORK", txtRework.Text);
            //差異原因
            rfcFunc.SetValue("REASON", comboBox1.SelectedValue);
            //數量單位
            rfcFunc.SetValue("QUANUNIT", txtQuanunit.Text);
            //整備
            rfcFunc.SetValue("ACTIVITY1", txtActivity1.Text);
            //整備單位
            rfcFunc.SetValue("ACTIUNIT1", txtActiunit1.Text);
            //機器
            rfcFunc.SetValue("ACTIVITY2", txtActivity2.Text);
            //機器單位
            rfcFunc.SetValue("ACTIUNIT2", txtActiunit2.Text);
            //人工
            rfcFunc.SetValue("ACTIVITY3", txtActivity3.Text);
            //人工單位
            rfcFunc.SetValue("ACTIUNIT3", txtActiunit3.Text);
            //製造費用-其他
            rfcFunc.SetValue("ACTIVITY4", txtActivity4.Text);
            //製造費用-其他單位
            rfcFunc.SetValue("ACTIUNIT4", txtActiunit4.Text);
            //製造費用-間接人工
            rfcFunc.SetValue("ACTIVITY5", txtActivity5.Text);
            //製造費用-間接人工單位
            rfcFunc.SetValue("ACTIUNIT5", txtActiunit5.Text);
            //製造費用-折舊
            rfcFunc.SetValue("ACTIVITY6", txtActivity6.Text);
            //製造費用-折舊單位
            rfcFunc.SetValue("ACTIUNIT6", txtActiunit6.Text);
            //過帳日期
            rfcFunc.SetValue("POSTG_DATE", Convert.ToDateTime(dtpPostgdate.Value.Date).ToString("yyyyMMdd"));
            //開時執行日期
            rfcFunc.SetValue("START_DATE", txtStart_Date.Text);
            //開始執行時間
            if(txtStart_Time.Text != "") 
            rfcFunc.SetValue("START_TIME", txtStart_Time.Text + "00");
            //完成執行日期
            rfcFunc.SetValue("FIN_DATE", txtEnd_Date.Text);
            //完成執行時間
            if (txtEnd_Time.Text != "")
            rfcFunc.SetValue("FIN_TIME", txtEnd_Time.Text + "00");
            //休息時間
            rfcFunc.SetValue("BREAK_TIME", txtBreak_Time.Text);
            //休息時間單位
            rfcFunc.SetValue("BREAK_UNIT", txtBreak_Unit.Text);
            //確認內文
            rfcFunc.SetValue("CONF_TEXT", txtConf_Text.Text);
            //外部確認者
            rfcFunc.SetValue("EX_CREATED_BY", windowsAccount);

            // Call function.
            rfcFunc.Invoke(rfcDest);

            //回傳參數
            string returnMessageType = rfcFunc.GetValue("STYPE").ToString();
            string returnStatus = rfcFunc.GetValue("STATUS").ToString();

            string returnMessage = "";
            switch (returnMessageType)
            {
                case "S": returnMessage = "成功"; break;
                case "E": returnMessage = "錯誤"; break;
                case "W": returnMessage = "警告"; break;
                case "I": returnMessage = "資訊"; break;
                case "A": returnMessage = "取消"; break;
            }

            if (MessageBox.Show(returnStatus, returnMessage) == DialogResult.OK)
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
            RfcConfigParameters rfcPara = new RfcConfigParameters();
            rfcPara.Add(RfcConfigParameters.Name, D_connSID);
            rfcPara.Add(RfcConfigParameters.AppServerHost, D_connIP);
            rfcPara.Add(RfcConfigParameters.Client,  D_connClient);
            rfcPara.Add(RfcConfigParameters.User, D_connUser);
            rfcPara.Add(RfcConfigParameters.Password, D_connPwd);
            rfcPara.Add(RfcConfigParameters.SystemNumber, D_connNum);
            rfcPara.Add(RfcConfigParameters.Language, D_connLanguage);
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcPara);
            RfcRepository rfcRepo = rfcDest.Repository;

            IRfcFunction rfcFunc = null;

            // rfc 函數名稱
            rfcFunc = rfcRepo.CreateFunction(D_RFCgetOrderDetail);
            //輸入參數：工單號碼
            rfcFunc.SetValue("P_AUFNR", txtAufnr.Text.ToString().Trim());
            // Call function.
            rfcFunc.Invoke(rfcDest);
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
            
            dataGridView1.DataSource = dt.DefaultView;
            dataGridView1.ReadOnly = true;

            // rfc 回傳參數
            string KDAUF = rfcFunc.GetValue("KDAUF").ToString().TrimStart('0');
            string KDPOS = rfcFunc.GetValue("KDPOS").ToString().TrimStart('0');
            string PSMNG = rfcFunc.GetValue("PSMNG").ToString().TrimEnd('0').TrimEnd('.');
            string DGLTS = rfcFunc.GetValue("DGLTS").ToString();
            string USER_LINE = rfcFunc.GetValue("USER_LINE").ToString();

            lblQty.Text = "工單數量：" + PSMNG;
            lblStatus.Text = "使用者自定狀態：" + USER_LINE;
            lblSoitme.Text = "銷售訂單/項次：" + KDAUF + " / " + KDPOS;
            lblEnddate.Text = "工單排程結束日期：" + DGLTS;
        }

        private void txtActivity2_TextChanged(object sender, EventArgs e)
        {
            //攤提工時
            //作業6 = 機器
            txtActivity6.Text = txtActivity2.Text;
            txtActiunit6.Text = txtActiunit2.Text;
        }      
           
        private void txtActivity3_TextChanged(object sender, EventArgs e)
        {
            //攤提工時
            //作業4 = 人工
            txtActivity4.Text = txtActivity3.Text;
            txtActiunit5.Text = txtActiunit3.Text;
            //作業5 = 人工
            txtActivity5.Text = txtActivity3.Text;
            txtActiunit4.Text = txtActiunit3.Text;
        }      

        private void txtStart_Date_KeyPress(object sender, KeyPressEventArgs e)
        {

            keyIsAccept = detectKey(e);

            if (keyIsAccept)
            {
                e.Handled =  true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private bool detectKey(KeyPressEventArgs e)
        {
            // e.KeyChar == (Char)48 ~ 57 -----> 0~9
            // e.KeyChar == (Char)8 -----------> Backpace
            // e.KeyChar == (Char)13-----------> Enter

            // 數字鍵或是倒退鍵            
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                return true;
            } else
            {
                return false;
            }
        }

        private void txtFin_Date_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyIsAccept = detectKey(e);

            if (keyIsAccept)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void txtStart_Time_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyIsAccept = detectKey(e);

            if (keyIsAccept)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void txtFin_Time_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyIsAccept = detectKey(e);

            if (keyIsAccept)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }
        private void txtAufnr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    btnPO_Click(sender, e);
                    Cursor.Current = Cursors.Default;
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
                Cursor.Current = Cursors.WaitCursor;
                btnPO_Click(sender, e);
                Cursor.Current = Cursors.Default;
            }
            catch (Exception)
            {
                MessageBox.Show("工單" + txtAufnr.Text + "不存在", "錯誤");
            }
        }

        int start_year, end_year, start_mon, start_day, end_mon, end_day;

        private void txtStart_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyIsAccept = detectKey(e);

            if (keyIsAccept)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }

        }

        private void txtMachine_Calc(object sender, EventArgs e)
        {
            // 計算機器加工時間，單位秒數
            // for 加工組
            try {
                double machineTimeInSec, orderQty, machineTimeInMin;
                orderQty = Convert.ToInt32(txtYield.Text);

                if (orderQty == 0) orderQty = 1; // 有只報工時、沒報數量的狀況；時間仍要照常計算，因此數量不能為0

                if (!string.IsNullOrEmpty(txtMachineTime.Text))
                {
                    machineTimeInSec = Convert.ToInt32(txtMachineTime.Text);
                    machineTimeInMin = Math.Round((orderQty * machineTimeInSec) / 60, 0);
                    txtActivity2.Text = machineTimeInMin.ToString();
                }
                else
                {
                    txtMachineTime.Text = "0";
                }                
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
                    double menTimeInMin, machineTimeInMin, calcMachineTime;
                    menTimeInMin = Convert.ToInt32(txtActivity3.Text);
                    machineTimeInMin = Convert.ToInt32(txtActivity2.Text);

                    calcMachineTime = menTimeInMin - machineTimeInMin;

                    if (calcMachineTime < 0) {
                        txtActivity3.Text = "1";
                    } else { 
                        txtActivity3.Text = calcMachineTime.ToString();
                    }
                }
                catch
                {
                    MessageBox.Show("時間不能為負數", "錯誤");
                }
            }
}

        int start_hour, start_min, end_hour, end_min, sec, calcDay,calcHour,calcMinute;
        int convertToMniute, totalPersonHour;

        private void btnCalcTime_Click(object sender, EventArgs e)
        {
            try
            {
                //沒輸入日期，就用現在的日期
                if (txtStart_Date.Text.Length == 0 | txtEnd_Date.Text.Length == 0)
                {
                    txtStart_Date.Text = DateTime.Now.ToString("yyyyMMdd");
                    txtEnd_Date.Text = DateTime.Now.ToString("yyyyMMdd");
                }
                //沒輸入年份，就補上年份
                if (txtStart_Date.Text.Length == 4 | txtEnd_Date.Text.Length == 4)
                {
                    txtStart_Date.Text = DateTime.Now.ToString("yyyy") + txtStart_Date.Text;
                    txtEnd_Date.Text = DateTime.Now.ToString("yyyy") + txtEnd_Date.Text;
                }
                //沒輸入年份也沒輸入月份前置0，就都補上
                if (txtStart_Date.Text.Length == 3 | txtEnd_Date.Text.Length == 3)
                {
                    txtStart_Date.Text = DateTime.Now.ToString("yyyy") + "0" + txtStart_Date.Text;
                    txtEnd_Date.Text = DateTime.Now.ToString("yyyy") + "0" + txtEnd_Date.Text;
                }

                //日期或時間格式不對
                if (txtStart_Date.Text.Length != 8 | txtEnd_Date.Text.Length != 8 |
                    txtStart_Time.Text.Length != 4 | txtEnd_Time.Text.Length != 4)
                {
                    MessageBox.Show("日期或時間請輸入完整格式！ 例: 日期 20150105 ; 時間 0800", "錯誤");
                }
                else {
                    int totalBreakTime = 0;

                    start_year = Convert.ToUInt16(txtStart_Date.Text.Substring(0, 4));
                    start_mon = Convert.ToInt16(txtStart_Date.Text.Substring(4, 2));
                    start_day = Convert.ToInt16(txtStart_Date.Text.Substring(6, 2));

                    end_year = Convert.ToInt16(txtEnd_Date.Text.Substring(0, 4));
                    end_mon = Convert.ToInt16(txtEnd_Date.Text.Substring(4, 2));
                    end_day = Convert.ToInt16(txtEnd_Date.Text.Substring(6, 2));

                    start_hour = Convert.ToInt16(txtStart_Time.Text.Substring(0, 2));
                    start_min = Convert.ToInt16(txtStart_Time.Text.Substring(2));

                    end_hour = Convert.ToInt16(txtEnd_Time.Text.Substring(0, 2));
                    end_min = Convert.ToInt16(txtEnd_Time.Text.Substring(2));

                    totalBreakTime = calcBreakTime(start_hour, start_min, end_hour, end_min);

                    DateTime startDateTime = new DateTime(start_year, start_mon, start_day, start_hour, start_min, sec);
                    DateTime endDateTime = new DateTime(end_year, end_mon, end_day, end_hour, end_min, sec);

                    TimeSpan timeSpan = endDateTime.Subtract(startDateTime);

                    //d3算出來的結果會是{00年:00月:00天:00時:00分:00秒}
                    //因此要個別取值出來再運算，把結果轉換成分鐘數
                    calcDay = Convert.ToUInt16(timeSpan.Days.ToString());
                    calcHour = Convert.ToUInt16(timeSpan.Hours.ToString());
                    calcMinute = Convert.ToUInt16(timeSpan.Minutes.ToString());

                    convertToMniute = ((calcDay * 24) + calcHour) * 60 + calcMinute;

                    totalPersonHour = Convert.ToInt16(txtPerson.Text) * ( convertToMniute - totalBreakTime);
                    txtBreak_Time.Text = Convert.ToString(totalBreakTime);
                    txtActivity3.Text = Convert.ToString(totalPersonHour);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("日期格式有誤，請檢查");
            }
        }

        private int calcBreakTime(int start_hour, int start_min, int end_hour, int end_min)
        {
            int totalBreakTime = 0;
            TimeSpan tsStart = new TimeSpan( 0, start_hour, start_min, 0);
            TimeSpan tsEnd = new TimeSpan(0, end_hour, end_min, 0);

            // work time
            TimeSpan amWorkStart = new TimeSpan(0, 8, 0, 0);
            TimeSpan amWorkEnd = new TimeSpan(0, 12, 0, 0);
            TimeSpan pmWorkStart = new TimeSpan(0, 13, 0, 0);
            TimeSpan pmWorkEnd = new TimeSpan(0, 17, 0, 0);
            TimeSpan ovWorkStart = new TimeSpan(0, 17, 30, 0);
            TimeSpan ovWorkEnd = new TimeSpan(0, 19, 30, 0);

            // break time
            TimeSpan amBreakStart = new TimeSpan(0, 10, 00, 0);
            TimeSpan amBreakEnd = new TimeSpan(0, 10, 10, 0);
            TimeSpan noonBreakStart = new TimeSpan(0, 12, 0, 0);
            TimeSpan noonBreakEnd = new TimeSpan(0, 13, 0, 0);
            TimeSpan pmBreakStart = new TimeSpan(0, 15, 0, 0);
            TimeSpan pmBreakEnd = new TimeSpan(0, 15, 10, 0);
            TimeSpan ovBreakStart = new TimeSpan(0, 17, 0, 0);
            TimeSpan ovBreakEnd = new TimeSpan(0, 17, 30, 0);

            if (tsStart<amBreakStart &&  tsEnd >amBreakEnd) totalBreakTime += 10;
            if (tsStart<noonBreakStart &&  tsEnd > noonBreakEnd) totalBreakTime += 60;
            if (tsStart<pmBreakStart && tsEnd > pmBreakEnd) totalBreakTime += 10;
            if (tsStart<ovBreakStart &&  tsEnd > ovWorkEnd) totalBreakTime += 30;

            return totalBreakTime;
        }

        private void txtFin_Date_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnCalcTime_Click(sender, e);
            }
        }

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

        private void textBox1_Leave(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                btnCalcTime_Click(sender, e);
                Cursor.Current = Cursors.Default;
            }
            catch (Exception)
            {
                MessageBox.Show("請檢查輸入時間是否正確", "錯誤");
            }
        }
     }
    }

