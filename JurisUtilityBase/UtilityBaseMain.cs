using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {

            string strW;
            string strSQL;
            string strDesc;
            long recAffected;
            DataSet rsdb;
            DataSet rsActngPeriod;
            string sConcat;
            int iNumAffected;
            long count;
            long lngBillTo;
            long lngPreCount;
            string strInputRecord;
            string strCliCode;
            string strMatCode;
            string strPC;
            DateTime dteBeg;
            DateTime dteEnd;
            int intFirstMonth;
            int intMonth;
            int intYear;
            DateTime dteFirstDate;
            string strRetEarningsAcct;
            int intRetEarningsAcct;
            int intSubAccts;
            string strRetEarningsSub;
            string CodeStr;
            string strMatSysNbr;
            string strBillingFld10;
            int intPos;


            DialogResult results = MessageBox.Show("WARNING: THIS PROGRAM WILL CHANGE THE FISCAL YEAR IN JURIS" + "\r\n"
                  + "ONLY RUN THIS UTILITY IF YOU HAVE BEEN ADVISED TO DO SO." + "\r\n"
                  + "ARE YOU SURE YOU WANT TO DO THIS?", "Continue?",  MessageBoxButtons.YesNo, MessageBoxIcon.Question);



            if (results == DialogResult.Yes)
            intFirstMonth = System.Convert.ToInt32(txtFirstMonth.Text);

            if (string.IsNullOrEmpty(txtFirstDate.Text))
                dteFirstDate = Convert.ToDateTime("01/01/1980");
            else
                dteFirstDate = Convert.ToDateTime(txtFirstDate.Text);

            rsdb = new DataSet();


            strSQL = "Select SpTxtValue from SysParam where SpName = 'RetEarnAcc'";
            rsdb = _jurisUtility.RecordsetFromSQL(strSQL);
            strRetEarningsAcct = rsdb.Tables[0].Rows[0][0].ToString();

            rsdb.Clear();

    strSQL = "select COUNT(*) as ct from ChartOfAccountsSubDefinition where COASDActive = 1";

    rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

    intSubAccts = Convert.ToInt32(rsdb.Tables[0].Rows[0][0].ToString());

    rsdb.Clear();

    strRetEarningsAcct = strRetEarningsAcct.Split(',')[0];

    if (strRetEarningsAcct.IndexOf("-") > -1)
    {
        strW = Strings.Mid(strRetEarningsAcct, Strings.InStr(strRetEarningsAcct, "-") + 1);
        strRetEarningsAcct = Strings.Left(strRetEarningsAcct, Strings.InStr(strRetEarningsAcct, "-") - 1);
        for (count = 0; count <= intSubAccts - 1; count++)
        {
            if (Strings.InStr(strW, "-") == 0)
                strRetEarningsSub[count] = strW;
            else
                strRetEarningsSub[count] = Strings.Left(strW, Strings.InStr(strW, "-") - 1);
            strW = Strings.Mid(strW, Strings.InStr(strW, "-") + 1);
        }
    }

    strSQL = "Select ChtSysNbr from ChartOfAccounts ";

    for (count = 0; count <= intSubAccts - 1; count++)
        strSQL = strSQL + " inner join COASubAccount" + Strings.Format(count + 1) + " on COAS" + Strings.Format(count + 1) + "ID = ChtSubAcct" + Strings.Format(count + 1);

    strSQL = strSQL + " Where ChtMainAcct = " + strRetEarningsAcct;

    for (count = 0; count <= intSubAccts - 1; count++)
        strSQL = strSQL + " and COAS" + Strings.Format(count + 1) + "Code = " + strRetEarningsSub[count];


    Debug.Print(strSQL);
    rsdb.Open(strSQL, Cn, ADODB.adOpenForwardOnly, ADODB.adLockReadOnly);
    rsdb.MoveFirst();

    intRetEarningsAcct = rsdb.ChtSysNbr;

    rsdb.Close();

    statBar.SimpleText = "Updating ActngPeriod";

    strSQL = "If Exists (Select * from sysObjects where name = 'tmpPeriod') drop table tmpPeriod";
    Cn.Execute(strSQL);
    strSQL = "select * into tmpPeriod from ActngPeriod Where PrdNbr <> 0";
    Cn.Execute(strSQL);














































































            string sql = "Create Table #Vch(rownbr int, vtype varchar(1) null, vdate datetime null, vendor varchar(50) null, ponbr varchar(50) null, invnbr varchar(50) null, invdate datetime null, duedate datetime null, " +
                        "   discdate datetime null, invamt decimal(20,2) null, nondisamt decimal(20,2) null, vchref varchar(8000) null, sepcheck varchar(1) null,apacct varchar(4) null, gldistacct varchar(20) null,glamt decimal(20,2) null, " +
                        " tbank varchar(4) null, expclient varchar(12) null, expmatter varchar(12) null,expcode varchar(4) null,exptask varchar(4) null, expunits decimal(12,2) null, expamount decimal(12,2) null, expnarrative varchar(8000), " +
                        " expnote varchar(60) null)";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            DataTable d2 = dataGridView1.DataSource as DataTable;
            int i = 1;

            if (d2.Rows.Count == 0)
            { MessageBox.Show("No vouchers to import.  Select a different file or close the application."); }
            else
            {
                foreach (DataRow dr in d2.Rows)
                {
                    string vtype = dr["Type"].ToString();
                    string VoucherDate = dr["VoucherDate"].ToString();
                    string VendorCode = dr["VendorCode"].ToString();
                    string PONbr = dr["PONbr"].ToString();
                    string InvoiceNbr = dr["InvoiceNbr"].ToString();
                    string DueDate = dr["DueDate"].ToString();
                    string InvoiceDate = dr["InvoiceDate"].ToString();
                    string DiscountDate = dr["DiscountDate"].ToString();
                    string InvoiceAmt = dr["InvoiceAmt"].ToString();
                    string NonDiscAmt = dr["NonDiscAmt"].ToString();
                    string VchReference = dr["VchReference"].ToString();
                    string SeparateCheck = dr["SeparateCheck"].ToString();
                    string APAcct = dr["APAcct"].ToString();
                    string GLDistAcct = dr["GLDistAcct"].ToString();
                    string GLAmt = dr["GLAmt"].ToString();
                    string TrustBank = dr["TrustBank"].ToString();
                    string ExpClient = dr["ExpClient"].ToString();
                    string ExpMatter = dr["ExpMatter"].ToString();
                    string ExpCode = dr["ExpCode"].ToString();
                    string ExpTaskCode = dr["ExpTaskCode"].ToString();
                    string ExpUnits = dr["ExpUnits"].ToString();
                    string ExpAmount = dr["ExpAmount"].ToString();
                    string ExpNarrative = dr["ExpNarrative"].ToString();
                    string ExpBillNote = dr["ExpBillNote"].ToString();

                    string s2 = "Insert into #Vch " +
                    "Values(" + i + ",'" + vtype + "',convert(datetime,'" + VoucherDate + "',101),'" + VendorCode + "','" + PONbr + "','" + InvoiceNbr + "',convert(datetime,'" + InvoiceDate + "',101),convert(datetime,'" + DueDate + "',101),convert(datetime,'" + DiscountDate + "',101), " +
                    "cast(isnull('" + InvoiceAmt + "','0') as decimal(12,2)), cast(isnull('" + NonDiscAmt + "','0') as decimal(12,2)), '" + VchReference + "','" + SeparateCheck + "','" + APAcct + "','" + GLDistAcct + "', cast(isnull('" + GLAmt + "','0') as money), " +
                    "'" + TrustBank + "','" + ExpClient + "','" + ExpMatter + "','" + ExpCode + "','" + ExpTaskCode + "',cast(isnull('" + ExpUnits + "','0') as decimal(12,2)),cast(isnull('" + ExpAmount + "','0') as decimal(12,2)),'" + ExpNarrative + "','" + ExpBillNote + "')";
                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                    i = i + 1;
                }
            }


            UpdateStatus("All MBF07 fields updated.", 1, 1);

            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            //if matter and originating timekeeper
            else if (false)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentOriginatingTimekeeper, 'DEF' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empinitials<>'ABC'";


            return reportSQL;
        }


    }
}
