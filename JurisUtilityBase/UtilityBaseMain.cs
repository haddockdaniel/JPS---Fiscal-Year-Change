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

        public DateTime origStartDate { get; set; }

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
                string dtO = "select convert(varchar(10),min(prdstartdate),101) as FirstDT, min(PrdYear) as NewDt from actngperiod";
                DataSet ds = _jurisUtility.RecordsetFromSQL(dtO);
                DataTable dt = ds.Tables[0];
                dtOrig.Text = dt.Rows[0]["FirstDT"].ToString();
                origStartDate = Convert.ToDateTime(dt.Rows[0]["FirstDT"].ToString());
                string DN = dt.Rows[0]["NewDt"].ToString();
                int d2 = Convert.ToInt32(DN) - 1;
                dateTimePicker1.Value = new DateTime(d2, 1,1);
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {

            string strW;
            string strSQL = "";
            DataSet rsdb;
            long count;
            DateTime dteFirstDate;
            string strRetEarningsAcct;
            int intRetEarningsAcct;
            int intSubAccts;
            string[] strRetEarningsSub = new String[1000];



            DialogResult results = MessageBox.Show("WARNING: THIS PROGRAM WILL CHANGE THE FISCAL YEAR IN JURIS" + "\r\n"
                  + "ONLY RUN THIS UTILITY IF YOU HAVE BEEN ADVISED TO DO SO." + "\r\n"
                  + "ARE YOU SURE YOU WANT TO DO THIS?", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);



            if (results == DialogResult.Yes)
            {
                Cursor.Current = Cursors.WaitCursor;

                if (dateTimePicker1.Value < Convert.ToDateTime("01/01/1980"))
                    dteFirstDate = Convert.ToDateTime("01/01/1980");
                else
                    dteFirstDate = dateTimePicker1.Value;

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
                    strW = strRetEarningsAcct.Substring(strRetEarningsAcct.IndexOf("-") + 1);
                    strRetEarningsAcct = strRetEarningsAcct.Substring(0, strRetEarningsAcct.IndexOf("-"));
                    for (count = 1; count <= intSubAccts; count++)
                    {
                        if (strW.IndexOf("-") == 0)
                            strRetEarningsSub[count] = strW;
                        else
                        {
                            if (strW.IndexOf("-") - 1 > 0)
                                strRetEarningsSub[count] = strW.Substring(0, strW.IndexOf("-") - 1);
                            else
                                strRetEarningsSub[count] = strW.Substring(0, strW.Length - 1);
                        }
                        strW = strW.Substring(strW.IndexOf("-") + 1);
                    }
                }


                strSQL = "Select ChtSysNbr from ChartOfAccounts ";

                for (count = 1; count <= intSubAccts; count++)
                    strSQL = strSQL + " inner join COASubAccount" + (count).ToString() + " on COAS" + (count).ToString() + "ID = ChtSubAcct" + (count).ToString();

                strSQL = strSQL + " Where ChtMainAcct = " + strRetEarningsAcct;

                for (count = 1; count <= intSubAccts; count++)
                    strSQL = strSQL + " and COAS" + (count).ToString() + "Code = " + strRetEarningsSub[count];

                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                intRetEarningsAcct = Convert.ToInt32(rsdb.Tables[0].Rows[0][0].ToString());

                rsdb.Clear();
                
                UpdateStatus("Updating the Acct Prd Table.", 1, 20);          
                toolStripStatusLabel.Text = "Updating the Acct Prd Table";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "select ayyear from ActngYear where AYYear = DATEPART(year, '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "')";
                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                //does the new start year exist?
                if (rsdb == null || rsdb.Tables.Count == 0 || rsdb.Tables[0].Rows.Count == 0)
                {
                    strSQL = "insert into ActngYear ([AYYear]  ,[AYNbrOfPrds] ,[AYCloseStatus]) values (" + "DATEPART(year, '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "'), 12, 'Y')";
                }

                rsdb.Clear();
                strSQL = "select  prdstartdate, prdenddate, prdnbr, prdyear, prdstate  into tmpPeriod from ActngPeriod";
                _jurisUtility.ExecuteSql(0, strSQL);
                 strSQL = "alter table tmpperiod add newDateStart datetime null";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                 strSQL = "alter table tmpperiod add newDateEnd datetime null";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                 strSQL = "alter table tmpperiod add newprd int null";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "alter table tmpperiod add newyr int null";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                int diffInMonths = ((dteFirstDate.Year - origStartDate.Year) * 12) + dteFirstDate.Month - origStartDate.Month;

                strSQL = "update tmpperiod set newdatestart =   DateAdd(month," + diffInMonths.ToString() +",prdstartdate) where prdnbr <> 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "update tmpperiod set newdateend =   DATEADD(d, -1, DATEADD(m, DATEDIFF(m, 0, newdatestart) + 1, 0)) where prdnbr <> 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "update tmpperiod set newprd = DATEPART(m, newdatestart) where prdnbr <> 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "update tmpperiod set newyr = DATEPART(year, newdatestart) where prdnbr <> 0 ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "update tmpperiod set newdatestart = '01/01/' + cast(DATEPART(year, prdstartdate) as varchar(5)), newdateend = '12/31/' + cast(DATEPART(year, prdstartdate) as varchar(5)) where prdnbr = 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "update tmpPeriod set newprd = 0, newyr = DATEPART(year, newdatestart) where prdnbr = 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


         //       strSQL = "Delete from JournalEntry "
       //  + "where JEDate < '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "' ";
         //       _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "update JournalEntry "
                       + "Set JEPrdNbr = AP.newprd, "
                       + " JEPrdYear = AP.newyr "
                       + "From JournalEntry "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on JEPrdYear = prdyear and JEPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);



                strSQL = "update JEBatchDetail "
                       + "Set JEBDPrdNbr = AP.newprd, "
                       + " JEBDPrdYear = AP.newyr "
                       + "From JEBatchDetail "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on JEBDPrdYear = prdyear and JEBDPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledExpenses.", 2, 20);
                toolStripStatusLabel.Text = "Updating BilledExpenses";
                statusStrip.Refresh();
                Application.DoEvents();

                strSQL = "update BilledExpenses "
                       + "Set BEPrdNbr = AP.newprd, "
                       + " BEPrdYear = AP.newyr "
                       + "From BilledExpenses "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on BEPrdYear = prdyear and BEPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledExpense", 3, 20);
                toolStripStatusLabel.Text = "Updating UnbilledExpense";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update UnbilledExpense "
                       + "Set UEPrdNbr = AP.newprd, "
                       + " UEPrdYear = AP.newyr "
                       + "From UnbilledExpense "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on UEPrdYear = prdyear and UEPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledTime", 4, 20);
                toolStripStatusLabel.Text = "Updating BilledTime";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update BilledTime "
                       + "Set BTPrdNbr = AP.newprd, "
                       + " BTPrdYear = AP.newyr "
                       + "From BilledTime "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on BTPrdYear = prdyear and BTPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledTime", 5, 20);
                toolStripStatusLabel.Text = "Updating UnbilledTime";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update UnbilledTime "
                       + "Set UTPrdNbr = AP.newprd, "
                       + " UTPrdYear =AP.newyr "
                       + "From UnbilledTime "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on UTPrdYear = prdyear and UTPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpBatchDetail", 6, 20);
                toolStripStatusLabel.Text = "Updating ExpBatchDetail";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update ExpBatchDetail "
                       + "Set EBDPrdNbr = AP.newprd, "
                       + " EBDPrdYear = AP.newyr "
                       + "From ExpBatchDetail "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on EBDPrdYear = prdyear and EBDPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpenseEntry", 7, 20);
                toolStripStatusLabel.Text = "Updating ExpenseEntry";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'ExpenseEntry') Update ExpenseEntry "
                       + "Set PeriodNbr = AP.newprd, "
                       + " PeriodYear = AP.newyr "
                       + "From ExpenseEntry "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on PeriodYear = prdyear and PeriodNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeBatchDetail", 8, 20);
                toolStripStatusLabel.Text = "Updating TimeBatchDetail";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update TimeBatchDetail "
                       + "Set TBDPrdNbr = AP.newprd, "
                       + " TBDPrdYear = AP.newyr "
                       + "From TimeBatchDetail "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on TBDPrdYear = prdyear and TBDPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeEntry", 9, 20);
                toolStripStatusLabel.Text = "Updating TimeEntry";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'TimeEntry') Update TimeEntry "
                       + "Set PeriodNumber = AP.newprd, "
                       + " PeriodYear = AP.newyr "
                       + "From TimeEntry "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on PeriodYear = prdyear and PeriodNumber = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating LedgerHistory", 10, 20);
                toolStripStatusLabel.Text = "Updating LedgerHistory";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update LedgerHistory "
                       + "Set LHPrdNbr = AP.newprd, "
                       + " LHPrdYear = AP.newyr "
                       + "From LedgerHistory "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on LHPrdYear = prdyear and LHPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CashReceipt", 11, 20);
                toolStripStatusLabel.Text = "Updating CashReceipt";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update CashReceipt "
                       + "Set CRPrdNbr = AP.newprd, "
                       + " CRPrdYear = AP.newyr "
                       + "From CashReceipt "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on CRPrdYear = prdyear and CRPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CreditMemo", 12, 20);
                toolStripStatusLabel.Text = "Updating CreditMemo";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update CreditMemo "
                       + "Set CMPrdNbr = AP.newprd, "
                       + " CMPrdYear = AP.newyr "
                       + "From CreditMemo "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on CMPrdYear = prdyear and CMPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrAdjBatchDetail", 13, 20);
                toolStripStatusLabel.Text = "Updating BilledExpenses";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update TrAdjBatchDetail "
                       + "Set TABDPrdNbr = AP.newprd, "
                       + " TABDPrdYear = AP.newyr "
                       + "From TrAdjBatchDetail "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on TABDPrdYear = prdyear and TABDPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrustSumByPrd", 14, 20);
                toolStripStatusLabel.Text = "Updating TrustSumByPrd";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpTrustSumByPrd') drop table tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpTrustSumByPrd from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpTrustSumByPrd "
                       + "Set TSPPrdNbr = AP.newprd, "
                       + "TSPPrdYear = AP.newyr "
                       + "From tmpTrustSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on TSPPrdYear = prdyear and TSPPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into TrustSumByPrd (TSPMatter, TSPBank, TSPPrdYear, TSPPrdNbr, TSPDeposits, TSPPayments, TSPAdjustments) "
                       + "Select TSPMatter, TSPBank, TSPPrdYear, TSPPrdNbr, TSPDeposits, TSPPayments, TSPAdjustments from tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating VenSumByPrd", 15, 20);

                toolStripStatusLabel.Text = "Updating VenSumByPrd";
                statusStrip.Refresh();
                Application.DoEvents();

                strSQL = "If Exists (Select * from sysObjects where name = 'tmpVenSumByPrd') drop table tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpVenSumByPrd from VenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from VenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpVenSumByPrd "
                       + "Set VSPPrdNbr = AP.newprd, "
                       + "VSPPrdYear = AP.newyr "
                       + "From tmpVenSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on VSPPrdYear = prdyear and VSPPrdNbr = prdnbr ";
                    ;
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "insert into VenSumByPrd (VSPVendor, VSPPrdYear, VSPPrdNbr, VSPVouchers, VSPPayments, VSPDiscountsTaken) "
                       + "Select VSPVendor, VSPPrdYear, VSPPrdNbr, VSPVouchers, VSPPayments, VSPDiscountsTaken from tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpSumByPrd", 16, 20);
                toolStripStatusLabel.Text = "Updating ExpSumByPrd";
                statusStrip.Refresh();
                Application.DoEvents();



                strSQL = "If Exists (Select * from sysObjects where name = 'tmpExpSumByPrd') drop table tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpExpSumByPrd from ExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from ExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpExpSumByPrd "
                       + "Set ESPPrdNbr = AP.newprd, "
                       + "ESPPrdYear = AP.newyr "
                       + "From tmpExpSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on ESPPrdYear = prdyear and ESPPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into ExpSumByPrd (ESPMatter, ESPExpCd, ESPPrdYear, "
                       + "ESPPrdNbr, ESPEntered, ESPBilledValue, ESPBilledAmt, ESPReceived, ESPAdjusted) "
                       + "Select ESPMatter, ESPExpCd, ESPPrdYear, "
                       + "ESPPrdNbr, ESPEntered, ESPBilledValue, ESPBilledAmt, ESPReceived, ESPAdjusted from tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating FeeSumByPrd", 17, 20);
                toolStripStatusLabel.Text = "Updating FeeSumByPrd";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpFeeSumByPrd') drop table tmpFeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpFeeSumByPrd from FeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from FeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpFeeSumByPrd "
                       + "Set FSPPrdNbr = AP.newprd, "
                       + "FSPPrdYear = AP.newyr "
                       + "From tmpFeeSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on FSPPrdYear = prdyear and FSPPrdNbr = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into FeeSumByPrd (FSPMatter, FSPTkpr, FSPTaskCd, FSPActivityCd, FSPPrdYear, FSPPrdNbr, "
                       + "FSPWorkedHrsEntered, FSPNonBilHrsEntered, FSPBilHrsEntered, FSPFeeEnteredStdValue, "
                       + "FSPFeeEnteredActualValue, FSPWorkedHrsBld, FSPHrsBilled, FSPFeeBldStdValue, "
                       + "FSPFeeBldActualValue, FSPFeeBldActualAmt, FSPFeeReceived, FSPFeeAdjusted) "
                       + "Select FSPMatter, FSPTkpr, FSPTaskCd, FSPActivityCd, FSPPrdYear, FSPPrdNbr, "
                       + "FSPWorkedHrsEntered, FSPNonBilHrsEntered, FSPBilHrsEntered, FSPFeeEnteredStdValue, "
                       + "FSPFeeEnteredActualValue, FSPWorkedHrsBld, FSPHrsBilled, FSPFeeBldStdValue, "
                       + "FSPFeeBldActualValue, FSPFeeBldActualAmt, FSPFeeReceived, FSPFeeAdjusted from tmpFeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpFeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating ChartBudget", 18, 20);

                toolStripStatusLabel.Text = "Updating ChartBudget";
                statusStrip.Refresh();
                Application.DoEvents();

                //reassign non zero
                strSQL = "If Exists (Select * from sysObjects where name = 'tmpChartBudget') drop table tmpChartBudget";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpChartBudget from ChartBudget";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpChartBudget "
                        + "Set ChbPeriod = AP.newprd, "
                        + "ChbPrdYear = AP.newyr "
                        + "From tmpChartBudget "
                        + "inner join (select * from tmpPeriod where PrdNbr <> 0) as AP on ChbPrdYear = prdyear and ChbPeriod = prdnbr ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from ChartBudget ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into chartbudget ([ChbAccount] ,[ChbPrdYear] ,[ChbPeriod] ,[ChbBudget] ,[ChbNetChange]) " +
                    " select [ChbAccount] ,[ChbPrdYear] ,[ChbPeriod] ,[ChbBudget] ,[ChbNetChange] from tmpChartBudget ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "select min(ayyear) as FirstYr from actngyear";
                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                DateTime tempdt = Convert.ToDateTime("01/01/" + rsdb.Tables[0].Rows[0][0].ToString());

                int diff = (dteFirstDate.Year - tempdt.Year);
                

                //reassign zero


                strSQL = "Update ChartBudget "
                        + " set ChbPrdYear = ChbPrdYear + " + diff.ToString() 
                        + " where  ChbPeriod = 0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

				

                UpdateStatus("Updating Accounting Year", 19, 20);
                toolStripStatusLabel.Text = "Updating Accounting Year";
                statusStrip.Refresh();
                Application.DoEvents();

                strSQL = "EXEC sp_MSforeachtable \"ALTER TABLE ? NOCHECK CONSTRAINT all\"";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from ActngPeriod";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into ActngPeriod (prdstartdate, prdenddate, prdnbr, prdyear, prdstate) " +
                "(SELECT newdateStart, newdateend, newprd, newyr, prdstate FROM tmpPeriod)" ;
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "EXEC sp_MSforeachtable \"ALTER TABLE ? CHECK CONSTRAINT all\"";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpPeriod";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "delete from actngyear where ayyear not in (select PrdYear from ActngPeriod)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "delete from documenttree where dtdocclass = 2000 and DTKeyL = '2020'";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "Update ActngYear "
                            + "set AYCloseStatus = 'N' "
                            + "where AYYear > = (select MIN(JEPrdYear)-1 as FirstYear from JournalEntry)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("All tables updated.", 20, 20);
                WriteLog("FiscalYearChangeTool: Accounting Year Changed to start in " + dateTimePicker1.Value.ToString("MM/dd/yyyy"));
                toolStripStatusLabel.Text = "All tables updated";
                statusStrip.Refresh();
                Application.DoEvents();
                Cursor.Current = Cursors.Default;

                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            }

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
           // if (string.IsNullOrEmpty(toAtty) || string.IsNullOrEmpty(fromAtty))
          //      MessageBox.Show("Please select from both Timekeeper drop downs", "Selection Error");
          //  else
          //  {
                //generates output of the report for before and after the change will be made to client
                string SQLTkpr = getReportSQL();

                DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                ReportDisplay rpds = new ReportDisplay(myRSTkpr);
                rpds.Show();

           // }
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


            return reportSQL;
        }


    }
}
