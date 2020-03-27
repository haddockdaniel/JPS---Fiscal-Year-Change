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
                    _jurisUtility.ExecuteNonQuery(0, strSQL);
                    strSQL = "insert into documenttree (dtdocid, [DTSystemCreated] ,[DTDocClass] ,[DTDocType],[DTParentID] ,[DTTitle] ,[DTKeyL],[DTKeyT]) " +
                       " values ((select max(dtdocid) + 1 from documenttree), 'Y', 2000, 'R', (select dtdocid from documenttree where dtdocclass = 2000 and dttitle = 'Accounting Periods'), " +
                       "'" + dateTimePicker1.Value.Year.ToString() + "', '" + dateTimePicker1.Value.Year.ToString() + "', null)";
                      _jurisUtility.ExecuteNonQuery(0, strSQL);
                }

                rsdb.Clear();

                strSQL = "select max(AYYear) from ActngYear";
                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                int finalYr = Convert.ToInt32(rsdb.Tables[0].Rows[0][0].ToString());

                rsdb.Clear();


                strSQL = "select prdenddate from actngperiod where prdnbr = (select spnbrvalue from sysparam where spname = 'CurAcctPrdNbr') and prdyear = (select spnbrvalue from sysparam where spname = 'CurAcctPrdYear')";
                
                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                string CurEndDate = rsdb.Tables[0].Rows[0][0].ToString();

                rsdb.Clear();

                strSQL = "select AYYear from ActngYear order by AYYear";
                rsdb = _jurisUtility.RecordsetFromSQL(strSQL);

                strSQL = "If Exists (Select * from sysObjects where name = 'tmpperiod') drop table tmpperiod";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "create table tmpperiod  (prdstartdate datetime null, prdenddate datetime null, prdnbr int null, prdyear int null, prdstate int null, oldPrd int null, oldYr int null)";
                _jurisUtility.ExecuteSql(0, strSQL);

                foreach (DataRow rr in rsdb.Tables[0].Rows)
                {
                    for (int i = 0; i < 13; i++)
                    {
                        if (i == 0)
                        {
                            int year = Convert.ToInt32(rr[0].ToString());
                            DateTime firstDay = new DateTime(year, 1, 1);
                            DateTime lastDay = new DateTime(year, 12, 31);
                            strSQL = "insert into tmpPeriod (prdstartdate, prdenddate, prdnbr, prdyear, prdstate) "
                             + "values ('" + firstDay.ToShortDateString() + "', '" + lastDay.ToShortDateString() + "', " + i.ToString() + ", " + rr[0].ToString() + ", 0 )";
                            _jurisUtility.ExecuteSql(0, strSQL);
                        }
                        else
                        {
                            DateTime ss = new DateTime(Convert.ToInt32(rr[0].ToString()), i, 1);
                            
                            strSQL = "insert into tmpPeriod (prdstartdate, prdenddate, prdnbr, prdyear, prdstate) "
                             + "values ('" + i.ToString() + "/01/" + rr[0].ToString() + "', '" + ss.AddMonths(1).AddDays(-1).ToShortDateString() + "', " + i.ToString() + ", " + rr[0].ToString() + ", 0 )";
                            _jurisUtility.ExecuteSql(0, strSQL);

                        }
                        
                    }

                }

                strSQL = "  update tmpperiod set oldprd = at.PrdNbr, oldyr = at.PrdYear " +
                        " from tmpperiod tt " +
                        " inner join ActngPeriod at on at.PrdStartDate = tt.PrdStartDate and " +
                        " at.PrdEndDate = tt.PrdEndDate and tt.PrdNbr <> 0 and at.PrdNbr <> 0";
                _jurisUtility.ExecuteSql(0, strSQL);
                //MessageBox.Show(";;");


                strSQL = "update JournalEntry "
                       + "Set JEPrdNbr = prdnbr, "
                       + " JEPrdYear = prdyear "
                       + "From JournalEntry "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on JEPrdYear = oldyr and JEPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "update JEBatchDetail "
                       + "Set JEBDPrdNbr = prdnbr, "
                       + " JEBDPrdYear = prdyear "
                       + "From JEBatchDetail "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on JEBDPrdYear = oldyr and JEBDPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledExpenses.", 2, 20);
                toolStripStatusLabel.Text = "Updating BilledExpenses";
                statusStrip.Refresh();
                Application.DoEvents();



                //update non zero periods
              //  strSQL = "Update ChartBudget "
                //       + "Set ChbNetChange = JENetChange "
                //       + "From ChartBudget "
                 //      + "inner join ( "
                 //      + "SELECT JEAccount, JEPrdYear, JEPrdNbr, sum(JEAmount) as JENetChange "
                 //      + "FROM JournalEntry "
                      // + "where JEDate >= convert(" + dteFirstDate.ToShortDateString() + ", varchar, 101) "
                //       + "group by JEAccount, JEPrdNbr, JEPrdYear) as JE "
                //       + "on JEAccount = ChbAccount and JEPrdYear = ChbPrdYear and JEPrdNbr = ChbPeriod";
            //    _jurisUtility.ExecuteNonQuery(0, strSQL);
             //   

                 
                 

                strSQL = "update BilledExpenses "
                       + "Set BEPrdNbr = prdnbr, "
                       + " BEPrdYear = prdyear "
                       + "From BilledExpenses "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on BEPrdYear = oldyr and BEPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledExpense", 3, 20);
                toolStripStatusLabel.Text = "Updating UnbilledExpense";
                statusStrip.Refresh();
                Application.DoEvents();


                //MessageBox.Show("kk");
                

                strSQL = "update UnbilledExpense "
                       + "Set UEPrdNbr = prdnbr, "
                       + " UEPrdYear = prdyear "
                       + "From UnbilledExpense "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on UEPrdYear = oldyr and UEPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledTime", 4, 20);
                toolStripStatusLabel.Text = "Updating BilledTime";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update BilledTime "
                       + "Set BTPrdNbr = prdnbr, "
                       + " BTPrdYear = prdyear "
                       + "From BilledTime "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on BTPrdYear = oldyr and BTPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledTime", 5, 20);
                toolStripStatusLabel.Text = "Updating UnbilledTime";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update UnbilledTime "
                       + "Set UTPrdNbr = prdnbr, "
                       + " UTPrdYear =prdyear "
                       + "From UnbilledTime "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on UTPrdYear = oldyr and UTPrdNbr = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpBatchDetail", 6, 20);
                toolStripStatusLabel.Text = "Updating ExpBatchDetail";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update ExpBatchDetail "
                       + "Set EBDPrdNbr = prdnbr, "
                       + " EBDPrdYear = prdyear "
                       + "From ExpBatchDetail "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on EBDPrdYear = oldyr and EBDPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpenseEntry", 7, 20);
                toolStripStatusLabel.Text = "Updating ExpenseEntry";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'ExpenseEntry') Update ExpenseEntry "
                       + "Set PeriodNbr = prdnbr, "
                       + " PeriodYear = prdyear "
                       + "From ExpenseEntry "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on PeriodYear = oldyr and PeriodNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeBatchDetail", 8, 20);
                toolStripStatusLabel.Text = "Updating TimeBatchDetail";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update TimeBatchDetail "
                       + "Set TBDPrdNbr = prdnbr, "
                       + " TBDPrdYear = prdyear "
                       + "From TimeBatchDetail "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on TBDPrdYear = oldyr and TBDPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeEntry", 9, 20);
                toolStripStatusLabel.Text = "Updating TimeEntry";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "If Exists (Select * from sysObjects where name = 'TimeEntry') Update TimeEntry "
                       + "Set PeriodNumber = prdnbr, "
                       + " PeriodYear = prdyear "
                       + "From TimeEntry "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on PeriodYear = oldyr and PeriodNumber = oldprd ";

                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating LedgerHistory", 10, 20);
                toolStripStatusLabel.Text = "Updating LedgerHistory";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "update LedgerHistory "
                       + "Set LHPrdNbr = prdnbr, "
                       + " LHPrdYear = prdyear "
                       + "From LedgerHistory "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on LHPrdYear = oldyr and LHPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CashReceipt", 11, 20);
                toolStripStatusLabel.Text = "Updating CashReceipt";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update CashReceipt "
                       + "Set CRPrdNbr = prdnbr, "
                       + " CRPrdYear = prdyear "
                       + "From CashReceipt "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on CRPrdYear = oldyr and CRPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CreditMemo", 12, 20);
                toolStripStatusLabel.Text = "Updating CreditMemo";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update CreditMemo "
                       + "Set CMPrdNbr = prdnbr, "
                       + " CMPrdYear = prdyear "
                       + "From CreditMemo "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on CMPrdYear = oldyr and CMPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrAdjBatchDetail", 13, 20);
                toolStripStatusLabel.Text = "Updating BilledExpenses";
                statusStrip.Refresh();
                Application.DoEvents();


                strSQL = "Update TrAdjBatchDetail "
                       + "Set TABDPrdNbr = prdnbr, "
                       + " TABDPrdYear = prdyear "
                       + "From TrAdjBatchDetail "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on TABDPrdYear = oldyr and TABDPrdNbr = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrustSumByPrd", 14, 20);
                toolStripStatusLabel.Text = "Updating TrustSumByPrd";
                statusStrip.Refresh();
                Application.DoEvents();

                        strSQL = "If Exists (Select * from sysObjects where name = 'tmpTrustSumByPrd') drop table tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL); ;     


        
                strSQL = "select * into tmpTrustSumByPrd from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpTrustSumByPrd "
                       + "Set TSPPrdNbr = prdnbr, "
                       + "TSPPrdYear = prdyear "
                       + "From tmpTrustSumByPrd "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on TSPPrdYear = oldyr and TSPPrdNbr = oldprd ";
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
                       + "Set VSPPrdNbr = prdnbr, "
                       + "VSPPrdYear = prdyear "
                       + "From tmpVenSumByPrd "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on VSPPrdYear = oldyr and VSPPrdNbr = oldprd ";

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
                       + "Set ESPPrdNbr = prdnbr, "
                       + "ESPPrdYear = prdyear "
                       + "From tmpExpSumByPrd "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on ESPPrdYear = oldyr and ESPPrdNbr = oldprd ";
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
                       + "Set FSPPrdNbr = prdnbr, "
                       + "FSPPrdYear = prdyear "
                       + "From tmpFeeSumByPrd "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on FSPPrdYear = oldyr and FSPPrdNbr = oldprd ";
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

                strSQL = "update ChartBudget "
                    + " set chbprdyear=prdyear from tmpperiod "
                    + " where chbprdyear=(select min(chbprdyear) from chartbudget) and chbperiod=0 "
                    + " and prdyear=(select year(prdstartdate) from actngperiod where prdnbr=0 and chbprdyear=actngperiod.prdyear)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpRetEarn') drop table tmpRetEarn";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "select chtsysnbr as SysNbr into tmpRetEarn "
+ " from(select sysnbr, mainacct, case when charindex('-', subacct) = 0 then subacct else left(subacct, charindex('-', subacct) - 1) end as Sub1, "
+ " case when charindex('-', subacct) = 0 then '' else right(subacct, len(subacct) - charindex('-', subacct)) end as Sub2  "
+ " from(select sysnbr, case when charindex('-', AcctNbr) = 0 then AcctNbr else left(acctnbr, charindex('-', acctnbr) - 1) end As Mainacct, "
+ " case when charindex('-', AcctNbr) = 0 then '' else right(acctnbr, len(acctnbr) - charindex('-', acctnbr)) end As SubAcct, AcctNbr "
+ " from(select spnbrvalue as SysNbr, case when charindex(',', sptxtvalue) = 0 then sptxtvalue else left(sptxtvalue, charindex(',', sptxtvalue) - 1) end as AcctNbr "
+ "  from sysparam where spname = 'RetEarnAcc')REA) R2) RetEarn, "
+ "  (select chtsysnbr, chtmainacct, coas1code, coas2code, coas3code, coas4code "
+ "   from chartofaccounts "
+ "  left outer  join COASubAccount1 on coas1id = chtsubacct1 "
+ " left outer join coasubaccount2 on coas2id = chtsubacct2 "
+ "  left outer join coasubaccount3 on coas3id = chtsubacct3 "
+ "  left outer join coasubaccount4 on coas4id = chtsubacct4) COA "
+ "  where(chtsysnbr = sysnbr and sysnbr <> 0) or(sysnbr = 0 and chtmainacct = mainacct and coas1code is null) or "
+ "       (sysnbr = 0 and chtmainacct = mainacct and coas1code = right('00000000' + sub1, 8) and coas2code is null) or "
 + "             (sysnbr = 0 and chtmainacct = mainacct and coas1code = right('00000000' + sub1, 8)  and coas2code = right('00000000' + sub2, 8)  and coas3code is null) ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "Insert into ChartBudget(chbaccount, chbprdyear, chbperiod, chbbudget, chbnetchange) "
                    + "select chtsysnbr, prdyear, 0, 0.00, 0.00 "
                    + " from (select chtsysnbr, prdyear, prdnbr from chartofaccounts, actngperiod where prdyear>=(select min(chbprdyear) from chartbudget) and prdnbr=0) CB "
                   + " left outer join chartbudget on chbaccount=chtsysnbr and chbprdyear=prdyear and chbperiod=prdnbr "
                   + " where chbaccount is null";


                strSQL = "Update ChartBudget "
                    + " set chbnetchange=0 where chbperiod=0 and chbprdyear<>(select min(chbprdyear) from chartbudget) ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "Update ChartBudget "
                        + "Set chbperiod = prdnbr, "
                       + " chbprdyear = prdyear "
                       + "From ChartBudget "
                       + "       inner join (select * from tmpperiod where PrdNbr <> 0) as AP on chbprdyear = oldyr and chbperiod = oldprd ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = "Update ChartBudget set chbnetchange=0.00 from chartbudget, tmpRetEarn where chbaccount=sysnbr and chbperiod=0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "Update ChartBudget"
                    + " set chbnetchange=begbalance "
                    + " from ( select chbaccount as Acct, AYYear, sum(case when chbprdyear<AYYear then chbnetchange else 0 end) as BegBalance " 
                    + "  from chartbudget " 
                    + " inner join chartofaccounts on chtsysnbr = chbaccount, actngyear " 
                    + " where chtfinstmttype = 'B'  and  ayyear> (select min(chbprdyear) from chartbudget) "
                    + " group by chbaccount, AYYear) CB "
                    + " where chbaccount=Acct and chbprdyear=ayyear and chbperiod=0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "Update ChartBudget set chbnetchange=0.00 from chartbudget "
                           + " inner join chartofaccounts on chtsysnbr = chbaccount  where chtfinstmttype = 'P' and chbperiod=0";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

          

                  strSQL = "Update ChartBudget "
                  +" set chbnetchange = Total " 
                  +" from( select chbprdyear as PYear, chbaccount as Account, sum(total * -1) as Total "
                   +" from(select sysnbr, chbprdyear as PYear, sum(chbnetchange) as Total "
                    +" from chartbudget, tmpRetEarn "
                    +" where chbperiod = 0 and chbaccount<>sysnbr"
                    +" group by sysnbr, chbprdyear) CB "
                    +" inner join chartbudget on chbprdyear = pyear "
                    +"  where chbprdyear = pyear  and chbperiod = 0 and chbaccount=sysnbr "
                    +" group by chbprdyear, chbaccount) CB "
                   +"  where account=chbaccount and chbprdyear=pyear and chbperiod=0";
                 

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
                "(SELECT prdstartdate, prdenddate, prdnbr, prdyear, prdstate FROM tmpPeriod)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "EXEC sp_MSforeachtable \"ALTER TABLE ? CHECK CONSTRAINT all\"";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpPeriod";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

               // strSQL = "delete from actngyear where ayyear not in (select PrdYear from ActngPeriod)";
              // _jurisUtility.ExecuteNonQuery(0, strSQL);

               // strSQL = "delete from documenttree where dtdocclass = 2000 and DTKeyL = '2020'";
              //  _jurisUtility.ExecuteNonQuery(0, strSQL);


                strSQL = "Update ActngYear "
                            + "set AYCloseStatus = 'Y' where ayyear<>(select max(jeprdyear) from journalentry)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                strSQL = " update sysparam set spnbrvalue = prdnbr "
+ " from actngperiod "
+ " where prdenddate = convert(varchar(10),'" + CurEndDate + "',101) and prdnbr<>0 and spname = 'CurAcctPrdNbr' ";
                _jurisUtility.ExecuteNonQuery(0, strSQL);



                strSQL = "update sysparam set spnbrvalue = prdyear "
+ " from actngperiod "
+ " where prdenddate = convert(varchar(10),'" + CurEndDate + "',101) and prdnbr<>0 and spname = 'CurAcctPrdYear' ";
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
