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
            DataSet rsdb;
            DataSet rsActngPeriod;
            long count;
            string dteBeg;
            string dteEnd;
            int intMonth;
            int intYear;
            DateTime dteFirstDate;
            string strRetEarningsAcct;
            int intRetEarningsAcct;
            int intSubAccts;
            string[] strRetEarningsSub = new String[1000];



            DialogResult results = MessageBox.Show("WARNING: THIS PROGRAM WILL CHANGE THE FISCAL YEAR IN JURIS" + "\r\n"
                  + "ONLY RUN THIS UTILITY IF YOU HAVE BEEN ADVISED TO DO SO." + "\r\n"
                  + "ARE YOU SURE YOU WANT TO DO THIS?", "Continue?",  MessageBoxButtons.YesNo, MessageBoxIcon.Question);



            if (results == DialogResult.Yes)
            {

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
                strSQL = "If Exists (Select * from sysObjects where name = 'tmpPeriod') drop table tmpPeriod";
                _jurisUtility.ExecuteSql(0, strSQL);
                strSQL = "select * into tmpPeriod from ActngPeriod Where PrdNbr <> 0";
                _jurisUtility.ExecuteSql(0, strSQL);

                strSQL = "SELECT * FROM ActngPeriod";
                rsActngPeriod = _jurisUtility.RecordsetFromSQL(strSQL);
                foreach (DataRow row in rsActngPeriod.Tables[0].Rows)
                {
                    if (Convert.ToInt32(row["PrdNbr"].ToString()) == 0)
                    {
                        intYear = Convert.ToInt32(row["PrdYear"].ToString());
                        dteBeg = "01/01/" + intYear.ToString();
                        dteEnd = "12/31/" + intYear.ToString();

                        string SQL = "update tmpPeriod set PrdStartDate = '" + dteBeg + "', PrdEndDate = '" + dteEnd + "' where PrdYear = " + intYear.ToString() + " and PrdNbr = 0";
                        _jurisUtility.ExecuteNonQuery(0, SQL);
                    }
                    else
                    {
                        intMonth = Convert.ToInt32(row["PrdNbr"].ToString());
                        intYear = Convert.ToInt32(row["PrdYear"].ToString());
                        dteBeg = row["PrdNbr"].ToString() + "/01/" + intYear.ToString();
                        DateTime dt = Convert.ToDateTime(row["PrdNbr"].ToString() + "/01/" + intYear.ToString());
                        DateTime endOfMonth = new DateTime(dt.Year,
                                               dt.Month,
                                               DateTime.DaysInMonth(dt.Year,
                                                                    dt.Month));
                        dteEnd = endOfMonth.ToString("MM/dd/yyyy");

                        string SQL = "update tmpPeriod set PrdStartDate = '" + dteBeg + "', PrdEndDate = '" + dteEnd + "' where PrdYear = " + intYear.ToString() + " and PrdNbr = " + row["PrdNbr"].ToString();
                        _jurisUtility.ExecuteNonQuery(0, SQL);
                    }
                }

                rsActngPeriod.Clear();
 
                strSQL = "Delete from JournalEntry "
            + "where JEDate < '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "' ";

                strSQL = "update JournalEntry "
                       + "Set JEPrdNbr = PrdNbr, "
                       + " JEPrdYear = PrdYear "
                       + "From JournalEntry "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(JEDate) as varchar) + '/1/' + cast(year(JEDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                


                strSQL = "update JEBatchDetail "
                       + "Set JEBDPrdNbr = PrdNbr, "
                       + " JEBDPrdYear = PrdYear "
                       + "From JEBatchDetail "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(JEBDDate) as varchar) + '/1/' + cast(year(JEBDDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledExpenses.", 2, 20);

                strSQL = "update BilledExpenses "
                       + "Set BEPrdNbr = PrdNbr, "
                       + " BEPrdYear = PrdYear "
                       + "From BilledExpenses "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(BEDate) as varchar) + '/1/' + cast(year(BEDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledExpense", 3, 20);


                strSQL = "update UnbilledExpense "
                       + "Set UEPrdNbr = PrdNbr, "
                       + " UEPrdYear = PrdYear "
                       + "From UnbilledExpense "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(UEDate) as varchar) + '/1/' + cast(year(UEDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating BilledTime", 4, 20);

                strSQL = "update BilledTime "
                       + "Set BTPrdNbr = PrdNbr, "
                       + " BTPrdYear = PrdYear "
                       + "From BilledTime "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(BTDate) as varchar) + '/1/' + cast(year(BTDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                UpdateStatus("Updating UnbilledTime", 5, 20);


                strSQL = "update UnbilledTime "
                       + "Set UTPrdNbr = PrdNbr, "
                       + " UTPrdYear = PrdYear "
                       + "From UnbilledTime "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(UTDate) as varchar) + '/1/' + cast(year(UTDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpBatchDetail", 6, 20);


                strSQL = "Update ExpBatchDetail "
                       + "Set EBDPrdNbr = PrdNbr, "
                       + " EBDPrdYear = PrdYear "
                       + "From ExpBatchDetail "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(EBDDate) as varchar) + '/1/' + cast(year(EBDDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpenseEntry", 7, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'ExpenseEntry') Update ExpenseEntry "
                       + "Set PeriodNbr = PrdNbr, "
                       + " PeriodYear = PrdYear "
                       + "From ExpenseEntry "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(EntryDate) as varchar) + '/1/' + cast(year(EntryDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeBatchDetail", 8, 20);


                strSQL = "Update TimeBatchDetail "
                       + "Set TBDPrdNbr = PrdNbr, "
                       + " TBDPrdYear = PrdYear "
                       + "From TimeBatchDetail "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(TBDDate) as varchar) + '/1/' + cast(year(TBDDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TimeEntry", 9, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'TimeEntry') Update TimeEntry "
                       + "Set PeriodNumber = PrdNbr, "
                       + " PeriodYear = PrdYear "
                       + "From TimeEntry "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(EntryDate) as varchar) + '/1/' + cast(year(EntryDate) as varchar) "
                       + "where PeriodNumber <> PrdNbr or PeriodYear <> PrdYear";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating LedgerHistory", 10, 20);


                strSQL = "update LedgerHistory "
                       + "Set LHPrdNbr = PrdNbr, "
                       + " LHPrdYear = PrdYear "
                       + "From LedgerHistory "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(LHDate) as varchar) + '/1/' + cast(year(LHDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CashReceipt", 11, 20);


                strSQL = "Update CashReceipt "
                       + "Set CRPrdNbr = PrdNbr, "
                       + " CRPrdYear = PrdYear "
                       + "From CashReceipt "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(CRDate) as varchar) + '/1/' + cast(year(CRDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating CreditMemo", 12, 20);


                strSQL = "Update CreditMemo "
                       + "Set CMPrdNbr = PrdNbr, "
                       + " CMPrdYear = PrdYear "
                       + "From CreditMemo "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(CMDate) as varchar) + '/1/' + cast(year(CMDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrAdjBatchDetail", 13, 20);


                strSQL = "Update TrAdjBatchDetail "
                       + "Set TABDPrdNbr = PrdNbr, "
                       + " TABDPrdYear = PrdYear "
                       + "From TrAdjBatchDetail "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on PrdStartDate = cast(month(TABDDate) as varchar) + '/1/' + cast(year(TABDDate) as varchar)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating TrustSumByPrd", 14, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpTrustSumByPrd') drop table tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpTrustSumByPrd from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from TrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpTrustSumByPrd "
                       + "Set TSPPrdNbr = AP.PrdNbr, "
                       + "TSPPrdYear = AP.PrdYear "
                       + "From tmpTrustSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as OP on TSPPrdYear = PrdYear and TSPPrdNbr = PrdNbr "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on AP.PrdStartDate = OP.PrdStartDate";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into TrustSumByPrd (TSPMatter, TSPBank, TSPPrdYear, TSPPrdNbr, TSPDeposits, TSPPayments, TSPAdjustments) "
                       + "Select TSPMatter, TSPBank, TSPPrdYear, TSPPrdNbr, TSPDeposits, TSPPayments, TSPAdjustments from tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpTrustSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating VenSumByPrd", 15, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpVenSumByPrd') drop table tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpVenSumByPrd from VenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from VenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpVenSumByPrd "
                       + "Set VSPPrdNbr = AP.PrdNbr, "
                       + "VSPPrdYear = AP.PrdYear "
                       + "From tmpVenSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as OP on VSPPrdYear = PrdYear and VSPPrdNbr = PrdNbr "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on AP.PrdStartDate = OP.PrdStartDate";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into VenSumByPrd (VSPVendor, VSPPrdYear, VSPPrdNbr, VSPVouchers, VSPPayments, VSPDiscountsTaken) "
                       + "Select VSPVendor, VSPPrdYear, VSPPrdNbr, VSPVouchers, VSPPayments, VSPDiscountsTaken from tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpVenSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating ExpSumByPrd", 16, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpExpSumByPrd') drop table tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpExpSumByPrd from ExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from ExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpExpSumByPrd "
                       + "Set ESPPrdNbr = AP.PrdNbr, "
                       + "ESPPrdYear = AP.PrdYear "
                       + "From tmpExpSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as OP on ESPPrdYear = PrdYear and ESPPrdNbr = PrdNbr "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on AP.PrdStartDate = OP.PrdStartDate";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "insert into ExpSumByPrd (ESPMatter, ESPExpCd, ESPPrdYear, "
                       + "ESPPrdNbr, ESPEntered, ESPBilledValue, ESPBilledAmt, ESPReceived, ESPAdjusted) "
                       + "Select ESPMatter, ESPExpCd, ESPPrdYear, "
                       + "ESPPrdNbr, ESPEntered, ESPBilledValue, ESPBilledAmt, ESPReceived, ESPAdjusted from tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "drop table tmpExpSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating FeeSumByPrd", 17, 20);


                strSQL = "If Exists (Select * from sysObjects where name = 'tmpFeeSumByPrd') drop table tmpFeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "select * into tmpFeeSumByPrd from FeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "delete from FeeSumByPrd";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update tmpFeeSumByPrd "
                       + "Set FSPPrdNbr = AP.PrdNbr, "
                       + "FSPPrdYear = AP.PrdYear "
                       + "From tmpFeeSumByPrd "
                       + "inner join (select * from tmpPeriod where PrdNbr <> 0) as OP on FSPPrdYear = PrdYear and FSPPrdNbr = PrdNbr "
                       + "inner join (select * from ActngPeriod where PrdNbr <> 0) as AP on AP.PrdStartDate = OP.PrdStartDate";
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



                strSQL = "insert into ChartBudget (ChbAccount, ChbPrdYear, ChbPeriod, ChbBudget, ChbNetChange) "
                       + "SELECT distinct JEAccount, JEPrdYear, JEPrdNbr, 0.00 as Budget, 0.00 as Net "
                       + "From JournalEntry "
                       + "left join ChartBudget on ChbAccount = JEAccount and ChbPrdYear = JEPrdYear and ChbPeriod = JEPrdNbr "
                       + "Where ChbAccount Is Null";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update ChartBudget "
                       + "Set ChbNetChange = 0.00";
                _jurisUtility.ExecuteNonQuery(0, strSQL);
                strSQL = "Update ChartBudget "
                       + "Set ChbNetChange = JENetChange "
                       + "From ChartBudget "
                       + "inner join ("
                       + "SELECT JEAccount, JEPrdYear, JEPrdNbr, sum(JEAmount) as JENetChange "
                       + "FROM JournalEntry "
                       + "where JEDate >= '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "' "
                       + "group by JEAccount, JEPrdNbr, JEPrdYear) as JE "
                       + "on JEAccount = ChbAccount and JEPrdYear = ChbPrdYear and JEPrdNbr = ChbPeriod";
                _jurisUtility.ExecuteNonQuery(0, strSQL);

                UpdateStatus("Updating Accounting Year", 19, 20);


                strSQL = "Update ActngYear "
                       + "set AYCloseStatus = 'N' "
                       + "where AYYear > = (select MIN(JEPrdYear)-1 as FirstYear from JournalEntry)";
                _jurisUtility.ExecuteNonQuery(0, strSQL);



                UpdateStatus("All tables updated.", 20, 20);
                WriteLog("FiscalYearChangeTool: Accounting Year Changed to start in month " + dateTimePicker1.Value.ToString("MM/dd/yyyy"));

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

            System.Environment.Exit(0);
          
        }



        private void label1_Click(object sender, EventArgs e)
        {

        }


    }
}
