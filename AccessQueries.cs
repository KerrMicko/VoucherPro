using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static VoucherPro.DataClass;
using static VoucherPro.AccessToDatabase;
using static VoucherPro.DataClass.CheckTableExpensesAndItems;
using QBFC16Lib;

namespace VoucherPro
{
    internal class AccessQueries
    {
        public List<CheckTable> GetCheckData(string refNumber)
        {
            List<CheckTable> check = new List<CheckTable>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    // AppliedToTxnTxnDate
                    string query = "SELECT TxnDate," +
                        "RefNumber, Amount, " +
                        "PayeeEntityRefFullName " +
                        "FROM [Check] WHERE RefNumber = ? " +
                        "UNION ALL " +
                        "SELECT TxnDate, " +
                        "RefNumber, " +
                        "Amount, " +
                        "PayeeEntityRefFullName " +
                        "FROM [BillPaymentCheck] WHERE RefNumber = ?";

                    using (OleDbCommand checkCommand = new OleDbCommand(query, accessConnection))
                    {
                        checkCommand.Parameters.AddWithValue("RefNumber", OdbcType.VarChar).Value = refNumber;
                        checkCommand.Parameters.AddWithValue("RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader checkReader = checkCommand.ExecuteReader())
                        {
                            while (checkReader.Read())
                            {
                                CheckTable newCheck = new CheckTable
                                {
                                    DateCreated = Convert.ToDateTime(checkReader["TxnDate"]).Date,
                                    RefNumber = checkReader["RefNumber"].ToString(),
                                    Amount = Convert.ToDouble(checkReader["Amount"]),
                                    PayeeFullName = checkReader["PayeeEntityRefFullName"].ToString(),
                                };

                                check.Add(newCheck);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return check;
        }

        public List<CheckTableGrid> GetCheckDataIVP(string refNumber)
        {
            List<CheckTableGrid> checkList = new List<CheckTableGrid>();
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                sessionManager.OpenConnection2("", "VoucherPro Check Data", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // ----------------------------------------------------------------
                // 1. QUERY FOR REGULAR CHECKS
                // ----------------------------------------------------------------
                ICheckQuery checkQuery = request.AppendCheckQueryRq();
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // ----------------------------------------------------------------
                // 2. QUERY FOR BILL PAYMENT CHECKS
                // ----------------------------------------------------------------
                IBillPaymentCheckQuery billPayQuery = request.AppendBillPaymentCheckQueryRq();
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // Execute Requests
                IMsgSetResponse response = sessionManager.DoRequests(request);

                // ----------------------------------------------------------------
                // PROCESS RESPONSE 1: REGULAR CHECKS
                // ----------------------------------------------------------------
                IResponse qbResponseCheck = response.ResponseList.GetAt(0);
                ICheckRetList checkRetList = qbResponseCheck.Detail as ICheckRetList;

                if (checkRetList != null)
                {
                    for (int i = 0; i < checkRetList.Count; i++)
                    {
                        ICheckRet checkRet = checkRetList.GetAt(i);
                        string docNum = checkRet.RefNumber.GetValue();

                        if (docNum != refNumber) continue;

                        CheckTableGrid newCheck = new CheckTableGrid
                        {
                            DateCreated = checkRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = checkRet.Amount.GetValue(),
                            PayeeFullName = checkRet.PayeeEntityRef != null ? checkRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        };
                        checkList.Add(newCheck);
                    }
                }

                // ----------------------------------------------------------------
                // PROCESS RESPONSE 2: BILL PAYMENT CHECKS
                // ----------------------------------------------------------------
                IResponse qbResponseBillPay = response.ResponseList.GetAt(1);
                IBillPaymentCheckRetList billPayRetList = qbResponseBillPay.Detail as IBillPaymentCheckRetList;

                if (billPayRetList != null)
                {
                    for (int i = 0; i < billPayRetList.Count; i++)
                    {
                        IBillPaymentCheckRet billPayRet = billPayRetList.GetAt(i);
                        string docNum = billPayRet.RefNumber.GetValue();

                        if (docNum != refNumber) continue;

                        CheckTableGrid newCheck = new CheckTableGrid
                        {
                            DateCreated = billPayRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = billPayRet.Amount.GetValue(),
                            PayeeFullName = billPayRet.PayeeEntityRef != null ? billPayRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        };
                        checkList.Add(newCheck);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from QuickBooks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sessionManager != null)
                {
                    try
                    {
                        sessionManager.EndSession();
                        sessionManager.CloseConnection();
                    }
                    catch { }
                }
            }

            return checkList;
        }


        public List<BillTable> GetBillData_LEADS(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_APV(accessConnectionString);
                string entityRefListID = string.Empty;
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    // Retrieve data from Access database
                    string query = @"select TOP 1000 
                        BillPaymentCheckLine.TxnDate AS BillPayment_TxnDate,
                        BillPaymentCheckLine.PayeeEntityRefFullname,
                        BillPaymentCheckLine.AddressAddr1,
                        BillPaymentCheckLine.AddressAddr2,
                        BillPaymentCheckLine.BankAccountRefFullName,
                        BillPaymentCheckLine.Amount,
                        BillPaymentCheckLine.Refnumber,
                        BillPaymentCheckLine.AppliedToTxnRefNumber,
                        BillPaymentCheckLine.AppliedToTxnTxnID,
                        BillPaymentCheckLine.APAccountRefFullName,
                        BillPaymentCheckLine.APAccountRefListID,
                        BillPaymentCheckLine.Memo,
                        Bill.Memo,
                        Bill.AmountDue,
                        Bill.DueDate,
                        Bill.VendorReflistID,
                        Bill.TxnDate AS Bill_TxnDate
                        FROM 
                        BillPaymentCheckLine
                        INNER JOIN 
                        Bill ON BillPaymentCheckLine.AppliedToTxnTxnID = Bill.TxnID
                        Where BillPaymentCheckLine.Refnumber = ?";

                    using (OleDbCommand command = new OleDbCommand(query, accessConnection))
                    {
                        command.Parameters.AddWithValue("RefNumber", OleDbType.VarChar).Value = refNumber;
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                BillTable newBill = new BillTable
                                {
                                    // BillPaymentCheckLine table
                                    DateCreated = reader["BillPayment_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["BillPayment_TxnDate"]).Date : DateTime.MinValue,
                                    DueDate = reader["Bill_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["Bill_TxnDate"]).Date : DateTime.MinValue,
                                    PayeeFullName = reader["PayeeEntityRefFullName"] != DBNull.Value ? reader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    Address = reader["AddressAddr1"] != DBNull.Value ? reader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = reader["AddressAddr2"] != DBNull.Value ? reader["AddressAddr2"].ToString() : string.Empty,
                                    BankAccount = reader["BankAccountRefFullName"] != DBNull.Value ? reader["BankAccountRefFullName"].ToString() : string.Empty,
                                    APAccountRefFullName = reader["APAccountRefFullName"] != DBNull.Value ? reader["APAccountRefFullName"].ToString() : string.Empty,
                                    Amount = reader["Amount"] != DBNull.Value ? Convert.ToDouble(reader["Amount"]) : 0.0,
                                    RefNumber = reader["RefNumber"] != DBNull.Value ? reader["RefNumber"].ToString() : string.Empty,
                                    AppliedRefNumber = reader["AppliedToTxnRefNumber"] != DBNull.Value ? reader["AppliedToTxnRefNumber"].ToString() : string.Empty,
                                    AppliedToTxnTxnID = reader["AppliedToTxnTxnID"] != DBNull.Value ? reader["AppliedToTxnTxnID"].ToString() : string.Empty,
                                    Memo = reader["BillPaymentCheckLine.Memo"] != DBNull.Value ? reader["BillPaymentCheckLine.Memo"].ToString() : string.Empty,
                                    BillMemo = reader["Bill.Memo"] != DBNull.Value ? reader["Bill.Memo"].ToString() : string.Empty,
                                    AmountDue = reader["AmountDue"] != DBNull.Value ? Convert.ToDouble(reader["AmountDue"]) : 0.0,

                                    // Increment
                                    IncrementalID = nextID.ToString("D6")
                                };


                                string secondQuery = @"SELECT TOP 1000
                                        BillItemLine.ItemLineItemRefFullName AS AccountRefFullName, 
                                        BillItemLine.ItemLineAmount AS Amount,
                                        BillItemLine.ItemLineDesc AS ItemExpenseMemo
                                    FROM 
                                        BillItemLine 
                                    WHERE 
                                        BillItemLine.TxnID = ?

                                    UNION ALL

                                    SELECT
                                        BillExpenseLine.ExpenseLineAccountRefFullName AS AccountRefFullName, 
                                        BillExpenseLine.ExpenseLineAmount AS Amount,
                                        BillExpenseLine.ExpenseLineMemo AS ItemExpenseMemo
                                    FROM 
                                        [BillExpenseLine]
                                    WHERE 
                                        BillExpenseLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillItemLine.TxnID", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];
                                        secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ItemLineItemRefFullName = itemLineItemRefFullName,
                                                    ItemLineAmount = itemLineAmount,
                                                    ItemLineMemo = itemLineItemMemo,
                                                });
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }
                                string ThirdQuery = @"SELECT 
                                           ListID
                                        FROM 
                                            Vendor
                                        WHERE 
                                            Vendor.ListID = ?";
                                using (OleDbConnection ThirdConnection = new OleDbConnection(accessConnectionString))
                                {
                                    ThirdConnection.Open();

                                    using (OleDbCommand ThirdCommand = new OleDbCommand(ThirdQuery, ThirdConnection))
                                    {
                                        ThirdCommand.Parameters.AddWithValue("Vendor.ListID", OleDbType.VarChar).Value = reader["VendorReflistID"];

                                        using (OleDbDataReader ThirdReader = ThirdCommand.ExecuteReader())
                                        {
                                            if (ThirdReader.HasRows) // Check if there are rows returned by the third query
                                            {
                                                while (ThirdReader.Read())
                                                {
                                                    //newBill.TinID = ThirdReader["CustomFieldTINNO"] != DBNull.Value ? ThirdReader["CustomFieldTINNO"].ToString() : string.Empty;
                                                    //newBill.POnumber = ThirdReader["CustomFieldPONO"] != DBNull.Value ? ThirdReader["CustomFieldPONO"].ToString() : string.Empty;
                                                }

                                                // Execute fourth query only if there are rows in the third query result
                                                /*string fourthQuery = @"SELECT 
                                                           Name, AccountNumber
                                                        FROM 
                                                            Account
                                                        WHERE 
                                                            Account.listID = ? ";

                                                using (OleDbConnection fourthConnection = new OleDbConnection(accessConnectionString))
                                                {
                                                    fourthConnection.Open();

                                                    using (OleDbCommand fourthCommand = new OleDbCommand(fourthQuery, fourthConnection))
                                                    {
                                                        fourthCommand.Parameters.AddWithValue("Account.ListID", OleDbType.VarChar).Value = reader["APAccountRefListID"];

                                                        using (OleDbDataReader fourthReader = fourthCommand.ExecuteReader())
                                                        {
                                                            if (fourthReader.HasRows) // Check if there are rows returned by the fourth query
                                                            {
                                                                while (fourthReader.Read())
                                                                {
                                                                    newBill.AccountName = fourthReader["Name"] != DBNull.Value ? fourthReader["Name"].ToString() : string.Empty;
                                                                    newBill.AccountNumber = fourthReader["AccountNumber"] != DBNull.Value ? fourthReader["AccountNumber"].ToString() : string.Empty;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    fourthConnection.Close();
                                                }*/
                                            }
                                        }
                                    }

                                    ThirdConnection.Close();
                                }
                                if (reader["VendorReflistID"] != DBNull.Value)
                                {
                                    entityRefListID = reader["VendorReflistID"].ToString();
                                }
                                bills.Add(newBill);
                            }
                        }
                    }
                    string transactionQuery = @"SELECT TOP 4 
                               Transaction.TxnID, 
                               Transaction.TxnDate, 
                               Transaction.Memo, 
                               Transaction.RefNumber, 
                               Transaction.Amount,
                               Transaction.EntityRefListID 
                           FROM 
                               [Transaction] 
                           WHERE 
                               [Transaction].RefNumber IS NOT NULL 
                               AND [Transaction].TxnType = 'billpaymentcheck' 
                               AND [Transaction].EntityRefListID = ?
                           ORDER BY 
                               [Transaction].TimeModified DESC";

                    using (OleDbCommand transactionCommand = new OleDbCommand(transactionQuery, accessConnection))
                    {
                        transactionCommand.Parameters.AddWithValue("Transaction.EntityRefListID", OleDbType.VarChar).Value = entityRefListID;

                        using (OleDbDataReader transactionReader = transactionCommand.ExecuteReader())
                        {
                            while (transactionReader.Read())
                            {
                                BillTable newTransaction = new BillTable
                                {
                                    DateCreatedHistory = transactionReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(transactionReader["TxnDate"]).Date : DateTime.MinValue,
                                    MemoHistory = transactionReader["Memo"] != DBNull.Value ? transactionReader["Memo"].ToString() : string.Empty,
                                    RefNumberHistory = transactionReader["Refnumber"] != DBNull.Value ? transactionReader["Refnumber"].ToString() : string.Empty,
                                    AmountHistory = transactionReader["Amount"] != DBNull.Value ? Convert.ToDouble(transactionReader["Amount"]) : 0.0,
                                };
                                string billHistoryQuery = @"SELECT 
                                                               HistoryCVNumber, 
                                                               HistoryAPVNumber, 
                                                               Remarks
                                                           FROM 
                                                               CheckHistory
                                                           WHERE 
                                                               CheckHistory.RefNumber = ? ";

                                using (OleDbConnection billhistoryConnection = new OleDbConnection(accessConnectionString))
                                {
                                    billhistoryConnection.Open();

                                    using (OleDbCommand billhistoryCommand = new OleDbCommand(billHistoryQuery, billhistoryConnection))
                                    {
                                        billhistoryCommand.Parameters.AddWithValue("CheckHistory.RefNumber", OleDbType.VarChar).Value = transactionReader["RefNumber"];

                                        using (OleDbDataReader billhistoryReader = billhistoryCommand.ExecuteReader())
                                        {
                                            while (billhistoryReader.Read())
                                            {
                                                newTransaction.HistoryCVNumber = billhistoryReader["HistoryCVNumber"] != DBNull.Value ? billhistoryReader["HistoryCVNumber"].ToString() : string.Empty;
                                                newTransaction.HistoryAPVNumber = billhistoryReader["HistoryAPVNumber"] != DBNull.Value ? billhistoryReader["HistoryAPVNumber"].ToString() : string.Empty;
                                                newTransaction.Remarks = billhistoryReader["Remarks"] != DBNull.Value ? billhistoryReader["Remarks"].ToString() : string.Empty;
                                            }
                                        }
                                    }

                                    billhistoryConnection.Close();
                                }
                                bills.Add(newTransaction);

                            }
                        }
                    }

                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from bill Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return bills;
        } // CV BillPaymentCheckLine

        public List<BillTable> GetBillData_KAYAK(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_APV(accessConnectionString);
                string entityRefListID = string.Empty;
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    // Retrieve data from Access database
                    string query = @"select TOP 1000 
                        BillPaymentCheckLine.TxnDate AS BillPayment_TxnDate,
                        BillPaymentCheckLine.PayeeEntityRefFullname,
                        BillPaymentCheckLine.AddressAddr1,
                        BillPaymentCheckLine.AddressAddr2,
                        BillPaymentCheckLine.BankAccountRefFullName,
                        BillPaymentCheckLine.Amount,
                        BillPaymentCheckLine.Refnumber,
                        BillPaymentCheckLine.AppliedToTxnRefNumber,
                        BillPaymentCheckLine.AppliedToTxnTxnID,
                        BillPaymentCheckLine.APAccountRefFullName,
                        BillPaymentCheckLine.APAccountRefListID,
                        BillPaymentCheckLine.Memo,
                        Bill.Memo,
                        Bill.AmountDue,
                        Bill.DueDate,
                        Bill.VendorReflistID,
                        Bill.TxnDate AS Bill_TxnDate
                        FROM 
                        BillPaymentCheckLine
                        INNER JOIN 
                        Bill ON BillPaymentCheckLine.AppliedToTxnTxnID = Bill.TxnID
                        Where BillPaymentCheckLine.Refnumber = ?";

                    using (OleDbCommand command = new OleDbCommand(query, accessConnection))
                    {
                        command.Parameters.AddWithValue("RefNumber", OleDbType.VarChar).Value = refNumber;
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                BillTable newBill = new BillTable
                                {
                                    // BillPaymentCheckLine table
                                    DateCreated = reader["BillPayment_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["BillPayment_TxnDate"]).Date : DateTime.MinValue,
                                    DueDate = reader["Bill_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["Bill_TxnDate"]).Date : DateTime.MinValue,
                                    PayeeFullName = reader["PayeeEntityRefFullName"] != DBNull.Value ? reader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    Address = reader["AddressAddr1"] != DBNull.Value ? reader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = reader["AddressAddr2"] != DBNull.Value ? reader["AddressAddr2"].ToString() : string.Empty,
                                    BankAccount = reader["BankAccountRefFullName"] != DBNull.Value ? reader["BankAccountRefFullName"].ToString() : string.Empty,
                                    APAccountRefFullName = reader["APAccountRefFullName"] != DBNull.Value ? reader["APAccountRefFullName"].ToString() : string.Empty,
                                    Amount = reader["Amount"] != DBNull.Value ? Convert.ToDouble(reader["Amount"]) : 0.0,
                                    RefNumber = reader["RefNumber"] != DBNull.Value ? reader["RefNumber"].ToString() : string.Empty,
                                    AppliedRefNumber = reader["AppliedToTxnRefNumber"] != DBNull.Value ? reader["AppliedToTxnRefNumber"].ToString() : string.Empty,
                                    AppliedToTxnTxnID = reader["AppliedToTxnTxnID"] != DBNull.Value ? reader["AppliedToTxnTxnID"].ToString() : string.Empty,
                                    Memo = reader["BillPaymentCheckLine.Memo"] != DBNull.Value ? reader["BillPaymentCheckLine.Memo"].ToString() : string.Empty,
                                    BillMemo = reader["Bill.Memo"] != DBNull.Value ? reader["Bill.Memo"].ToString() : string.Empty,
                                    AmountDue = reader["AmountDue"] != DBNull.Value ? Convert.ToDouble(reader["AmountDue"]) : 0.0,

                                    // Increment
                                    IncrementalID = nextID.ToString("D6")
                                };
                                string accountNumberQuery = "SELECT AccountNumber FROM Account WHERE ListID = ?";
                                using (OleDbCommand accountNumberCmd = new OleDbCommand(accountNumberQuery, accessConnection))
                                {
                                    accountNumberCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = reader["APAccountRefListID"];
                                    object accountNumberResult = accountNumberCmd.ExecuteScalar();
                                    if (accountNumberResult != null && accountNumberResult != DBNull.Value)
                                    {
                                        newBill.AccountNumber = accountNumberResult.ToString();
                                    }
                                }

                                string secondQuery = @"SELECT TOP 1000
                                        BillItemLine.ItemLineItemRefFullName AS AccountRefFullName, 
                                        BillItemLine.ItemLineAmount AS Amount,
                                        BillItemLine.ItemLineDesc AS ItemExpenseMemo
                                    FROM 
                                        BillItemLine 
                                    WHERE 
                                        BillItemLine.TxnID = ?

                                    UNION ALL

                                    SELECT
                                        BillExpenseLine.ExpenseLineAccountRefFullName AS AccountRefFullName, 
                                        BillExpenseLine.ExpenseLineAmount AS Amount,
                                        BillExpenseLine.ExpenseLineMemo AS ItemExpenseMemo
                                    FROM 
                                        [BillExpenseLine]
                                    WHERE 
                                        BillExpenseLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillItemLine.TxnID", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];
                                        secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ItemLineItemRefFullName = itemLineItemRefFullName,
                                                    ItemLineAmount = itemLineAmount,
                                                    ItemLineMemo = itemLineItemMemo,
                                                });
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }
                                bills.Add(newBill);
                            }
                        }
                    }

                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from bill Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return bills;
        }

        public List<BillTable> GetBillData_CPI(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();
            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_APV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string query = @"select TOP 1000 
                BillPaymentCheckLine.TxnDate AS BillPayment_TxnDate,
                BillPaymentCheckLine.PayeeEntityRefFullname,
                BillPaymentCheckLine.AddressAddr1,
                BillPaymentCheckLine.AddressAddr2,
                BillPaymentCheckLine.BankAccountRefFullName,
                BillPaymentCheckLine.Amount,
                BillPaymentCheckLine.Refnumber,
                BillPaymentCheckLine.AppliedToTxnRefNumber,
                BillPaymentCheckLine.AppliedToTxnTxnID,
                BillPaymentCheckLine.APAccountRefFullName,
                BillPaymentCheckLine.APAccountRefListID,
                BillPaymentCheckLine.Memo,
                Bill.Memo,
                Bill.AmountDue,
                Bill.DueDate,
                Bill.VendorReflistID,
                Bill.TxnDate AS Bill_TxnDate
                FROM 
                BillPaymentCheckLine
                INNER JOIN 
                Bill ON BillPaymentCheckLine.AppliedToTxnTxnID = Bill.TxnID
                Where BillPaymentCheckLine.Refnumber = ?";

                    using (OleDbCommand command = new OleDbCommand(query, accessConnection))
                    {
                        command.Parameters.AddWithValue("RefNumber", OleDbType.VarChar).Value = refNumber;
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                BillTable newBill = new BillTable
                                {
                                    DateCreated = reader["BillPayment_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["BillPayment_TxnDate"]).Date : DateTime.MinValue,
                                    DueDate = reader["Bill_TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["Bill_TxnDate"]).Date : DateTime.MinValue,
                                    PayeeFullName = reader["PayeeEntityRefFullName"]?.ToString() ?? string.Empty,
                                    Address = reader["AddressAddr1"]?.ToString() ?? string.Empty,
                                    Address2 = reader["AddressAddr2"]?.ToString() ?? string.Empty,
                                    BankAccount = reader["BankAccountRefFullName"]?.ToString() ?? string.Empty,
                                    APAccountRefFullName = reader["APAccountRefFullName"]?.ToString() ?? string.Empty,
                                    Amount = reader["Amount"] != DBNull.Value ? Convert.ToDouble(reader["Amount"]) : 0.0,
                                    RefNumber = reader["RefNumber"]?.ToString() ?? string.Empty,
                                    AppliedRefNumber = reader["AppliedToTxnRefNumber"]?.ToString() ?? string.Empty,
                                    AppliedToTxnTxnID = reader["AppliedToTxnTxnID"]?.ToString() ?? string.Empty,
                                    Memo = reader["BillPaymentCheckLine.Memo"]?.ToString() ?? string.Empty,
                                    BillMemo = reader["Bill.Memo"]?.ToString() ?? string.Empty,
                                    AmountDue = reader["AmountDue"] != DBNull.Value ? Convert.ToDouble(reader["AmountDue"]) : 0.0,
                                    IncrementalID = nextID.ToString("D6")
                                };

                                // Fetch Account Number
                                string accountNumberQuery = "SELECT AccountNumber FROM Account WHERE ListID = ?";
                                using (OleDbCommand accountNumberCmd = new OleDbCommand(accountNumberQuery, accessConnection))
                                {
                                    accountNumberCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = reader["APAccountRefListID"];
                                    object accountNumberResult = accountNumberCmd.ExecuteScalar();
                                    if (accountNumberResult != null && accountNumberResult != DBNull.Value)
                                        newBill.AccountNumber = accountNumberResult.ToString();
                                }

                                // --- UPDATED QUERY: Includes Class Columns ---
                                string secondQuery = @"SELECT TOP 1000
                                BillItemLine.ItemLineItemRefFullName AS AccountRefFullName, 
                                BillItemLine.ItemLineAmount AS Amount,
                                BillItemLine.ItemLineDesc AS ItemExpenseMemo,
                                BillItemLine.ItemLineClassRefFullName AS LineClass
                            FROM 
                                BillItemLine 
                            WHERE 
                                BillItemLine.TxnID = ?

                            UNION ALL

                            SELECT
                                BillExpenseLine.ExpenseLineAccountRefFullName AS AccountRefFullName, 
                                BillExpenseLine.ExpenseLineAmount AS Amount,
                                BillExpenseLine.ExpenseLineMemo AS ItemExpenseMemo,
                                BillExpenseLine.ExpenseLineClassRefFullName AS LineClass
                            FROM 
                                [BillExpenseLine]
                            WHERE 
                                BillExpenseLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();
                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("p1", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];
                                        secondCommand.Parameters.AddWithValue("p2", OleDbType.VarChar).Value = reader["AppliedToTxnTxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string partName = secondReader["AccountRefFullName"]?.ToString() ?? string.Empty;
                                                double amount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string memo = secondReader["ItemExpenseMemo"]?.ToString() ?? string.Empty;

                                                // MAPPING THE CLASS
                                                string classValue = secondReader["LineClass"]?.ToString() ?? string.Empty;

                                                // DEBUG LOG: Verify the fetch
                                                Console.WriteLine($"FETCH DEBUG: Part='{partName}' | Class='{classValue}'");

                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ItemLineItemRefFullName = partName,
                                                    ItemLineAmount = amount,
                                                    ItemLineMemo = memo,
                                                    // Assigning to both to ensure the final Insert logic picks it up
                                                    ItemLineClassRefFullName = classValue,
                                                    ExpenseLineClassRefFullName = classValue,
                                                    ExpenseLineItemRefFullName = partName, // Map this for Expense check
                                                    ExpenseLineAmount = amount,
                                                    ExpenseLineMemo = memo
                                                });
                                            }
                                        }
                                    }
                                    secondConnection.Close();
                                }
                                bills.Add(newBill);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            return bills;
        }

        public List<BillTable> GetBillData_KAYAKdirect(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<BillTable> bills = new List<BillTable>();

            try
            {
                // 1. Get Incremental ID
                string accessConnectionString = GetAccessConnectionString();
                string nextIDStr = GetNextIncrementalID_APV(accessConnectionString).ToString("D6");

                sessionManager.OpenConnection2("", "KAYAK Bill Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // ====================================================
                // 1. QUERY BILL PAYMENT CHECK
                // ====================================================
                IMsgSetRequest req1 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                IBillPaymentCheckQuery bpcQuery = req1.AppendBillPaymentCheckQueryRq();
                bpcQuery.IncludeLineItems.SetValue(true);
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                IMsgSetResponse resp1 = sessionManager.DoRequests(req1);
                IResponse r1 = resp1.ResponseList.GetAt(0);
                IBillPaymentCheckRetList bpList = r1.Detail as IBillPaymentCheckRetList;

                if (bpList == null || bpList.Count == 0) return bills;

                IBillPaymentCheckRet bp = bpList.GetAt(0);

                // Header data from Payment Check
                DateTime payDate = bp.TxnDate?.GetValue() ?? DateTime.MinValue;
                string payee = bp.PayeeEntityRef?.FullName?.GetValue() ?? "";
                string bankAccount = bp.BankAccountRef?.FullName?.GetValue() ?? "";
                string paymentMemo = bp.Memo?.GetValue() ?? "";
                double checkAmount = bp.Amount?.GetValue() ?? 0;

                // Collect all Applied Bill TxnIDs
                List<string> appliedTxnIDs = new List<string>();
                if (bp.AppliedToTxnRetList != null)
                {
                    for (int k = 0; k < bp.AppliedToTxnRetList.Count; k++)
                    {
                        string tId = bp.AppliedToTxnRetList.GetAt(k).TxnID?.GetValue();
                        if (!string.IsNullOrEmpty(tId)) appliedTxnIDs.Add(tId);
                    }
                }

                if (appliedTxnIDs.Count == 0) return bills;

                // ====================================================
                // 2. QUERY ALL LINKED BILLS IN ONE BATCH
                // ====================================================
                IMsgSetRequest req2 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                IBillQuery billQuery = req2.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);

                foreach (string id in appliedTxnIDs)
                {
                    billQuery.ORBillQuery.TxnIDList.Add(id);
                }

                IMsgSetResponse resp2 = sessionManager.DoRequests(req2);
                IBillRetList billList = resp2.ResponseList.GetAt(0).Detail as IBillRetList;

                if (billList == null) return bills;

                // ====================================================
                // 3. PROCESS EACH BILL RETRIEVED
                // ====================================================
                for (int bIndex = 0; bIndex < billList.Count; bIndex++)
                {
                    IBillRet bill = billList.GetAt(bIndex);

                    BillTable bt = new BillTable
                    {
                        IncrementalID = nextIDStr,
                        DateCreated = payDate,
                        DueDate = payDate,
                        PayeeFullName = payee,
                        BankAccount = bankAccount,
                        RefNumber = refNumber, // Check Ref
                        AppliedRefNumber = bill.RefNumber?.GetValue() ?? "", // Bill Ref
                        AppliedToTxnTxnID = bill.TxnID?.GetValue() ?? "",
                        Memo = paymentMemo,
                        BillMemo = bill.Memo?.GetValue() ?? "",
                        Amount = checkAmount,
                        AmountDue = bill.AmountDue?.GetValue() ?? 0,
                        APAccountRefFullName = bill.APAccountRef?.FullName?.GetValue() ?? "",
                        // Get AP Account Number
                        AccountNumber = GetAccountNumber(sessionManager, bill.APAccountRef?.ListID?.GetValue())
                    };

                    // Expense Lines
                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int i = 0; i < bill.ExpenseLineRetList.Count; i++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(i);
                            bt.ItemDetails.Add(new ItemDetail
                            {
                                ItemLineItemRefFullName = exp.AccountRef?.FullName?.GetValue() ?? "",
                                ItemLineAmount = exp.Amount?.GetValue() ?? 0,
                                ItemLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ItemLineMemo = exp.Memo?.GetValue() ?? "",
                                // Maintain compatibility with your previous Expense naming
                                ExpenseLineAmount = exp.Amount?.GetValue() ?? 0,
                                ExpenseLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? ""
                            });
                        }
                    }

                    // Item Lines
                    if (bill.ORItemLineRetList != null)
                    {
                        for (int i = 0; i < bill.ORItemLineRetList.Count; i++)
                        {
                            var orItem = bill.ORItemLineRetList.GetAt(i);
                            if (orItem.ItemLineRet != null)
                            {
                                var item = orItem.ItemLineRet;
                                bt.ItemDetails.Add(new ItemDetail
                                {
                                    ItemLineItemRefFullName = item.ItemRef?.FullName?.GetValue() ?? "",
                                    ItemLineAmount = item.Amount?.GetValue() ?? 0,
                                    ItemLineClassRefFullName = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemLineMemo = item.Desc?.GetValue() ?? ""
                                });
                            }
                        }
                    }
                    bills.Add(bt);
                }
            }
            catch (Exception ex) { MessageBox.Show("KAYAK Bill Retrieval Error: " + ex.Message); }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); } catch { }
            }

            return bills;
        }

        public List<BillTable> GetBillData_IVP(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<BillTable> bills = new List<BillTable>();

            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"[DEBUG] START: GetBillData_IVP for RefNumber: {refNumber}");

            try
            {
                sessionManager.OpenConnection2("", "Bill Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                Console.WriteLine("[DEBUG] Session Opened Successfully.");

                // ====================================================
                // 1. QUERY BILL PAYMENT CHECK USING RefNumber
                // ====================================================
                IMsgSetRequest req1 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req1.Attributes.OnError = ENRqOnError.roeContinue;

                IBillPaymentCheckQuery bpcQuery = req1.AppendBillPaymentCheckQueryRq();
                bpcQuery.IncludeLineItems.SetValue(true);

                // exact match
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                Console.WriteLine("[DEBUG] Sending BillPaymentCheck Query...");
                IMsgSetResponse resp1 = sessionManager.DoRequests(req1);
                IResponse r1 = resp1.ResponseList.GetAt(0);

                IBillPaymentCheckRetList bpList = r1.Detail as IBillPaymentCheckRetList;

                if (bpList == null || bpList.Count == 0)
                {
                    MessageBox.Show("Bill Payment Check not found: " + refNumber);
                    return bills;
                }

                IBillPaymentCheckRet bp = bpList.GetAt(0);

                // HEADER FROM BILL PAYMENT CHECK (These stay constant for all bills in this check)
                DateTime payDate = bp.TxnDate?.GetValue() ?? DateTime.MinValue;
                string payee = bp.PayeeEntityRef?.FullName?.GetValue() ?? "";
                string address1 = bp.Address?.Addr1?.GetValue() ?? "";
                string address2 = bp.Address?.Addr2?.GetValue() ?? "";
                string bankAccount = bp.BankAccountRef?.FullName?.GetValue() ?? "";
                string memo = bp.Memo?.GetValue() ?? "";
                double amountPaid = bp.Amount?.GetValue() ?? 0;

                // ====================================================
                // *** CHANGED: GET ALL APPLIED BILL TxnIDs (NOT JUST INDEX 0)
                // ====================================================
                List<string> appliedTxnIDs = new List<string>();

                if (bp.AppliedToTxnRetList != null && bp.AppliedToTxnRetList.Count > 0)
                {
                    Console.WriteLine($"[DEBUG] AppliedToTxn List Count: {bp.AppliedToTxnRetList.Count}");
                    // Loop through ALL applied transactions
                    for (int k = 0; k < bp.AppliedToTxnRetList.Count; k++)
                    {
                        var applied = bp.AppliedToTxnRetList.GetAt(k);
                        string tId = applied.TxnID?.GetValue();
                        if (!string.IsNullOrEmpty(tId))
                        {
                            appliedTxnIDs.Add(tId);
                            Console.WriteLine($"[DEBUG] Found Applied Bill TxnID: {tId}");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Applied Bill found from Bill Payment Check.");
                    return bills;
                }

                // ====================================================
                // 2. QUERY BILL(S) USING THE COLLECTED TxnIDs
                // ====================================================
                IMsgSetRequest req2 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req2.Attributes.OnError = ENRqOnError.roeContinue;

                IBillQuery billQuery = req2.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);

                // *** CHANGED: Add ALL TxnIDs to the query list
                foreach (string id in appliedTxnIDs)
                {
                    billQuery.ORBillQuery.TxnIDList.Add(id);
                }

                Console.WriteLine($"[DEBUG] Sending Bill Query for {appliedTxnIDs.Count} bills...");
                IMsgSetResponse resp2 = sessionManager.DoRequests(req2);
                IResponse r2 = resp2.ResponseList.GetAt(0);

                IBillRetList billList = r2.Detail as IBillRetList;

                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bills not found for the provided TxnIDs.");
                    return bills;
                }

                // ====================================================
                // *** CHANGED: LOOP THROUGH ALL RETRIEVED BILLS
                // ====================================================
                Console.WriteLine($"[DEBUG] Retrieved {billList.Count} Bill(s). Processing...");

                for (int bIndex = 0; bIndex < billList.Count; bIndex++)
                {
                    IBillRet bill = billList.GetAt(bIndex);

                    // BILL HEADER FIELDS
                    DateTime billDate = bill.TxnDate?.GetValue() ?? DateTime.MinValue;
                    DateTime dueDate = bill.DueDate?.GetValue() ?? DateTime.MinValue;
                    double amountDue = bill.AmountDue?.GetValue() ?? 0;
                    string billMemo = bill.Memo?.GetValue() ?? "";
                    string billAPAccount = bill.APAccountRef?.FullName?.GetValue() ?? "";
                    string billRefNumber = bill.RefNumber?.GetValue() ?? "";
                    string specificTxnID = bill.TxnID?.GetValue() ?? "";

                    Console.WriteLine($"[DEBUG] Processing Bill #{bIndex + 1}: Ref {billRefNumber}");

                    // Create BillTable object for THIS specific bill
                    BillTable bt = new BillTable
                    {
                        DateCreated = payDate,
                        DueDate = payDate, // Or dueDate depending on your report requirement
                        PayeeFullName = payee,
                        Address = address1,
                        Address2 = address2,
                        BankAccount = bankAccount,
                        APAccountRefFullName = billAPAccount,
                        Amount = amountPaid, // This is the Check Total
                        RefNumber = refNumber, // This is the Check Ref Number
                        AppliedRefNumber = billRefNumber, // This is the specific Bill Ref Number
                        AppliedToTxnTxnID = specificTxnID,
                        Memo = memo,
                        BillMemo = billMemo,
                        AmountDue = amountDue, // The amount of this specific bill
                    };

                    // Process Expense Lines for THIS bill
                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int i = 0; i < bill.ExpenseLineRetList.Count; i++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(i);
                            bt.ItemDetails.Add(new ItemDetail
                            {
                                ItemLineItemRefFullName = exp.AccountRef?.FullName?.GetValue() ?? "",
                                ItemLineAmount = exp.Amount?.GetValue() ?? 0,
                                ItemLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ItemLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ItemLineMemo = exp.Memo?.GetValue() ?? "",
                            });
                        }
                    }

                    // Process Item Lines for THIS bill
                    if (bill.ORItemLineRetList != null)
                    {
                        for (int i = 0; i < bill.ORItemLineRetList.Count; i++)
                        {
                            var orItem = bill.ORItemLineRetList.GetAt(i);
                            if (orItem.ItemLineRet != null)
                            {
                                var item = orItem.ItemLineRet;
                                bt.ItemDetails.Add(new ItemDetail
                                {
                                    ItemLineItemRefFullName = item.ItemRef?.FullName?.GetValue() ?? "",
                                    ItemLineAmount = item.Amount?.GetValue() ?? 0,
                                    ItemLineClassRefFullName = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemLineCustomerJob = item.CustomerRef?.FullName?.GetValue() ?? "",
                                    ItemLineMemo = item.Desc?.GetValue() ?? "",
                                });
                            }
                        }
                    }

                    // Add THIS bill to the main list
                    bills.Add(bt);
                }

                Console.WriteLine($"[DEBUG] Successfully added {bills.Count} bills to the return list.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] EXCEPTION: {ex.Message}");
                MessageBox.Show("Error retrieving Bill data: " + ex.Message);
            }
            finally
            {
                try
                {
                    sessionManager.EndSession();
                    sessionManager.CloseConnection();
                }
                catch { }
            }

            return bills;
        }

        public List<ItemReciept> GetItemRecieptData_LEADS(string refNumber)
        {
            List<ItemReciept> ItemReceipt = new List<ItemReciept>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                //int nextID = GetNextIncrementalID_CV(accessConnectionString);
                string entityRefListID = string.Empty;

                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string itemQuery = @"SELECT TOP 1000
                                              ItemReceipt.VendorRefFullName,
                                              ItemReceipt.VendorRefListID,
                                              ItemReceipt.TxnID AS ItemReceiptTxnID,
                                              ItemReceipt.TxnDate,
                                              ItemReceipt.RefNumber,
                                              ItemReceipt.Memo,
                                              ItemReceipt.TotalAmount,
                                              ItemReceipt.APAccountRefFullName,
                                              ItemReceipt.APAccountRefListID,
                                              ItemReceiptItemLine.TxnID AS ItemLineTxnID,
                                              ItemReceiptItemLine.ItemLineItemRefFullName,
                                              ItemReceiptItemLine.ItemLineCustomerRefFullName,
                                              ItemReceiptItemLine.ItemLineClassRefFullName,
                                              ItemReceiptItemLine.ItemLineDesc,
                                              ItemReceiptItemLine.ItemLineQuantity,
                                              ItemReceiptItemLine.ItemLineCost,
                                              ItemReceiptItemLine.ItemLineUnitOfMeasure,
                                              ItemReceiptItemLine.ItemLineAmount
                                    FROM
                                              [ItemReceipt]
                                    INNER JOIN 
                                              ItemReceiptItemLine ON ItemReceipt.TxnID = ItemReceiptItemLine.TxnID
                                    WHERE
                                              ItemReceipt.RefNumber = ? ";

                    using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, accessConnection))
                    {
                        itemCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader itemReader = itemCommand.ExecuteReader())
                        {
                            while (itemReader.Read())
                            {
                                ItemReciept newCheckItem = new ItemReciept
                                {
                                    DateCreated = itemReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(itemReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = itemReader["APAccountRefFullName"] != DBNull.Value ? itemReader["APAccountRefFullName"].ToString() : string.Empty,
                                    PayeeFullName = itemReader["VendorRefFullName"] != DBNull.Value ? itemReader["VendorRefFullName"].ToString() : string.Empty,
                                    RefNumber = itemReader["RefNumber"] != DBNull.Value ? itemReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = itemReader["TotalAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["TotalAmount"]) : 0.0,
                                    Memo = itemReader["Memo"] != DBNull.Value ? itemReader["Memo"].ToString() : string.Empty,

                                    Item = itemReader["ItemLineItemRefFullName"] != DBNull.Value ? itemReader["ItemLineItemRefFullName"].ToString() : string.Empty,
                                    ItemDescription = itemReader["ItemLineDesc"] != DBNull.Value ? itemReader["ItemLineDesc"].ToString() : string.Empty,
                                    ItemClass = itemReader["ItemLineClassRefFullName"] != DBNull.Value ? itemReader["ItemLineClassRefFullName"].ToString() : string.Empty,
                                    ItemCustomerJob = itemReader["ItemLineCustomerRefFullName"] != DBNull.Value ? itemReader["ItemLineCustomerRefFullName"].ToString() : string.Empty,
                                    ItemAmount = itemReader["ItemLineAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineAmount"]) : 0.0,
                                    ItemQuantity = itemReader["ItemLineQuantity"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineQuantity"]) : 0.0,
                                    ItemCost = itemReader["ItemLineCost"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineCost"]) : 0.0,
                                    ItemUM = itemReader["ItemLineUnitOfMeasure"] != DBNull.Value ? itemReader["ItemLineUnitOfMeasure"].ToString() : string.Empty,
                                    ReceiptItemType = ReceiptItemType.ReceiptItem,

                                    //IncrementalID = nextID,
                                    //IncrementalID = nextID.ToString("D6")
                                };
                                string SecondQuery = @"SELECT 
                                        VendorAddressAddr1, VendorAddressAddr2, VendorAddressAddr3, VendorAddressAddr4, VendorAddressCity
                                     FROM 
                                         Vendor
                                     WHERE 
                                         Vendor.ListID = ?";
                                using (OleDbConnection SecondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    SecondConnection.Open();

                                    using (OleDbCommand SecondCommand = new OleDbCommand(SecondQuery, SecondConnection))
                                    {
                                        SecondCommand.Parameters.AddWithValue("Vendor.ListID", OleDbType.VarChar).Value = itemReader["VendorRefListID"];

                                        using (OleDbDataReader SecondReader = SecondCommand.ExecuteReader())
                                        {
                                            while (SecondReader.Read())
                                            {
                                                newCheckItem.Addr1 = SecondReader["VendorAddressAddr1"] != DBNull.Value ? SecondReader["VendorAddressAddr1"].ToString() : string.Empty;
                                                newCheckItem.Addr2 = SecondReader["VendorAddressAddr2"] != DBNull.Value ? SecondReader["VendorAddressAddr2"].ToString() : string.Empty;
                                                newCheckItem.Addr3 = SecondReader["VendorAddressAddr3"] != DBNull.Value ? SecondReader["VendorAddressAddr3"].ToString() : string.Empty;
                                                newCheckItem.Addr4 = SecondReader["VendorAddressAddr4"] != DBNull.Value ? SecondReader["VendorAddressAddr4"].ToString() : string.Empty;
                                                newCheckItem.AddrCity = SecondReader["VendorAddressCity"] != DBNull.Value ? SecondReader["VendorAddressCity"].ToString() : string.Empty;
                                            }
                                        }
                                    }

                                    SecondConnection.Close();
                                }
                                ItemReceipt.Add(newCheckItem);
                            }
                        }
                    }

                    string expenseQuery = @"SELECT TOP 1000
                                              ItemReceipt.VendorRefFullName,
                                              ItemReceipt.VendorRefListID,
                                              ItemReceipt.TxnID AS ItemReceiptTxnID,
                                              ItemReceipt.TxnDate,
                                              ItemReceipt.RefNumber,
                                              ItemReceipt.Memo,
                                              ItemReceipt.TotalAmount,
                                              ItemReceipt.APAccountRefFullName,
                                              ItemReceipt.APAccountRefListID,
                                              ItemReceiptExpenseLine.TxnID AS ExpenseLineTxnID,
                                              ItemReceiptExpenseLine.ExpenseLineAccountRefFullName,
                                              ItemReceiptExpenseLine.ExpenseLineAmount,
                                              ItemReceiptExpenseLine.ExpenseLineMemo
                                    FROM
                                              ItemReceipt
                                    INNER JOIN 
                                              ItemReceiptExpenseLine ON ItemReceipt.TxnID = ItemReceiptExpenseLine.TxnID
                                    WHERE
                                              ItemReceipt.RefNumber = ? ";


                    using (OleDbCommand expenseCommand = new OleDbCommand(expenseQuery, accessConnection))
                    {
                        expenseCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader expenseReader = expenseCommand.ExecuteReader())
                        {
                            while (expenseReader.Read())
                            {
                                ItemReciept newItemRecieptExpense = new ItemReciept
                                {
                                    DateCreated = expenseReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(expenseReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = expenseReader["APAccountRefFullName"] != DBNull.Value ? expenseReader["APAccountRefFullName"].ToString() : string.Empty,
                                    PayeeFullName = expenseReader["VendorRefFullName"] != DBNull.Value ? expenseReader["VendorRefFullName"].ToString() : string.Empty,
                                    RefNumber = expenseReader["RefNumber"] != DBNull.Value ? expenseReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = expenseReader["TotalAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["TotalAmount"]) : 0.0,
                                    Memo = expenseReader["Memo"] != DBNull.Value ? expenseReader["Memo"].ToString() : string.Empty,

                                    Account = expenseReader["ExpenseLineAccountRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineAccountRefFullName"].ToString() : string.Empty,
                                    ExpensesAmount = expenseReader["ExpenseLineAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["ExpenseLineAmount"]) : 0.0,
                                    ExpensesMemo = expenseReader["ExpenseLineMemo"] != DBNull.Value ? expenseReader["ExpenseLineMemo"].ToString() : string.Empty,
                                    ReceiptItemType = ReceiptItemType.RecieptExpense,

                                    //IncrementalID = nextID, // Assign the CV000001 here
                                    //IncrementalID = nextID.ToString("D6"),
                                };
                                string SecondQuery = @"SELECT 
                                        VendorAddressAddr1, VendorAddressAddr2, VendorAddressAddr3, VendorAddressAddr4, VendorAddressCity
                                     FROM 
                                         Vendor
                                     WHERE 
                                         Vendor.ListID = ?";
                                using (OleDbConnection SecondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    SecondConnection.Open();

                                    using (OleDbCommand SecondCommand = new OleDbCommand(SecondQuery, SecondConnection))
                                    {
                                        SecondCommand.Parameters.AddWithValue("Vendor.ListID", OleDbType.VarChar).Value = expenseReader["VendorRefListID"];

                                        using (OleDbDataReader SecondReader = SecondCommand.ExecuteReader())
                                        {
                                            while (SecondReader.Read())
                                            {
                                                newItemRecieptExpense.Addr1 = SecondReader["VendorAddressAddr1"] != DBNull.Value ? SecondReader["VendorAddressAddr1"].ToString() : string.Empty;
                                                newItemRecieptExpense.Addr2 = SecondReader["VendorAddressAddr2"] != DBNull.Value ? SecondReader["VendorAddressAddr2"].ToString() : string.Empty;
                                                newItemRecieptExpense.Addr3 = SecondReader["VendorAddressAddr3"] != DBNull.Value ? SecondReader["VendorAddressAddr3"].ToString() : string.Empty;
                                                newItemRecieptExpense.Addr4 = SecondReader["VendorAddressAddr4"] != DBNull.Value ? SecondReader["VendorAddressAddr4"].ToString() : string.Empty;
                                                newItemRecieptExpense.AddrCity = SecondReader["VendorAddressCity"] != DBNull.Value ? SecondReader["VendorAddressCity"].ToString() : string.Empty;
                                            }
                                        }
                                    }

                                    SecondConnection.Close();
                                }

                                ItemReceipt.Add(newItemRecieptExpense);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ItemReceipt;
        }
        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_LEADS(string refNumber)
        {
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_CV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string itemQuery = "SELECT TOP 1000 Check.TxnDate, " +
                        "Check.AccountRefFullName, " +
                        "Check.PayeeEntityRefFullName, " +
                        "Check.RefNumber, " +
                        "Check.Amount, " +
                        "Check.AddressAddr1, " +
                        "Check.AddressAddr2, " +
                        "Check.Memo, " +
                        "CheckItemLine.ItemLineItemRefFullName, " +
                        "CheckItemLine.ItemLineDesc, " +
                        "CheckItemLine.ItemLineClassRefFullName, " +
                        "CheckItemLine.ItemLineItemRefListID, " +
                        "CheckItemLine.ItemLineAmount, " +
                        "CheckItemLine.PayeeEntityReflistID " +
                        "FROM [Check] " +
                        "INNER JOIN CheckItemLine ON [Check].RefNumber = CheckItemLine.RefNumber " +
                        "WHERE [Check].RefNumber = ? ";

                    using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, accessConnection))
                    {
                        itemCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader itemReader = itemCommand.ExecuteReader())
                        {
                            while (itemReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckItem = new CheckTableExpensesAndItems
                                {
                                    DateCreated = itemReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(itemReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = itemReader["AccountRefFullname"] != DBNull.Value ? itemReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = itemReader["PayeeEntityRefFullName"] != DBNull.Value ? itemReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = itemReader["RefNumber"] != DBNull.Value ? itemReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = itemReader["Amount"] != DBNull.Value ? Convert.ToDouble(itemReader["Amount"]) : 0.0,
                                    Address = itemReader["AddressAddr1"] != DBNull.Value ? itemReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = itemReader["AddressAddr2"] != DBNull.Value ? itemReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = itemReader["Memo"] != DBNull.Value ? itemReader["Memo"].ToString() : string.Empty,

                                    Item = itemReader["ItemLineItemRefFullName"] != DBNull.Value ? itemReader["ItemLineItemRefFullName"].ToString() : string.Empty,
                                    ItemDescription = itemReader["ItemLineDesc"] != DBNull.Value ? itemReader["ItemLineDesc"].ToString() : string.Empty,
                                    ItemClass = itemReader["ItemLineClassRefFullName"] != DBNull.Value ? itemReader["ItemLineClassRefFullName"].ToString() : string.Empty,
                                    ItemAmount = itemReader["ItemLineAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineAmount"]) : 0.0,
                                    ItemType = ItemType.Item,

                                    //IncrementalID = nextID,
                                    IncrementalID = nextID.ToString("D6")
                                };
                                string secondQuery = @"SELECT 
                                                        AssetAccountRefFullname
                                                    FROM 
                                                        Item
                                                    WHERE 
                                                        Item.listID = ?";
                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("Item.ListID", OleDbType.VarChar).Value = itemReader["ItemLineItemRefListID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                newCheckItem.AccountName = secondReader["AssetAccountRefFullname"] != DBNull.Value ? secondReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                //newCheckItem.AccountNumber = secondReader["AccountNumber"] != DBNull.Value ? secondReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }
                                checks.Add(newCheckItem);
                            }
                        }
                    }

                    string expenseQuery = "SELECT TOP 1000 Check.TxnDate, Check.AccountRefFullName, " +
                                      "Check.PayeeEntityRefFullName, Check.RefNumber, Check.Amount, " +
                                      "Check.AddressAddr1, Check.Memo, " +
                                      "Check.AddressAddr2," +
                                      "CheckExpenseLine.ExpenseLineAccountRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineAccountRefListID, " +
                                      "CheckExpenseLine.ExpenseLineAmount, CheckExpenseLine.ExpenseLineMemo, " +
                                      "CheckExpenseLine.ExpenseLineCustomerRefFullName, " +
                                      "CheckExpenseLine.PayeeEntityReflistID " +
                                      "FROM [Check] " +
                                      "INNER JOIN CheckExpenseLine ON [Check].RefNumber = CheckExpenseLine.RefNumber " +
                                      //"WHERE Check.RefNumber = ? AND Check.TimeCreated >= ? AND Check.TimeCreated < ?";
                                      "WHERE [Check].RefNumber = ?";


                    using (OleDbCommand expenseCommand = new OleDbCommand(expenseQuery, accessConnection))
                    {
                        expenseCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader expenseReader = expenseCommand.ExecuteReader())
                        {
                            while (expenseReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckExpense = new CheckTableExpensesAndItems
                                {
                                    DateCreated = expenseReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(expenseReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = expenseReader["AccountRefFullname"] != DBNull.Value ? expenseReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = expenseReader["PayeeEntityRefFullName"] != DBNull.Value ? expenseReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = expenseReader["RefNumber"] != DBNull.Value ? expenseReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = expenseReader["Amount"] != DBNull.Value ? Convert.ToDouble(expenseReader["Amount"]) : 0.0,
                                    Address = expenseReader["AddressAddr1"] != DBNull.Value ? expenseReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = expenseReader["AddressAddr2"] != DBNull.Value ? expenseReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = expenseReader["Memo"] != DBNull.Value ? expenseReader["Memo"].ToString() : string.Empty,

                                    Account = expenseReader["ExpenseLineAccountRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineAccountRefFullName"].ToString() : string.Empty,
                                    ExpensesAmount = expenseReader["ExpenseLineAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["ExpenseLineAmount"]) : 0.0,
                                    ExpensesMemo = expenseReader["ExpenseLineMemo"] != DBNull.Value ? expenseReader["ExpenseLineMemo"].ToString() : string.Empty,
                                    ExpensesCustomerJob = expenseReader["ExpenseLineCustomerRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineCustomerRefFullName"].ToString() : string.Empty,
                                    ItemType = ItemType.Expense,

                                    //IncrementalID = nextID, // Assign the CV000001 here
                                    IncrementalID = nextID.ToString("D6"),
                                };
                                checks.Add(newCheckExpense);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return checks;
        } // CV Check Expense Item

        /*public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_KAYAK(string refNumber)
        {
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_CV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string itemQuery = "SELECT TOP 1000 Check.TxnDate, " +
                        "Check.AccountRefFullName, " +
                        "Check.AccountRefListID, " +
                        "Check.PayeeEntityRefFullName, " +
                        "Check.RefNumber, " +
                        "Check.Amount, " +
                        "Check.AddressAddr1, " +
                        "Check.AddressAddr2, " +
                        "Check.Memo, " +
                        "CheckItemLine.ItemLineItemRefFullName, " +
                        "CheckItemLine.ItemLineDesc, " +
                        "CheckItemLine.ItemLineClassRefFullName, " +
                        "CheckItemLine.ItemLineItemRefListID, " +
                        "CheckItemLine.ItemLineAmount, " +
                        "CheckItemLine.PayeeEntityReflistID " +
                        "FROM [Check] " +
                        "INNER JOIN CheckItemLine ON [Check].RefNumber = CheckItemLine.RefNumber " +
                        "WHERE [Check].RefNumber = ? ";

                    using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, accessConnection))
                    {
                        itemCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader itemReader = itemCommand.ExecuteReader())
                        {
                            while (itemReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckItem = new CheckTableExpensesAndItems
                                {
                                    DateCreated = itemReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(itemReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = itemReader["AccountRefFullname"] != DBNull.Value ? itemReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = itemReader["PayeeEntityRefFullName"] != DBNull.Value ? itemReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = itemReader["RefNumber"] != DBNull.Value ? itemReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = itemReader["Amount"] != DBNull.Value ? Convert.ToDouble(itemReader["Amount"]) : 0.0,
                                    Address = itemReader["AddressAddr1"] != DBNull.Value ? itemReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = itemReader["AddressAddr2"] != DBNull.Value ? itemReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = itemReader["Memo"] != DBNull.Value ? itemReader["Memo"].ToString() : string.Empty,

                                    Item = itemReader["ItemLineItemRefFullName"] != DBNull.Value ? itemReader["ItemLineItemRefFullName"].ToString() : string.Empty,
                                    ItemDescription = itemReader["ItemLineDesc"] != DBNull.Value ? itemReader["ItemLineDesc"].ToString() : string.Empty,
                                    ItemClass = itemReader["ItemLineClassRefFullName"] != DBNull.Value ? itemReader["ItemLineClassRefFullName"].ToString() : string.Empty,
                                    ItemAmount = itemReader["ItemLineAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineAmount"]) : 0.0,
                                    ItemType = ItemType.Item,

                                    //IncrementalID = nextID,
                                    IncrementalID = nextID.ToString("D6"),


                                };
                                string bankAccountRefListID = itemReader["AccountRefListID"] != DBNull.Value ? itemReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckItem.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }

                                string secondQuery = @"SELECT Name, AssetAccountRefFullname, AssetAccountRefListID FROM Item WHERE ListID = ?";
                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();
                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = itemReader["ItemLineItemRefListID"];
                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                newCheckItem.AccountName = secondReader["AssetAccountRefFullname"] != DBNull.Value ? secondReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                newCheckItem.ItemName = secondReader["Name"] != DBNull.Value ? secondReader["Name"].ToString() : string.Empty;
                                                string assetAccountRefListID = secondReader["AssetAccountRefListID"] != DBNull.Value ? secondReader["AssetAccountRefListID"].ToString() : string.Empty;

                                                if (!string.IsNullOrEmpty(assetAccountRefListID))
                                                {
                                                    // Get AccountNumber from Account table using AssetAccountRefListID
                                                    string getAssetAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                                    using (OleDbCommand accCmd = new OleDbCommand(getAssetAccountNumberQuery, secondConnection))
                                                    {
                                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = assetAccountRefListID;
                                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                                        {
                                                            while (accReader.Read())
                                                            {
                                                                newCheckItem.AssetAccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                                newCheckItem.AssetAccountName = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    secondConnection.Close();
                                }
                                checks.Add(newCheckItem);
                            }
                        }
                    }

                    string expenseQuery = "SELECT TOP 1000 Check.TxnDate, Check.AccountRefListID, Check.AccountRefFullName, " +
                                      "Check.PayeeEntityRefFullName, Check.RefNumber, Check.Amount, " +
                                      "Check.AddressAddr1, Check.Memo, " +
                                      "Check.AddressAddr2," +
                                      "CheckExpenseLine.ExpenseLineAccountRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineClassRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineAccountRefListID, " +
                                      "CheckExpenseLine.ExpenseLineAmount, CheckExpenseLine.ExpenseLineMemo, " +
                                      "CheckExpenseLine.ExpenseLineCustomerRefFullName, " +
                                      "CheckExpenseLine.PayeeEntityReflistID " +
                                      "FROM [Check] " +
                                      "INNER JOIN CheckExpenseLine ON [Check].RefNumber = CheckExpenseLine.RefNumber " +
                                      //"WHERE Check.RefNumber = ? AND Check.TimeCreated >= ? AND Check.TimeCreated < ?";
                                      "WHERE [Check].RefNumber = ?";


                    using (OleDbCommand expenseCommand = new OleDbCommand(expenseQuery, accessConnection))
                    {
                        expenseCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader expenseReader = expenseCommand.ExecuteReader())
                        {
                            while (expenseReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckExpense = new CheckTableExpensesAndItems
                                {
                                    DateCreated = expenseReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(expenseReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = expenseReader["AccountRefFullname"] != DBNull.Value ? expenseReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = expenseReader["PayeeEntityRefFullName"] != DBNull.Value ? expenseReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = expenseReader["RefNumber"] != DBNull.Value ? expenseReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = expenseReader["Amount"] != DBNull.Value ? Convert.ToDouble(expenseReader["Amount"]) : 0.0,
                                    Address = expenseReader["AddressAddr1"] != DBNull.Value ? expenseReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = expenseReader["AddressAddr2"] != DBNull.Value ? expenseReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = expenseReader["Memo"] != DBNull.Value ? expenseReader["Memo"].ToString() : string.Empty,

                                    Account = expenseReader["ExpenseLineAccountRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineAccountRefFullName"].ToString() : string.Empty,
                                    ExpenseClass = expenseReader["ExpenseLineClassRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineClassRefFullName"].ToString() : string.Empty,
                                    ExpensesAmount = expenseReader["ExpenseLineAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["ExpenseLineAmount"]) : 0.0,
                                    ExpensesMemo = expenseReader["ExpenseLineMemo"] != DBNull.Value ? expenseReader["ExpenseLineMemo"].ToString() : string.Empty,
                                    ExpensesCustomerJob = expenseReader["ExpenseLineCustomerRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineCustomerRefFullName"].ToString() : string.Empty,
                                    ItemType = ItemType.Expense,

                                    //IncrementalID = nextID, // Assign the CV000001 here
                                    IncrementalID = nextID.ToString("D6"),
                                };
                                string bankAccountRefListID = expenseReader["AccountRefListID"] != DBNull.Value ? expenseReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckExpense.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                string getExpenseAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                using (OleDbConnection accountConn = new OleDbConnection(accessConnectionString))
                                {
                                    accountConn.Open();
                                    using (OleDbCommand accCmd = new OleDbCommand(getExpenseAccountNumberQuery, accountConn))
                                    {
                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = expenseReader["ExpenseLineAccountRefListID"];
                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                        {
                                            while (accReader.Read())
                                            {
                                                newCheckExpense.AccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                newCheckExpense.AccountNameCheck = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                checks.Add(newCheckExpense);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return checks;
        } // CV Check Expense Item*/

        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_KAYAK(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            try
            {
                string accessConnectionString = GetAccessConnectionString();
                string nextIDStr = GetNextIncrementalID_CV(accessConnectionString).ToString("D6");

                sessionManager.OpenConnection2("", "KAYAK Check Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery checkQuery = request.AppendCheckQueryRq();
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);
                checkQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);

                ICheckRetList list = qbResponse.Detail as ICheckRetList;
                if (list == null || list.Count == 0) return checks;

                for (int i = 0; i < list.Count; i++)
                {
                    ICheckRet check = list.GetAt(i);

                    // HEADER DATA
                    string bankAccountName = check.AccountRef?.FullName?.GetValue() ?? "";
                    string bankAccountListID = check.AccountRef?.ListID?.GetValue() ?? "";
                    // Fetch Bank Account Number
                    string bankAccountNumber = GetAccountNumber(sessionManager, bankAccountListID);

                    // EXPENSE LINES
                    if (check.ExpenseLineRetList != null)
                    {
                        for (int e = 0; e < check.ExpenseLineRetList.Count; e++)
                        {
                            IExpenseLineRet exp = check.ExpenseLineRetList.GetAt(e);
                            string expAccListID = exp.AccountRef?.ListID?.GetValue() ?? "";

                            checks.Add(new CheckTableExpensesAndItems
                            {
                                IncrementalID = nextIDStr,
                                DateCreated = check.TxnDate?.GetValue() ?? DateTime.MinValue,
                                BankAccount = bankAccountName,
                                BankAccountNumber = bankAccountNumber,
                                PayeeFullName = check.PayeeEntityRef?.FullName?.GetValue() ?? "",
                                RefNumber = check.RefNumber?.GetValue() ?? "",
                                TotalAmount = check.Amount?.GetValue() ?? 0,

                                // Line Specifics
                                ExpenseClass = exp.ClassRef?.FullName?.GetValue() ?? "",
                                Account = exp.AccountRef?.FullName?.GetValue() ?? "",
                                AccountNumber = GetAccountNumber(sessionManager, expAccListID),
                                ExpensesAmount = exp.Amount?.GetValue() ?? 0,
                                ItemType = ItemType.Expense
                            });
                        }
                    }

                    // ITEM LINES
                    if (check.ORItemLineRetList != null)
                    {
                        for (int iLine = 0; iLine < check.ORItemLineRetList.Count; iLine++)
                        {
                            IORItemLineRet orItemLine = (IORItemLineRet)check.ORItemLineRetList.GetAt(iLine);
                            if (orItemLine.ItemLineRet != null)
                            {
                                IItemLineRet item = orItemLine.ItemLineRet;
                                // Items usually map to an Income/Expense account internally
                                // If you need the Asset Account Number for the Item:
                                string itemAccNumber = GetItemAssetAccountNumber(sessionManager, item.ItemRef?.ListID?.GetValue());

                                checks.Add(new CheckTableExpensesAndItems
                                {
                                    IncrementalID = nextIDStr,
                                    DateCreated = check.TxnDate?.GetValue() ?? DateTime.MinValue,
                                    BankAccount = bankAccountName,
                                    BankAccountNumber = bankAccountNumber,
                                    PayeeFullName = check.PayeeEntityRef?.FullName?.GetValue() ?? "",
                                    RefNumber = check.RefNumber?.GetValue() ?? "",

                                    // Item Specifics
                                    ItemClass = item.ClassRef?.FullName?.GetValue() ?? "",
                                    Item = item.ItemRef?.FullName?.GetValue() ?? "",
                                    AssetAccountNumber = itemAccNumber,
                                    ItemAmount = item.Amount?.GetValue() ?? 0,
                                    ItemType = ItemType.Item
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { sessionManager.EndSession(); sessionManager.CloseConnection(); }

            return checks;
        }

        private string GetAccountNumber(QBSessionManager sessionManager, string listID)
        {
            if (string.IsNullOrEmpty(listID)) return "";

            IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
            IAccountQuery accQuery = request.AppendAccountQueryRq();
            accQuery.ORAccountListQuery.ListIDList.Add(listID);

            IMsgSetResponse response = sessionManager.DoRequests(request);
            IResponse qbResponse = response.ResponseList.GetAt(0);
            IAccountRetList accList = qbResponse.Detail as IAccountRetList;

            return accList?.GetAt(0)?.AccountNumber?.GetValue() ?? "";
        }

        private string GetItemAssetAccountNumber(QBSessionManager sessionManager, string itemListID)
        {
            if (string.IsNullOrEmpty(itemListID)) return "";

            IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
            IItemQuery itemQuery = request.AppendItemQueryRq();
            itemQuery.ORListQuery.ListIDList.Add(itemListID);

            IMsgSetResponse response = sessionManager.DoRequests(request);
            IResponse qbResponse = response.ResponseList.GetAt(0);

            // This is a simplified check for AssetAccountRef within the Item details
            // You may need to cast to IItemInventoryRet depending on your Item Types
            return ""; // Logic similar to GetAccountNumber but targeting ItemRet.AssetAccountRef
        }

        /*public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_IVP(string refNumber)
        {
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_CV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string itemQuery = "SELECT TOP 1000 Check.TxnDate, " +
                        "Check.AccountRefFullName, " +
                        "Check.AccountRefListID, " +
                        "Check.PayeeEntityRefFullName, " +
                        "Check.RefNumber, " +
                        "Check.Amount, " +
                        "Check.AddressAddr1, " +
                        "Check.AddressAddr2, " +
                        "Check.Memo, " +
                        "CheckItemLine.ItemLineItemRefFullName, " +
                        "CheckItemLine.ItemLineDesc, " +
                        "CheckItemLine.ItemLineClassRefFullName, " +
                        "CheckItemLine.ItemLineItemRefListID, " +
                        "CheckItemLine.ItemLineAmount, " +
                        "CheckItemLine.PayeeEntityReflistID " +
                        "FROM [Check] " +
                        "INNER JOIN CheckItemLine ON [Check].RefNumber = CheckItemLine.RefNumber " +
                        "WHERE [Check].RefNumber = ? ";

                    using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, accessConnection))
                    {
                        itemCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader itemReader = itemCommand.ExecuteReader())
                        {
                            while (itemReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckItem = new CheckTableExpensesAndItems
                                {
                                    DateCreated = itemReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(itemReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = itemReader["AccountRefFullname"] != DBNull.Value ? itemReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = itemReader["PayeeEntityRefFullName"] != DBNull.Value ? itemReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = itemReader["RefNumber"] != DBNull.Value ? itemReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = itemReader["Amount"] != DBNull.Value ? Convert.ToDouble(itemReader["Amount"]) : 0.0,
                                    Address = itemReader["AddressAddr1"] != DBNull.Value ? itemReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = itemReader["AddressAddr2"] != DBNull.Value ? itemReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = itemReader["Memo"] != DBNull.Value ? itemReader["Memo"].ToString() : string.Empty,

                                    Item = itemReader["ItemLineItemRefFullName"] != DBNull.Value ? itemReader["ItemLineItemRefFullName"].ToString() : string.Empty,
                                    ItemDescription = itemReader["ItemLineDesc"] != DBNull.Value ? itemReader["ItemLineDesc"].ToString() : string.Empty,
                                    ItemClass = itemReader["ItemLineClassRefFullName"] != DBNull.Value ? itemReader["ItemLineClassRefFullName"].ToString() : string.Empty,
                                    ItemAmount = itemReader["ItemLineAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineAmount"]) : 0.0,
                                    ItemType = ItemType.Item,

                                    //IncrementalID = nextID,
                                    IncrementalID = nextID.ToString("D6"),


                                };
                                string bankAccountRefListID = itemReader["AccountRefListID"] != DBNull.Value ? itemReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckItem.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }

                                string secondQuery = @"SELECT Name, AssetAccountRefFullname, AssetAccountRefListID FROM Item WHERE ListID = ?";
                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();
                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = itemReader["ItemLineItemRefListID"];
                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                newCheckItem.AccountName = secondReader["AssetAccountRefFullname"] != DBNull.Value ? secondReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                newCheckItem.ItemName = secondReader["Name"] != DBNull.Value ? secondReader["Name"].ToString() : string.Empty;
                                                string assetAccountRefListID = secondReader["AssetAccountRefListID"] != DBNull.Value ? secondReader["AssetAccountRefListID"].ToString() : string.Empty;

                                                if (!string.IsNullOrEmpty(assetAccountRefListID))
                                                {
                                                    // Get AccountNumber from Account table using AssetAccountRefListID
                                                    string getAssetAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                                    using (OleDbCommand accCmd = new OleDbCommand(getAssetAccountNumberQuery, secondConnection))
                                                    {
                                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = assetAccountRefListID;
                                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                                        {
                                                            while (accReader.Read())
                                                            {
                                                                newCheckItem.AssetAccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                                newCheckItem.AssetAccountName = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    secondConnection.Close();
                                }
                                checks.Add(newCheckItem);
                            }
                        }
                    }

                    string expenseQuery = "SELECT TOP 1000 Check.TxnDate, Check.AccountRefListID, Check.AccountRefFullName, " +
                                      "Check.PayeeEntityRefFullName, Check.RefNumber, Check.Amount, " +
                                      "Check.AddressAddr1, Check.Memo, " +
                                      "Check.AddressAddr2," +
                                      "CheckExpenseLine.ExpenseLineAccountRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineClassRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineAccountRefListID, " +
                                      "CheckExpenseLine.ExpenseLineAmount, CheckExpenseLine.ExpenseLineMemo, " +
                                      "CheckExpenseLine.ExpenseLineCustomerRefFullName, " +
                                      "CheckExpenseLine.PayeeEntityReflistID " +
                                      "FROM [Check] " +
                                      "INNER JOIN CheckExpenseLine ON [Check].RefNumber = CheckExpenseLine.RefNumber " +
                                      //"WHERE Check.RefNumber = ? AND Check.TimeCreated >= ? AND Check.TimeCreated < ?";
                                      "WHERE [Check].RefNumber = ?";


                    using (OleDbCommand expenseCommand = new OleDbCommand(expenseQuery, accessConnection))
                    {
                        expenseCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader expenseReader = expenseCommand.ExecuteReader())
                        {
                            while (expenseReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckExpense = new CheckTableExpensesAndItems
                                {
                                    DateCreated = expenseReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(expenseReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = expenseReader["AccountRefFullname"] != DBNull.Value ? expenseReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = expenseReader["PayeeEntityRefFullName"] != DBNull.Value ? expenseReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = expenseReader["RefNumber"] != DBNull.Value ? expenseReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = expenseReader["Amount"] != DBNull.Value ? Convert.ToDouble(expenseReader["Amount"]) : 0.0,
                                    Address = expenseReader["AddressAddr1"] != DBNull.Value ? expenseReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = expenseReader["AddressAddr2"] != DBNull.Value ? expenseReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = expenseReader["Memo"] != DBNull.Value ? expenseReader["Memo"].ToString() : string.Empty,

                                    Account = expenseReader["ExpenseLineAccountRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineAccountRefFullName"].ToString() : string.Empty,
                                    ExpenseClass = expenseReader["ExpenseLineClassRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineClassRefFullName"].ToString() : string.Empty,
                                    ExpensesAmount = expenseReader["ExpenseLineAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["ExpenseLineAmount"]) : 0.0,
                                    ExpensesMemo = expenseReader["ExpenseLineMemo"] != DBNull.Value ? expenseReader["ExpenseLineMemo"].ToString() : string.Empty,
                                    ExpensesCustomerJob = expenseReader["ExpenseLineCustomerRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineCustomerRefFullName"].ToString() : string.Empty,
                                    ItemType = ItemType.Expense,

                                    //IncrementalID = nextID, // Assign the CV000001 here
                                    IncrementalID = nextID.ToString("D6"),
                                };
                                string bankAccountRefListID = expenseReader["AccountRefListID"] != DBNull.Value ? expenseReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckExpense.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                string getExpenseAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                using (OleDbConnection accountConn = new OleDbConnection(accessConnectionString))
                                {
                                    accountConn.Open();
                                    using (OleDbCommand accCmd = new OleDbCommand(getExpenseAccountNumberQuery, accountConn))
                                    {
                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = expenseReader["ExpenseLineAccountRefListID"];
                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                        {
                                            while (accReader.Read())
                                            {
                                                newCheckExpense.AccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                newCheckExpense.AccountNameCheck = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                checks.Add(newCheckExpense);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return checks;
        }*/

        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_IVP(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            try
            {
                Console.WriteLine("--- Starting QuickBooks Session ---");
                string AppName = "QuickBooks Check Retrieval";
                sessionManager.OpenConnection2("", AppName, ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Build request
                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery checkQuery = request.AppendCheckQueryRq();

                // Filter by RefNumber
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion
                    .SetValue(ENMatchCriterion.mcStartsWith);

                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber
                    .SetValue(refNumber);

                // Include line items
                checkQuery.IncludeLineItems.SetValue(true);

                Console.WriteLine($"Querying for RefNumber starting with: {refNumber}");
                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);

                ICheckRetList list = qbResponse.Detail as ICheckRetList;

                if (list == null || list.Count == 0)
                {
                    Console.WriteLine("No checks found.");
                    return checks;
                }

                Console.WriteLine($"Found {list.Count} check(s).");

                for (int i = 0; i < list.Count; i++)
                {
                    ICheckRet check = list.GetAt(i);

                    // HEADER DATA
                    DateTime txnDate = check.TxnDate?.GetValue() ?? DateTime.MinValue;
                    string bankAccount = check.AccountRef?.FullName?.GetValue() ?? "";
                    string payee = check.PayeeEntityRef?.FullName?.GetValue() ?? "";
                    string memo = check.Memo?.GetValue() ?? "";
                    string address1 = check.Address?.Addr1?.GetValue() ?? "";
                    string address2 = check.Address?.Addr2?.GetValue() ?? "";
                    double totalAmount = check.Amount?.GetValue() ?? 0;
                    string currentRef = check.RefNumber?.GetValue() ?? "";
                    string duedate = check.TxnDate?.GetValue().ToString("yyyy-MM-dd") ?? "";

                    Console.WriteLine($"\n[Check #{i + 1}] Ref: {currentRef} | Payee: {payee} | Total: {totalAmount}");

                    // EXPENSE LINES
                    if (check.ExpenseLineRetList != null)
                    {
                        for (int e = 0; e < check.ExpenseLineRetList.Count; e++)
                        {
                            IExpenseLineRet exp = check.ExpenseLineRetList.GetAt(e);

                            string expAccount = exp.AccountRef?.FullName?.GetValue() ?? "";
                            double expAmount = exp.Amount?.GetValue() ?? 0;

                            Console.WriteLine($"   -> [Expense Line] Account: {expAccount} | Amount: {expAmount}");

                            checks.Add(new CheckTableExpensesAndItems
                            {
                                DateCreated = txnDate,
                                BankAccount = bankAccount,
                                PayeeFullName = payee,
                                RefNumber = refNumber,
                                TotalAmount = totalAmount,
                                DueDate = txnDate,
                                Memo = memo,
                                Address = address1,
                                Address2 = address2,

                                Account = expAccount,
                                ExpenseClass = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpensesAmount = expAmount,
                                ExpensesMemo = exp.Memo?.GetValue() ?? "",
                                ExpensesCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",

                                ItemType = ItemType.Expense
                            });
                        }
                    }

                    // ITEM LINES
                    if (check.ORItemLineRetList != null)
                    {
                        for (int iLine = 0; iLine < check.ORItemLineRetList.Count; iLine++)
                        {
                            // 1. Cast to the "OR" wrapper first
                            IORItemLineRet orItemLine = (IORItemLineRet)check.ORItemLineRetList.GetAt(iLine);

                            // 2. Check if the wrapper contains a standard ItemLineRet
                            if (orItemLine.ItemLineRet != null)
                            {
                                IItemLineRet item = orItemLine.ItemLineRet;

                                string itemName = item.ItemRef?.FullName?.GetValue() ?? "";
                                double itemAmount = item.Amount?.GetValue() ?? 0;

                                Console.WriteLine($"   -> [Item Line] Item: {itemName} | Amount: {itemAmount}");

                                checks.Add(new CheckTableExpensesAndItems
                                {
                                    DateCreated = txnDate,
                                    BankAccount = bankAccount,
                                    PayeeFullName = payee,
                                    RefNumber = refNumber,
                                    TotalAmount = totalAmount,
                                    DueDate = txnDate,
                                    Memo = memo,
                                    Address = address1,
                                    Address2 = address2,

                                    Item = itemName,
                                    ItemDescription = item.Desc?.GetValue() ?? "",
                                    ItemClass = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemAmount = itemAmount,

                                    ItemType = ItemType.Item
                                });
                            }
                            else if (orItemLine.ItemGroupLineRet != null)
                            {
                                Console.WriteLine("   -> [Item Group] Found a Group/Bundle (Skipping logic not implemented)");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CRITICAL ERROR: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                Console.WriteLine("--- Closing Session ---");
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); }
                catch { }
            }

            return checks;
        }


        public List<JournalGridItem> GetJournalEntryForGrid(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<JournalGridItem> gridItems = new List<JournalGridItem>();

            try
            {
                Console.WriteLine("--- [START] DATA RETRIEVAL ---");

                sessionManager.OpenConnection2("", "QB Journal Grid", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery jeQuery = request.AppendJournalEntryQueryRq();

                // 1. QUERY BROADLY
                // We are forced to use mcStartsWith because your SDK lacks mcValues
                jeQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                jeQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);
                jeQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);
                IJournalEntryRetList list = qbResponse.Detail as IJournalEntryRetList;

                if (list != null)
                {
                    for (int i = 0; i < list.Count; i++)
                    {
                        IJournalEntryRet je = list.GetAt(i);
                        string docNum = je.RefNumber.GetValue();

                        // 2. FILTER STRICTLY (Manual Exact Match)
                        // If QuickBooks returns "JV0010" but we wanted "JV001", skip it.
                        if (docNum != refNumber)
                        {
                            continue;
                        }

                        // If we get here, it is the correct RefNumber. Extract lines.
                        DateTime date = je.TxnDate.GetValue();

                        if (je.ORJournalLineList != null)
                        {
                            for (int j = 0; j < je.ORJournalLineList.Count; j++)
                            {
                                IORJournalLine orLine = je.ORJournalLineList.GetAt(j);
                                JournalGridItem item = new JournalGridItem
                                {
                                    Date = date,
                                    Num = docNum,
                                    Type = "General Journal"
                                };

                                if (orLine.JournalDebitLine != null)
                                {
                                    var line = orLine.JournalDebitLine;
                                    item.AccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                                    item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                                    item.Memo = line.Memo?.GetValue() ?? "";
                                    item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                                    item.Debit = line.Amount?.GetValue() ?? 0;
                                    item.Credit = 0;
                                }
                                else if (orLine.JournalCreditLine != null)
                                {
                                    var line = orLine.JournalCreditLine;
                                    item.AccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                                    item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                                    item.Memo = line.Memo?.GetValue() ?? "";
                                    item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                                    item.Debit = 0;
                                    item.Credit = line.Amount?.GetValue() ?? 0;
                                }

                                gridItems.Add(item);
                            }
                        }

                        // 3. STOP IMMEDIATELY
                        // We found one "JV001". Even if there is a duplicate "JV001" later in the list, 
                        // we ignore it to prevent the Double Table issue.
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); } catch { }
            }

            return gridItems;
        }



        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_CPI(string refNumber)
        {
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_CV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    string itemQuery = "SELECT TOP 1000 Check.TxnDate, " +
                        "Check.AccountRefFullName, " +
                        "Check.AccountRefListID, " +
                        "Check.PayeeEntityRefFullName, " +
                        "Check.RefNumber, " +
                        "Check.Amount, " +
                        "Check.AddressAddr1, " +
                        "Check.AddressAddr2, " +
                        "Check.Memo, " +
                        "CheckItemLine.ItemLineItemRefFullName, " +
                        "CheckItemLine.ItemLineDesc, " +
                        "CheckItemLine.ItemLineClassRefFullName, " +
                        "CheckItemLine.ItemLineItemRefListID, " +
                        "CheckItemLine.ItemLineAmount, " +
                        "CheckItemLine.PayeeEntityReflistID " +
                        "FROM [Check] " +
                        "INNER JOIN CheckItemLine ON [Check].RefNumber = CheckItemLine.RefNumber " +
                        "WHERE [Check].RefNumber = ? ";

                    using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, accessConnection))
                    {
                        itemCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader itemReader = itemCommand.ExecuteReader())
                        {
                            while (itemReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckItem = new CheckTableExpensesAndItems
                                {
                                    DateCreated = itemReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(itemReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = itemReader["AccountRefFullname"] != DBNull.Value ? itemReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = itemReader["PayeeEntityRefFullName"] != DBNull.Value ? itemReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = itemReader["RefNumber"] != DBNull.Value ? itemReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = itemReader["Amount"] != DBNull.Value ? Convert.ToDouble(itemReader["Amount"]) : 0.0,
                                    Address = itemReader["AddressAddr1"] != DBNull.Value ? itemReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = itemReader["AddressAddr2"] != DBNull.Value ? itemReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = itemReader["Memo"] != DBNull.Value ? itemReader["Memo"].ToString() : string.Empty,

                                    Item = itemReader["ItemLineItemRefFullName"] != DBNull.Value ? itemReader["ItemLineItemRefFullName"].ToString() : string.Empty,
                                    ItemDescription = itemReader["ItemLineDesc"] != DBNull.Value ? itemReader["ItemLineDesc"].ToString() : string.Empty,
                                    ItemClass = itemReader["ItemLineClassRefFullName"] != DBNull.Value ? itemReader["ItemLineClassRefFullName"].ToString() : string.Empty,
                                    ItemAmount = itemReader["ItemLineAmount"] != DBNull.Value ? Convert.ToDouble(itemReader["ItemLineAmount"]) : 0.0,
                                    ItemType = ItemType.Item,

                                    //IncrementalID = nextID,
                                    IncrementalID = nextID.ToString("D6"),


                                };
                                string bankAccountRefListID = itemReader["AccountRefListID"] != DBNull.Value ? itemReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckItem.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }

                                string secondQuery = @"SELECT Name, AssetAccountRefFullname, AssetAccountRefListID FROM Item WHERE ListID = ?";
                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();
                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = itemReader["ItemLineItemRefListID"];
                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                newCheckItem.AccountName = secondReader["AssetAccountRefFullname"] != DBNull.Value ? secondReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                newCheckItem.ItemName = secondReader["Name"] != DBNull.Value ? secondReader["Name"].ToString() : string.Empty;
                                                string assetAccountRefListID = secondReader["AssetAccountRefListID"] != DBNull.Value ? secondReader["AssetAccountRefListID"].ToString() : string.Empty;

                                                if (!string.IsNullOrEmpty(assetAccountRefListID))
                                                {
                                                    // Get AccountNumber from Account table using AssetAccountRefListID
                                                    string getAssetAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                                    using (OleDbCommand accCmd = new OleDbCommand(getAssetAccountNumberQuery, secondConnection))
                                                    {
                                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = assetAccountRefListID;
                                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                                        {
                                                            while (accReader.Read())
                                                            {
                                                                newCheckItem.AssetAccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                                newCheckItem.AssetAccountName = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    secondConnection.Close();
                                }
                                checks.Add(newCheckItem);
                            }
                        }
                    }

                    string expenseQuery = "SELECT TOP 1000 Check.TxnDate, Check.AccountRefListID, Check.AccountRefFullName, " +
                                      "Check.PayeeEntityRefFullName, Check.RefNumber, Check.Amount, " +
                                      "Check.AddressAddr1, Check.Memo, " +
                                      "Check.AddressAddr2," +
                                      "CheckExpenseLine.ExpenseLineAccountRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineClassRefFullName, " +
                                      "CheckExpenseLine.ExpenseLineAccountRefListID, " +
                                      "CheckExpenseLine.ExpenseLineAmount, CheckExpenseLine.ExpenseLineMemo, " +
                                      "CheckExpenseLine.ExpenseLineCustomerRefFullName, " +
                                      "CheckExpenseLine.PayeeEntityReflistID " +
                                      "FROM [Check] " +
                                      "INNER JOIN CheckExpenseLine ON [Check].RefNumber = CheckExpenseLine.RefNumber " +
                                      //"WHERE Check.RefNumber = ? AND Check.TimeCreated >= ? AND Check.TimeCreated < ?";
                                      "WHERE [Check].RefNumber = ?";


                    using (OleDbCommand expenseCommand = new OleDbCommand(expenseQuery, accessConnection))
                    {
                        expenseCommand.Parameters.AddWithValue("Check.RefNumber", OdbcType.VarChar).Value = refNumber;
                        using (OleDbDataReader expenseReader = expenseCommand.ExecuteReader())
                        {
                            while (expenseReader.Read())
                            {
                                CheckTableExpensesAndItems newCheckExpense = new CheckTableExpensesAndItems
                                {
                                    DateCreated = expenseReader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(expenseReader["TxnDate"]).Date : DateTime.MinValue,
                                    BankAccount = expenseReader["AccountRefFullname"] != DBNull.Value ? expenseReader["AccountRefFullname"].ToString() : string.Empty,
                                    PayeeFullName = expenseReader["PayeeEntityRefFullName"] != DBNull.Value ? expenseReader["PayeeEntityRefFullName"].ToString() : string.Empty,
                                    RefNumber = expenseReader["RefNumber"] != DBNull.Value ? expenseReader["RefNumber"].ToString() : string.Empty,
                                    TotalAmount = expenseReader["Amount"] != DBNull.Value ? Convert.ToDouble(expenseReader["Amount"]) : 0.0,
                                    Address = expenseReader["AddressAddr1"] != DBNull.Value ? expenseReader["AddressAddr1"].ToString() : string.Empty,
                                    Address2 = expenseReader["AddressAddr2"] != DBNull.Value ? expenseReader["AddressAddr2"].ToString() : string.Empty,
                                    Memo = expenseReader["Memo"] != DBNull.Value ? expenseReader["Memo"].ToString() : string.Empty,

                                    Account = expenseReader["ExpenseLineAccountRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineAccountRefFullName"].ToString() : string.Empty,
                                    ExpenseClass = expenseReader["ExpenseLineClassRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineClassRefFullName"].ToString() : string.Empty,
                                    ExpensesAmount = expenseReader["ExpenseLineAmount"] != DBNull.Value ? Convert.ToDouble(expenseReader["ExpenseLineAmount"]) : 0.0,
                                    ExpensesMemo = expenseReader["ExpenseLineMemo"] != DBNull.Value ? expenseReader["ExpenseLineMemo"].ToString() : string.Empty,
                                    ExpensesCustomerJob = expenseReader["ExpenseLineCustomerRefFullName"] != DBNull.Value ? expenseReader["ExpenseLineCustomerRefFullName"].ToString() : string.Empty,
                                    ItemType = ItemType.Expense,

                                    //IncrementalID = nextID, // Assign the CV000001 here
                                    IncrementalID = nextID.ToString("D6"),
                                };
                                string bankAccountRefListID = expenseReader["AccountRefListID"] != DBNull.Value ? expenseReader["AccountRefListID"].ToString() : string.Empty;

                                if (!string.IsNullOrEmpty(bankAccountRefListID))
                                {
                                    string getBankAccountNumberQuery = @"SELECT AccountNumber FROM Account WHERE ListID = ?";
                                    using (OleDbCommand bankAccCmd = new OleDbCommand(getBankAccountNumberQuery, accessConnection))
                                    {
                                        bankAccCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = bankAccountRefListID;
                                        using (OleDbDataReader bankReader = bankAccCmd.ExecuteReader())
                                        {
                                            while (bankReader.Read())
                                            {
                                                newCheckExpense.BankAccountNumber = bankReader["AccountNumber"] != DBNull.Value ? bankReader["AccountNumber"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                string getExpenseAccountNumberQuery = @"SELECT AccountNumber, Name FROM Account WHERE ListID = ?";
                                using (OleDbConnection accountConn = new OleDbConnection(accessConnectionString))
                                {
                                    accountConn.Open();
                                    using (OleDbCommand accCmd = new OleDbCommand(getExpenseAccountNumberQuery, accountConn))
                                    {
                                        accCmd.Parameters.AddWithValue("ListID", OleDbType.VarChar).Value = expenseReader["ExpenseLineAccountRefListID"];
                                        using (OleDbDataReader accReader = accCmd.ExecuteReader())
                                        {
                                            while (accReader.Read())
                                            {
                                                newCheckExpense.AccountNumber = accReader["AccountNumber"] != DBNull.Value ? accReader["AccountNumber"].ToString() : string.Empty;
                                                newCheckExpense.AccountNameCheck = accReader["Name"] != DBNull.Value ? accReader["Name"].ToString() : string.Empty;
                                            }
                                        }
                                    }
                                }
                                checks.Add(newCheckExpense);
                            }
                        }
                    }
                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data to Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return checks;
        } // CV Check Expense Item
        public List<BillTable> GetAccountsPayableData_LEADS(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_APV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    // Retrieve data from Access database
                    string query = @"SELECT TOP 1000 
                       Bill.TxnDate,
                       Bill.RefNumber,
                       Bill.VendorRefFullName, 
                       Bill.TermsRefFullName,
                       Bill.DueDate, 
                       Bill.AmountDue, 
                       Bill.Memo,
                       Bill.IsPaid,
                       Bill.APAccountRefFullName,
                       Bill.APAccountRefListID,
                       Bill.TxnID
                   FROM Bill
                   WHERE Bill.RefNumber = ?";

                    using (OleDbCommand command = new OleDbCommand(query, accessConnection))
                    {
                        command.Parameters.AddWithValue("RefNumber", OleDbType.VarChar).Value = refNumber;
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                BillTable newBill = new BillTable
                                {
                                    // BillPaymentCheckLine table
                                    DateCreated = reader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["TxnDate"]).Date : DateTime.MinValue,
                                    RefNumber = reader["RefNumber"] != DBNull.Value ? reader["RefNumber"].ToString() : string.Empty,
                                    Vendor = reader["VendorRefFullName"] != DBNull.Value ? reader["VendorRefFullName"].ToString() : string.Empty,
                                    APAccountRefFullName = reader["APAccountRefFullName"] != DBNull.Value ? reader["APAccountRefFullName"].ToString() : string.Empty,
                                    TermsRefFullName = reader["TermsRefFullName"] != DBNull.Value ? reader["TermsRefFullName"].ToString() : string.Empty,
                                    DueDate = reader["DueDate"] != DBNull.Value ? Convert.ToDateTime(reader["DueDate"]).Date : DateTime.MinValue,
                                    AmountDue = reader["AmountDue"] != DBNull.Value ? Convert.ToDouble(reader["AmountDue"]) : 0.0,
                                    Memo = reader["Memo"] != DBNull.Value ? reader["Memo"].ToString() : string.Empty,
                                    IsPaid = reader["IsPaid"] != DBNull.Value ? Convert.ToBoolean(reader["IsPaid"]) : false,

                                    // Increment
                                    IncrementalID = nextID.ToString("D6")
                                };

                                string secondQuery = @"SELECT TOP 1000
                                        BillItemLine.ItemLineItemRefFullName AS AccountRefFullName, 
                                        BillItemLine.ItemLineAmount AS Amount,
                                        BillItemLine.ItemLineClassRefFullName AS ClassRefFullName,
                                        BillItemLine.ItemLineitemRefListID AS BillRefListID,
                                        BillItemLine.ItemLineDesc AS ItemExpenseMemo
                                    FROM 
                                        BillItemLine 
                                    WHERE 
                                        BillItemLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillItemLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];
                                        //secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineClassRefFullName = secondReader["ClassRefFullName"] != DBNull.Value ? secondReader["ClassRefFullName"].ToString() : string.Empty;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                Console.WriteLine($"Item: {itemLineItemRefFullName}, Class: {itemLineClassRefFullName}, Amount: {itemLineAmount}, Memo: {itemLineItemMemo}");


                                                //Console.WriteLine(itemLineAmount);
                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ItemLineItemRefFullName = itemLineItemRefFullName,
                                                    ItemLineAmount = itemLineAmount,
                                                    ItemLineClassRefFullName = itemLineClassRefFullName,
                                                    ItemLineMemo = itemLineItemMemo,
                                                });

                                                string fourthQuery = @"SELECT 
                                                        AssetAccountRefFullname
                                                    FROM 
                                                        item
                                                    WHERE 
                                                        item.listID = ?";

                                                string billRefListID = secondReader["BillRefListID"].ToString();

                                                using (OleDbConnection fourthConnection = new OleDbConnection(accessConnectionString))
                                                {
                                                    fourthConnection.Open();

                                                    using (OleDbCommand fourthCommand = new OleDbCommand(fourthQuery, fourthConnection))
                                                    {
                                                        fourthCommand.Parameters.AddWithValue("Account.ListID", OleDbType.VarChar).Value = billRefListID;

                                                        using (OleDbDataReader fourthReader = fourthCommand.ExecuteReader())
                                                        {
                                                            while (fourthReader.Read())
                                                            {
                                                                // Store additional account particulars in lists
                                                                string accountNameParticulars = fourthReader["AssetAccountRefFullname"] != DBNull.Value ? fourthReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                                //string accountNumberParticulars = fourthReader["AccountNumber"] != DBNull.Value ? fourthReader["AccountNumber"].ToString() : string.Empty;

                                                                newBill.AccountNameParticularsList.Add(accountNameParticulars);
                                                                //newBill.AccountNumberParticularsList.Add(accountNumberParticulars);
                                                            }
                                                        }
                                                    }

                                                    fourthConnection.Close();
                                                }
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }
                                string secondQuery2 = @"SELECT
                                        BillExpenseLine.ExpenseLineAccountRefFullName AS AccountRefFullName, 
                                        BillExpenseLine.ExpenseLineAmount AS Amount,
                                        BillExpenseLine.ExpenseLineClassRefFullName AS ClassRefFullName,
                                        BillExpenseLine.ExpenseLineAccountRefListID AS BillRefListID,
                                        BillExpenseLine.ExpenseLineMemo AS ItemExpenseMemo
                                    FROM 
                                        [BillExpenseLine]
                                    WHERE 
                                        BillExpenseLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery2, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineClassRefFullName = secondReader["ClassRefFullName"] != DBNull.Value ? secondReader["ClassRefFullName"].ToString() : string.Empty;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                Console.WriteLine($"Item: {itemLineItemRefFullName}, Amount: {itemLineAmount}, Memo: {itemLineItemMemo}");


                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ExpenseLineItemRefFullName = itemLineItemRefFullName,
                                                    ExpenseLineAmount = itemLineAmount,
                                                    ExpenseLineClassRefFullName = itemLineClassRefFullName,
                                                    ExpenseLineMemo = itemLineItemMemo,
                                                });
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }

                                string ThirdQuery = @"SELECT 
                                               Name, AccountNumber
                                            FROM 
                                                Account
                                            WHERE 
                                                Account.listID = ? ";

                                using (OleDbConnection ThirdConnection = new OleDbConnection(accessConnectionString))
                                {
                                    ThirdConnection.Open();

                                    using (OleDbCommand ThirdCommand = new OleDbCommand(ThirdQuery, ThirdConnection))
                                    {
                                        ThirdCommand.Parameters.AddWithValue("Account.ListID", OleDbType.VarChar).Value = reader["APAccountRefListID"];

                                        using (OleDbDataReader ThirdReader = ThirdCommand.ExecuteReader())
                                        {
                                            if (ThirdReader.HasRows) // Check if there are rows returned by the third query
                                            {
                                                while (ThirdReader.Read())
                                                {
                                                    newBill.AccountName = ThirdReader["Name"] != DBNull.Value ? ThirdReader["Name"].ToString() : string.Empty;
                                                    newBill.AccountNumber = ThirdReader["AccountNumber"] != DBNull.Value ? ThirdReader["AccountNumber"].ToString() : string.Empty;
                                                }
                                            }
                                        }
                                    }

                                    ThirdConnection.Close();
                                }
                                bills.Add(newBill);
                            }
                        }
                    }

                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return bills;
        }
        public List<BillTable> GetAccountsPayableData_CPI(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                int nextID = GetNextIncrementalID_APV(accessConnectionString);
                using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                {
                    accessConnection.Open();

                    // Retrieve data from Access database
                    string query = @"SELECT TOP 1000 
                       Bill.TxnDate,
                       Bill.RefNumber,
                       Bill.VendorRefFullName, 
                       Bill.TermsRefFullName,
                       Bill.DueDate, 
                       Bill.AmountDue, 
                       Bill.Memo,
                       Bill.IsPaid,
                       Bill.APAccountRefFullName,
                       Bill.APAccountRefListID,
                       Bill.TxnID
                   FROM Bill
                   WHERE Bill.RefNumber = ?";

                    using (OleDbCommand command = new OleDbCommand(query, accessConnection))
                    {
                        command.Parameters.AddWithValue("RefNumber", OleDbType.VarChar).Value = refNumber;
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                BillTable newBill = new BillTable
                                {
                                    // BillPaymentCheckLine table
                                    DateCreated = reader["TxnDate"] != DBNull.Value ? Convert.ToDateTime(reader["TxnDate"]).Date : DateTime.MinValue,
                                    RefNumber = reader["RefNumber"] != DBNull.Value ? reader["RefNumber"].ToString() : string.Empty,
                                    Vendor = reader["VendorRefFullName"] != DBNull.Value ? reader["VendorRefFullName"].ToString() : string.Empty,
                                    APAccountRefFullName = reader["APAccountRefFullName"] != DBNull.Value ? reader["APAccountRefFullName"].ToString() : string.Empty,
                                    TermsRefFullName = reader["TermsRefFullName"] != DBNull.Value ? reader["TermsRefFullName"].ToString() : string.Empty,
                                    DueDate = reader["DueDate"] != DBNull.Value ? Convert.ToDateTime(reader["DueDate"]).Date : DateTime.MinValue,
                                    AmountDue = reader["AmountDue"] != DBNull.Value ? Convert.ToDouble(reader["AmountDue"]) : 0.0,
                                    Memo = reader["Memo"] != DBNull.Value ? reader["Memo"].ToString() : string.Empty,
                                    IsPaid = reader["IsPaid"] != DBNull.Value ? Convert.ToBoolean(reader["IsPaid"]) : false,

                                    // Increment
                                    IncrementalID = nextID.ToString("D6")
                                };

                                string secondQuery = @"SELECT TOP 1000
                                        BillItemLine.ItemLineItemRefFullName AS AccountRefFullName, 
                                        BillItemLine.ItemLineAmount AS Amount,
                                        BillItemLine.ItemLineClassRefFullName AS ClassRefFullName,
                                        BillItemLine.ItemLineitemRefListID AS BillRefListID,
                                        BillItemLine.ItemLineDesc AS ItemExpenseMemo
                                    FROM 
                                        BillItemLine 
                                    WHERE 
                                        BillItemLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillItemLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];
                                        //secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineClassRefFullName = secondReader["ClassRefFullName"] != DBNull.Value ? secondReader["ClassRefFullName"].ToString() : string.Empty;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                Console.WriteLine($"Item: {itemLineItemRefFullName}, Class: {itemLineClassRefFullName}, Amount: {itemLineAmount}, Memo: {itemLineItemMemo}");


                                                //Console.WriteLine(itemLineAmount);
                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ItemLineItemRefFullName = itemLineItemRefFullName,
                                                    ItemLineAmount = itemLineAmount,
                                                    ItemLineClassRefFullName = itemLineClassRefFullName,
                                                    ItemLineMemo = itemLineItemMemo,
                                                });

                                                string fourthQuery = @"SELECT 
                                                        AssetAccountRefFullname
                                                    FROM 
                                                        item
                                                    WHERE 
                                                        item.listID = ?";

                                                string billRefListID = secondReader["BillRefListID"].ToString();

                                                using (OleDbConnection fourthConnection = new OleDbConnection(accessConnectionString))
                                                {
                                                    fourthConnection.Open();

                                                    using (OleDbCommand fourthCommand = new OleDbCommand(fourthQuery, fourthConnection))
                                                    {
                                                        fourthCommand.Parameters.AddWithValue("Account.ListID", OleDbType.VarChar).Value = billRefListID;

                                                        using (OleDbDataReader fourthReader = fourthCommand.ExecuteReader())
                                                        {
                                                            while (fourthReader.Read())
                                                            {
                                                                // Store additional account particulars in lists
                                                                string accountNameParticulars = fourthReader["AssetAccountRefFullname"] != DBNull.Value ? fourthReader["AssetAccountRefFullname"].ToString() : string.Empty;
                                                                //string accountNumberParticulars = fourthReader["AccountNumber"] != DBNull.Value ? fourthReader["AccountNumber"].ToString() : string.Empty;

                                                                newBill.AccountNameParticularsList.Add(accountNameParticulars);
                                                                //newBill.AccountNumberParticularsList.Add(accountNumberParticulars);
                                                            }
                                                        }
                                                    }

                                                    fourthConnection.Close();
                                                }
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }
                                string secondQuery2 = @"SELECT
                                        BillExpenseLine.ExpenseLineAccountRefFullName AS AccountRefFullName, 
                                        BillExpenseLine.ExpenseLineAmount AS Amount,
                                        BillExpenseLine.ExpenseLineClassRefFullName AS ClassRefFullName,
                                        BillExpenseLine.ExpenseLineAccountRefListID AS BillRefListID,
                                        BillExpenseLine.ExpenseLineMemo AS ItemExpenseMemo
                                    FROM 
                                        [BillExpenseLine]
                                    WHERE 
                                        BillExpenseLine.TxnID = ?";

                                using (OleDbConnection secondConnection = new OleDbConnection(accessConnectionString))
                                {
                                    secondConnection.Open();

                                    using (OleDbCommand secondCommand = new OleDbCommand(secondQuery2, secondConnection))
                                    {
                                        secondCommand.Parameters.AddWithValue("BillExpenseLine.TxnID", OleDbType.VarChar).Value = reader["TxnID"];

                                        using (OleDbDataReader secondReader = secondCommand.ExecuteReader())
                                        {
                                            while (secondReader.Read())
                                            {
                                                string itemLineItemRefFullName = secondReader["AccountRefFullName"] != DBNull.Value ? secondReader["AccountRefFullName"].ToString() : string.Empty;
                                                double itemLineAmount = secondReader["Amount"] != DBNull.Value ? Convert.ToDouble(secondReader["Amount"]) : 0.0;
                                                string itemLineClassRefFullName = secondReader["ClassRefFullName"] != DBNull.Value ? secondReader["ClassRefFullName"].ToString() : string.Empty;
                                                string itemLineItemMemo = secondReader["ItemExpenseMemo"] != DBNull.Value ? secondReader["ItemExpenseMemo"].ToString() : string.Empty;

                                                Console.WriteLine($"Item: {itemLineItemRefFullName}, Amount: {itemLineAmount}, Memo: {itemLineItemMemo}");


                                                newBill.ItemDetails.Add(new ItemDetail
                                                {
                                                    ExpenseLineItemRefFullName = itemLineItemRefFullName,
                                                    ExpenseLineAmount = itemLineAmount,
                                                    ExpenseLineClassRefFullName = itemLineClassRefFullName,
                                                    ExpenseLineMemo = itemLineItemMemo,
                                                });
                                            }
                                        }
                                    }

                                    secondConnection.Close();
                                }

                                string ThirdQuery = @"SELECT 
                                               Name, AccountNumber
                                            FROM 
                                                Account
                                            WHERE 
                                                Account.listID = ? ";

                                using (OleDbConnection ThirdConnection = new OleDbConnection(accessConnectionString))
                                {
                                    ThirdConnection.Open();

                                    using (OleDbCommand ThirdCommand = new OleDbCommand(ThirdQuery, ThirdConnection))
                                    {
                                        ThirdCommand.Parameters.AddWithValue("Account.ListID", OleDbType.VarChar).Value = reader["APAccountRefListID"];

                                        using (OleDbDataReader ThirdReader = ThirdCommand.ExecuteReader())
                                        {
                                            if (ThirdReader.HasRows) // Check if there are rows returned by the third query
                                            {
                                                while (ThirdReader.Read())
                                                {
                                                    newBill.AccountName = ThirdReader["Name"] != DBNull.Value ? ThirdReader["Name"].ToString() : string.Empty;
                                                    newBill.AccountNumber = ThirdReader["AccountNumber"] != DBNull.Value ? ThirdReader["AccountNumber"].ToString() : string.Empty;
                                                }
                                            }
                                        }
                                    }

                                    ThirdConnection.Close();
                                }
                                bills.Add(newBill);
                            }
                        }
                    }

                    accessConnection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from Access database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return bills;
        }


        public int GetNextIncrementalID_CV(string accessConnectionString)
        {
            int incrementalID = 0;

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                string query = "SELECT FormattedID FROM CVIncrement";
                OleDbCommand command = new OleDbCommand(query, connection);

                try
                {
                    connection.Open();
                    object result = command.ExecuteScalar();

                    if (result != null)
                    {
                        int currentID = Convert.ToInt32(result);
                        // Increment the ID
                        //incrementalID = "CV" + currentID.ToString("D6"); // Format to CV000001
                        incrementalID = currentID; // Format to CV000001
                    }
                    else
                    {
                        // If no record exists, create one with FormattedID set to 0
                        query = "INSERT INTO CVIncrement (FormattedID) VALUES (0)";
                        command = new OleDbCommand(query, connection);
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            incrementalID = 0;
                        }
                        else
                        {
                            Console.WriteLine("Error creating a new record.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }

            return incrementalID;
        }

        public int GetNextIncrementalID_APV(string accessConnectionString)
        {
            int incrementalID = 0;

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                string query = "SELECT FormattedID FROM APVIncrement";
                OleDbCommand command = new OleDbCommand(query, connection);

                try
                {
                    connection.Open();
                    object result = command.ExecuteScalar();

                    if (result != null)
                    {
                        int currentID = Convert.ToInt32(result);
                        // Increment the ID
                        //incrementalID = "CV" + currentID.ToString("D6"); // Format to CV000001
                        incrementalID = currentID; // Format to CV000001
                    }
                    else
                    {
                        // If no record exists, create one with FormattedID set to 0
                        query = "INSERT INTO APVIncrement (FormattedID) VALUES (0)";
                        command = new OleDbCommand(query, connection);
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            incrementalID = 0;
                        }
                        else
                        {
                            Console.WriteLine("Error creating a new record.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }

            return incrementalID;
        }

        
    }
}
