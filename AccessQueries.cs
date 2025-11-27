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

        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_KAYAK(string refNumber)
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
                                                string assetAccountRefListID = secondReader["AssetAccountRefListID"] != DBNull.Value? secondReader["AssetAccountRefListID"].ToString() : string.Empty;

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
                                                                newCheckItem.AssetAccountNumber = accReader["AccountNumber"] != DBNull.Value? accReader["AccountNumber"].ToString() : string.Empty;
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
