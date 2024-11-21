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

namespace VoucherPro
{
    internal class AccessQueries
    {
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
    }
}
