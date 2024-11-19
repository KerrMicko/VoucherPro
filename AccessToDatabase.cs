using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VoucherPro
{
    internal class AccessToDatabase
    {
        public static string GetAccessConnectionString()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = "CheckDatabase.accdb";
            string resourcePath = Path.Combine(baseDirectory, fileName);
            string accessConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={resourcePath};Persist Security Info=False;";
            return accessConnectionString;
        }

        public static string GetQBConnectionString()
        {
            string qbConnectionString = "DSN=QuickBooks Data;";
            return qbConnectionString;
        }

        //List<string> tableNamesWithItemReceipt = new List<string> { "Account", "Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine", "Check", "CheckExpenseLine", "CheckItemLine", "Company", "Item", "ItemReceipt", "ItemReceiptExpenseLine", "ItemReceiptItemLine",  "Transaction", "Vendor" };
        List<string> tableNamesWithItemReceipt = new List<string> { "Account", 
            "Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine", 
            "Check", "CheckExpenseLine", "CheckItemLine", "Company", 
            "Item", "ItemReceipt", "ItemReceiptExpenseLine", "ItemReceiptItemLine", 
            "Transaction", "Vendor" };

        List<string> tableNamesWithoutItemReceipt = new List<string> { "Account", 
            "Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine", 
            "Check", "CheckExpenseLine", "CheckItemLine", "Company", "Item", "Vendor" };

        public async Task FetchAndSaveData()
        {
            string odbcConnectionString = GetQBConnectionString();
            string accessConnectionString = GetAccessConnectionString();

            List<string> tableNames;

            if (GlobalVariables.includeItemReceipt)
            {
                tableNames = tableNamesWithItemReceipt;
            }
            else
            {
                tableNames = tableNamesWithoutItemReceipt;
            }

            using (OdbcConnection odbcConnection = new OdbcConnection(odbcConnectionString))
            {
                try
                {
                    await odbcConnection.OpenAsync();

                    using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                    {
                        await accessConnection.OpenAsync();

                        foreach (var tableName in tableNames)
                        {
                            // Ensure table name is enclosed in square brackets
                            string formattedTableName = $"[{tableName}]";

                            // Create ODBC command to select data from the QuickBooks Data DSN
                            OdbcCommand odbcCommand = odbcConnection.CreateCommand();
                            odbcCommand.CommandText = $"SELECT * FROM {tableName}";

                            // Execute the ODBC command to fetch the data
                            using (OdbcDataReader reader = (OdbcDataReader)await odbcCommand.ExecuteReaderAsync())
                            {
                                // Create the destination table name with "QB_" prefix
                                //string destinationTableName = $"QB_{tableName}";
                                //string destinationTableName = $"{tableName}";
                                string destinationTableName = formattedTableName;

                                // Construct the Access SQL command for inserting data
                                OleDbCommand accessCommand = accessConnection.CreateCommand();
                                string columnNames = GetColumnNames(reader);
                                string parameterNames = GetParameterNames(reader);
                                accessCommand.CommandText = $"INSERT INTO {destinationTableName} ({columnNames}) VALUES ({parameterNames})";

                                // Add parameters to the Access command
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    accessCommand.Parameters.Add(new OleDbParameter($"@param{i}", reader.GetFieldType(i)));
                                }

                                // Transfer data row by row
                                while (await reader.ReadAsync())
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        accessCommand.Parameters[$"@param{i}"].Value = reader.GetValue(i);
                                    }

                                    try
                                    {
                                        await accessCommand.ExecuteNonQueryAsync();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"An error occurred while inserting data into {destinationTableName}: {ex.Message}\nSQL: {accessCommand.CommandText}");
                                        return;
                                    }
                                }
                            }
                        }
                        accessConnection.Close();
                    }
                    odbcConnection.Close();

                    MessageBox.Show("Data transfer completed successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching data: " + ex.Message);
                }
            }
        }

        // Helper method to get column names from the OdbcDataReader
        private string GetColumnNames(OdbcDataReader reader)
        {
            List<string> columnNames = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                columnNames.Add($"[{reader.GetName(i)}]");
            }
            return string.Join(", ", columnNames);
        }

        // Helper method to get parameter names for the SQL command
        private string GetParameterNames(OdbcDataReader reader)
        {
            List<string> parameterNames = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                parameterNames.Add($"@param{i}");
            }
            return string.Join(", ", parameterNames);
        }

        public void DeleteSpecifiedTablesData()
        {
            string accessConnectionString = GetAccessConnectionString();

            List<string> tableNames;

            if (GlobalVariables.includeItemReceipt)
            {
                tableNames = tableNamesWithItemReceipt;
            }
            else
            {
                tableNames = tableNamesWithoutItemReceipt;
            }

            using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
            {
                accessConnection.Open();

                foreach (var tableName in tableNames)
                {
                    // Construct the SQL command to delete all data from the table
                    OleDbCommand accessCommand = accessConnection.CreateCommand();
                    accessCommand.CommandText = $"DELETE FROM [{tableName}]";

                    try
                    {
                        accessCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred while deleting data from {tableName}: {ex.Message}");
                        return;
                    }
                }
                accessConnection.Close();
            }

            Console.WriteLine("All data from specified tables has been deleted.");
        }

        public class AmountToWordsConverter
        {
            private static string[] units = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };
            private static string[] teens = { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
            private static string[] tens = { "", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };
            private static string[] thousandsGroups = { "", " Thousand", " Million", " Billion" };

            public static string Convert(double amount)
            {
                if (amount == 0)
                    return "Zero Pesos Only";

                if (amount < 0)
                    return "Negative amount, cannot convert to words";

                int pesos = (int)Math.Floor(amount);
                int centavos = (int)Math.Round((amount - pesos) * 100);

                string pesoWords = ConvertToWords(pesos);
                string centavoWords = ConvertToWords(centavos);

                string result = "";
                if (centavos > 0)
                {
                    result = pesoWords + " Pesos";
                    //result = pesoWords + " and " + centavos + "/100 Pesos Only"; kanan terrys
                    result += " and " + centavoWords + " Centavos Only";
                }
                else
                {
                    result = pesoWords + " Pesos Only";
                }

                return result;
            }

            private static string ConvertToWords(int number)
            {
                if (number == 0)
                    return "Zero";

                if (number < 0)
                    return "Negative " + ConvertToWords(Math.Abs(number));

                string words = "";

                for (int i = 0; number > 0; i++)
                {
                    if (number % 1000 != 0)
                    {
                        words = ConvertHundreds(number % 1000) + thousandsGroups[i] + " " + words;
                    }
                    number /= 1000;
                }

                return words.Trim();
            }

            private static string ConvertHundreds(int number)
            {
                string words = "";

                if (number >= 100)
                {
                    words += units[number / 100] + " Hundred ";
                    number %= 100;
                }

                if (number >= 10 && number <= 19)
                {
                    words += teens[number - 10] + " ";
                    number = 0;
                }

                if (number >= 20)
                {
                    words += tens[number / 10] + " ";
                    number %= 10;
                }

                if (number >= 1 && number <= 9)
                {
                    words += units[number] + " ";
                }

                return words;
            }
        }
    }
}
