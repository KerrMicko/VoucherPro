using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VoucherPro.AccessToDatabase;
using static VoucherPro.DataClass;
using System.Windows.Forms;

namespace VoucherPro.Clients
{
    public class Layouts_CPI
    {
        Font font_Six = new Font("Microsoft Sans Serif", 6, FontStyle.Regular);
        Font font_Seven = new Font("Microsoft Sans Serif", 7, FontStyle.Regular);
        Font font_SevenBold = new Font("Microsoft Sans Serif", 7, FontStyle.Bold);
        Font font_Eight = new Font("Microsoft Sans Serif", 8, FontStyle.Regular);
        Font font_EightBold = new Font("Microsoft Sans Serif", 8, FontStyle.Bold);
        Font font_Nine = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
        Font font_NineBold = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
        Font font_Ten = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
        Font font_TenBold = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
        Font font_Eleven = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
        Font font_ElevenBold = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);
        Font font_Twelve = new Font("Microsoft Sans Serif", 12, FontStyle.Regular);
        Font font_TwelveBold = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);

        public void PrintPage_CPI(object sender, PrintPageEventArgs e, int layoutIndex, string seriesNumber, object data)
        {
            StringFormat sfAlignRight = new StringFormat { Alignment = StringAlignment.Far | StringAlignment.Far };
            StringFormat sfAlignCenterRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignLeftCenter = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

            switch (layoutIndex)
            {
                case 1:
                    if (data is List<CheckTableExpensesAndItems>)
                    {
                        Layout_CheckVoucher_Check(e, data as List<CheckTableExpensesAndItems>, seriesNumber, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    }
                    else if (data is List<BillTable>)
                    {
                        Layout_CheckVoucher_Bill(e, data as List<BillTable>, seriesNumber, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    }
                    break;
                default:
                    throw new ArgumentException("Invalid layout index");
            }
        }

        private void Layout_CheckVoucher_Check(PrintPageEventArgs e, List<CheckTableExpensesAndItems> checkData, string seriesNumber, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, StringFormat sfAlignRight)
        {
            Font font_Header = font_EightBold;
            Font font_Data = font_Eight;
            Font font_Data2 = font_Seven;


            Image image = Properties.Resources.CPI_LOGO;
            Bitmap resizedBitmap = null;

            if (image != null)
            {
                int imageWidth = 140;
                int imageHeight = (int)((double)image.Height / image.Width * imageWidth);
                resizedBitmap = new Bitmap(image, new Size(imageWidth, imageHeight));
            }

            if (resizedBitmap != null)
            {
                e.Graphics.DrawImage(resizedBitmap, new PointF(50, 40)); // an logo
            }


            //string companyName = "Kayak Constuction Corp.";
            //string companyTIN = "VAT Reg. TIN: 243-354-422-00000";
            string companyAddress = "3F Renaissance Tower 1000 Meralco Avenue, Brgy Ugong, 1604 \n Pasig City, Philippines";
            string companyTelNo = "Tel. No.: 63 2 631 8401 - 08";
            string cvText = "CHECK VOUCHER";

            //e.Graphics.DrawString(companyName, font_TwelveBold, Brushes.Black, new PointF(48, 40));
            //e.Graphics.DrawString(companyTIN, font_Eight, Brushes.Black, new PointF(200, 65));
            e.Graphics.DrawString(companyAddress, font_Seven, Brushes.Black, new PointF(50, 60));
            e.Graphics.DrawString(companyTelNo, font_Seven, Brushes.Black, new PointF(50, 70));

            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));

            // 1st Table - Details
            int tableWidth = 750;
            int tableHeight = 40;
            int firstTableYPos = 180 + tableHeight + 7;
            //e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 50, 150, tableHeight + 10); // CV Ref. No.
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 100, 150, tableHeight); // Print Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 140, tableWidth - 450, tableHeight); // Payee
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 450, 140, 150, tableHeight); // Bank
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 300, 140, 150, tableHeight); // Check Number
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 140, 150, tableHeight); // Check Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 180, tableWidth - 150, tableHeight); // Amount in Words
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 180, 150, tableHeight); // Amount

            // 1st Table Header
            //e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 50 + 2, 150, tableHeight + 10));
            e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + tableWidth - 190, 100 - 8 - 1, 150, tableHeight), sfAlignCenter);

            e.Graphics.DrawString("Payee", font_Header, Brushes.Black, new RectangleF(50 + 3, 140 + 2, tableWidth - 450, tableHeight));
            e.Graphics.DrawString("Bank", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 450, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Number", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 300, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Date", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 140 + 2, 150, tableHeight));

            e.Graphics.DrawString("Amount in Words", font_Header, Brushes.Black, new RectangleF(50 + 3, 180 + 2, tableWidth - 150, tableHeight));
            e.Graphics.DrawString("Amount", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 180 + 2, 150, tableHeight));

            // 1st Table Data
            string payee = checkData[0].PayeeFullName;
            string bankAccount = checkData[0].BankAccount;
            string checkNumber = checkData[0].RefNumber;
            double amount = checkData[0].TotalAmount;
            string amountInWords = AmountToWordsConverter.Convert(amount);
            DateTime chequeDate = checkData[0].DateCreated;
            DateTime dateTime = DateTime.Now; //PRINT DATE


            //e.Graphics.DrawString(seriesNumber, font_TenBold, Brushes.Black, new RectangleF(50 + tableWidth - 150, 50 + 6, 150, tableHeight + 10), sfAlignCenter); // CV Ref. No.
            e.Graphics.DrawString(seriesNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 + 8, 150, tableHeight), sfAlignCenter); // Print Date

            e.Graphics.DrawString(payee, font_Data, Brushes.Black, new RectangleF(50 + 15, 140 + 6, tableWidth - 450, tableHeight), sfAlignLeftCenter); // Payee
            e.Graphics.DrawString(bankAccount, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 450, 140 + 6, 150, tableHeight), sfAlignCenter); // Bank
            e.Graphics.DrawString(checkNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 300, 140 + 6, 150, tableHeight), sfAlignCenter); // Check Number
            e.Graphics.DrawString(chequeDate.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 140 + 8, 150, tableHeight), sfAlignCenter); // Check Date

            e.Graphics.DrawString(amountInWords, font_Data, Brushes.Black, new RectangleF(50 + 15, 180 + 6, tableWidth - 150, tableHeight), sfAlignLeftCenter); // Amount in words
            e.Graphics.DrawString("₱", font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 + 10, 180 + 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 - 10, 180 + 6, 150, tableHeight), sfAlignCenterRight); // Amount

            // 2nd Table - Particulars
            int perItemHeight = 25;

            int secondTableHeight = 130 - 90; // 75

            int itemsLimit = 8;

            Dictionary<string, Tuple<double, string>> groupedItemData = new Dictionary<string, Tuple<double, string>>();

            try
            {
                for (int i = 0; i < checkData.Count; i++)
                {
                    string itemName = checkData[i].AccountName;
                    string itemAccountNumber = checkData[i].AccountNumber;
                    double itemAmount = checkData[i].ItemAmount;
                    string itemClass = checkData[i].ItemClass;
                    string itemAssetName = checkData[i].AssetAccountName;
                    string itemAssetNumber = checkData[i].AssetAccountNumber;

                    if (itemAssetName != "" && itemAmount != 0)
                    {
                        if (groupedItemData.ContainsKey(itemAssetName))
                        {
                            // Update the existing value by adding the item amount and keeping the itemAssetNumber
                            var existingData = groupedItemData[itemAssetName];
                            groupedItemData[itemAssetName] = new Tuple<double, string>(existingData.Item1 + itemAmount, existingData.Item2);
                        }
                        else
                        {
                            // Add new entry with both amount and asset account number
                            groupedItemData[itemAssetName] = new Tuple<double, string>(itemAmount, itemAssetNumber);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while grouping entries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                foreach (var entry in groupedItemData)
                {
                    secondTableHeight += 25;
                }
                foreach (var account in checkData)
                {
                    if (!string.IsNullOrEmpty(account.Account))
                    {

                        secondTableHeight += perItemHeight;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while adding height: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            int secondTableYPos = firstTableYPos + 40 + secondTableHeight;

            // 2nd Table Headers
            e.Graphics.DrawRectangle(Pens.Black, 50, firstTableYPos, tableWidth - (300 + 180), 20); // Particular header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 270, firstTableYPos, 185, 20); // Class header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 455, firstTableYPos, 145, 20); // Debit header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 600, firstTableYPos, 150, 20); // Credit header

            e.Graphics.DrawString("Particular", font_Header, Brushes.Black, new RectangleF(25, firstTableYPos, tableWidth - (300 + 100), 20), sfAlignCenter);
            e.Graphics.DrawString("Class", font_Header, Brushes.Black, new RectangleF(50 + 300 + 10, firstTableYPos, 100, 20), sfAlignCenter);
            e.Graphics.DrawString("Debit", font_Header, Brushes.Black, new RectangleF(50 + 450, firstTableYPos, 150, 20), sfAlignCenter);
            e.Graphics.DrawString("Credit", font_Header, Brushes.Black, new RectangleF(50 + 600, firstTableYPos, 150, 20), sfAlignCenter);

            e.Graphics.DrawLine(Pens.Black, 50 + 270, firstTableYPos + 20, 50 + 270, secondTableYPos); // Line ha class
            e.Graphics.DrawLine(Pens.Black, 50 + 455, firstTableYPos + 20, 50 + 455, secondTableYPos); // Line ha debit
            e.Graphics.DrawLine(Pens.Black, 50 + 600, firstTableYPos + 20, 50 + 600, secondTableYPos); // Line ha credit
            e.Graphics.DrawLine(Pens.Black, 50, secondTableYPos, tableWidth + 50, secondTableYPos); // Line ha ubos

            // 2nd Table Data

            //string particularAccount = checkData[0].AccountNumber + " - " + checkData[0].AccountName;
            string particularBank = checkData[0].BankAccount;
            string particularMemo = checkData[0].Memo; // Remark or Memo

            int itemAccountHeight = 0;

            int pos = 0;
            int amountPos = 0;

            double debitTotalAmount = 0;
            double creditTotalAmount = 0;
            /*for (int i = 0; i < checkData.Count; i++)
            {
                e.Graphics.DrawString(checkData[i].Item + checkData[i].Account, font_eight, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + itemAccountHeight, tableWidth - (300 + 100), perItemHeight)); // Item
                //e.Graphics.DrawString(i.ToString(), font_eight, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + itemAccountHeight, tableWidth - (300 + 100), perItemHeight)); // Item

                double itemAmount = checkData[i].ItemAmount;
                double expensesAmount = checkData[i].ExpensesAmount;

                if (itemAmount != 0)
                {
                    if (itemAmount > 0)
                    {
                        e.Graphics.DrawString(itemAmount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + itemAccountHeight, 150, perItemHeight), sfAlignRight); // Debit
                        //debitTotalAmount += itemAmount;
                    }
                    else
                    {
                        double absoluteAmount = Math.Abs(itemAmount);
                        e.Graphics.DrawString(absoluteAmount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 20 + 4 + itemAccountHeight, 150, perItemHeight), sfAlignRight); // Credit
                        //creditTotalAmount += absoluteAmount;
                    }
                }
                else if (expensesAmount != 0)
                {
                    if (expensesAmount > 0)
                    {
                        e.Graphics.DrawString(expensesAmount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + itemAccountHeight, 150, perItemHeight), sfAlignRight); // Debit
                        //debitTotalAmount += expensesAmount;
                    }
                    else
                    {
                        double absoluteAmount = Math.Abs(expensesAmount);
                        e.Graphics.DrawString(absoluteAmount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 20 + 4 + itemAccountHeight, 150, perItemHeight), sfAlignRight); // Credit
                        //creditTotalAmount += absoluteAmount;
                    }
                }
                //e.Graphics.DrawRectangle(Pens.Red, 50, firstTableYPos + 20 + itemAccountHeight, tableWidth - (300 + 100), perItemHeight); // Particular 
                itemAccountHeight += 25;
            }*/

            foreach (var item in groupedItemData)
            {
                double totalAmount = item.Value.Item1;  // Total amount (Item1 in Tuple)
                string assetAccountNumber = item.Value.Item2;  // Asset Account Number (Item2 in Tuple)

                // Draw the account name (item.Key)
                e.Graphics.DrawString($"{assetAccountNumber}" + " - " + $"{item.Key}", font_Data, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 150), 25));

                // Draw the total amount (item.Value.Item1)
                e.Graphics.DrawString($"{totalAmount:N2}", font_Data, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + pos, 150, perItemHeight), sfAlignRight); // Credit


                // Accumulate totals if needed
                if (totalAmount > 0)
                {
                    debitTotalAmount += totalAmount;
                }

                pos += 25;
            }



            amountPos += pos;

            foreach (var check in checkData)
            {
                string accNum = check.AccountNumber?.Trim().TrimEnd('-') ?? "";
                string accName = check.AccountNameCheck?.Trim().TrimStart('-') ?? "";
                string accountText = string.Join(" - ", new[] { accNum, accName }.Where(s => !string.IsNullOrEmpty(s)));

                e.Graphics.DrawString(accountText, font_Data, Brushes.Black, new RectangleF(55, firstTableYPos + 24 + pos, tableWidth - 485, perItemHeight));
                //e.Graphics.DrawRectangle(Pens.Blue, 55, firstTableYPos + 24 + pos, tableWidth - 485, perItemHeight);

                e.Graphics.DrawString(check.ItemClass, font_Data2, Brushes.Black, new RectangleF(50 + 270, firstTableYPos + 20 + 4 + pos, tableWidth - (500 + 65), perItemHeight), sfAlignCenter); // Itemclass
                //e.Graphics.DrawRectangle(Pens.Red, 50 + 270, firstTableYPos + 20 + 4 + pos, tableWidth - (500 + 65), perItemHeight);


                double expensesAmount = check.ExpensesAmount;
                if (expensesAmount != 0)
                {
                    if (expensesAmount > 0)
                    {
                        e.Graphics.DrawString(expensesAmount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + amountPos, 150, perItemHeight), sfAlignRight); // Debit
                        debitTotalAmount += expensesAmount;
                    }
                    else
                    {
                        double absoluteAmount = Math.Abs(expensesAmount);
                        e.Graphics.DrawString(absoluteAmount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 20 + 4 + amountPos, 150, perItemHeight), sfAlignRight); // Credit
                        //creditTotalAmount += absoluteAmount;
                    }
                    pos += 25;
                    amountPos += 25;
                }
            }


            // Total Amount
            creditTotalAmount += amount;

            //e.Graphics.DrawString(particularAccount, font_eight, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4, tableWidth - (300 + 100), perItemHeight)); // Item1
            e.Graphics.DrawString(particularBank, font_Data, Brushes.Black, new RectangleF(50 + 10 - 2, firstTableYPos + 30 + pos, tableWidth - (300 + 110), perItemHeight)); // Item1 Bank
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 30 + pos, 150, perItemHeight), sfAlignRight); // Credit - bank

            e.Graphics.DrawString("*Remarks: ", font_Data, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 50 - 5 + pos, tableWidth - (300 + 50), 60)); // Item1 Remark / Memo
            e.Graphics.DrawString(particularMemo, font_Seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 50 - 3 + pos, tableWidth - (300 + 220), 60)); // Item1 Remark / Memo

            // Debit Credit Total
            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 450 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(debitTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 450 - 5, secondTableYPos + 7, 150, tableHeight), sfAlignRight);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 25, 50 + 450 - 5 + 150, secondTableYPos + 25);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 28, 50 + 450 - 5 + 150, secondTableYPos + 28);

            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 600 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(creditTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 600 - 5, secondTableYPos + 7, 150, tableHeight), sfAlignRight);
            e.Graphics.DrawLine(Pens.Black, 50 + 600 + 5, secondTableYPos + 25, 50 + 600 - 5 + 150, secondTableYPos + 25);
            e.Graphics.DrawLine(Pens.Black, 50 + 600 + 5, secondTableYPos + 28, 50 + 600 - 5 + 150, secondTableYPos + 28);

            // Others Header
            int signWidth = 187; // 150 if 5 columns
            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 2, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 3, secondTableYPos + 45, signWidth, tableHeight - 20);

            e.Graphics.DrawString("Prepared By:", font_Header, Brushes.Black, new RectangleF(50, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Checked By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Approved By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Received By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);

            // Others Data
            int othersYPos = secondTableYPos + 45 + 20 + 75;
            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 2, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 3, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);

            AccessToDatabase accessToDatabase = new AccessToDatabase();

            var (PreparedByName, PreparedByPosition,
                ReviewedByName, ReviewedByPosition,
                RecommendingApprovalName, RecommendingApprovalPosition,
                ApprovedByName, ApprovedByPosition,
                ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

            e.Graphics.DrawString(PreparedByName, font_SevenBold, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(PreparedByPosition, font_Seven, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            e.Graphics.DrawString(ReviewedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ReviewedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            e.Graphics.DrawString(ApprovedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ApprovedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            e.Graphics.DrawString(ReceivedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ReceivedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

        }

        private void Layout_CheckVoucher_Bill(PrintPageEventArgs e, List<BillTable> billData, string seriesNumber, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, StringFormat sfAlignRight)
        {
            Font font_Header = font_EightBold;
            Font font_Data = font_Eight;

            Image image = Properties.Resources.CPI_LOGO;
            Bitmap resizedBitmap = null;

            if (image != null)
            {
                int imageWidth = 140;
                int imageHeight = (int)((double)image.Height / image.Width * imageWidth);
                resizedBitmap = new Bitmap(image, new Size(imageWidth, imageHeight));
            }

            if (resizedBitmap != null)
            {
                e.Graphics.DrawImage(resizedBitmap, new PointF(50, 40)); // an logo
            }


            //string companyName = "Kayak Constuction Corp.";
            //string companyTIN = "VAT Reg. TIN: 243-354-422-00000";
            string companyAddress = "3F Renaissance Tower 1000 Meralco Avenue, Brgy Ugong, 1604 Pasig City, Philippines";
            string companyTelNo = "Tel. No.: 63 2 631 8401 - 08";
            string cvText = "CHECK VOUCHER";


            //e.Graphics.DrawString(companyName, font_TwelveBold, Brushes.Black, new PointF(48, 40));
            //e.Graphics.DrawString(companyTIN, font_Eight, Brushes.Black, new PointF(200, 65));
            e.Graphics.DrawString(companyAddress, font_Seven, Brushes.Black, new PointF(60, 80));
            e.Graphics.DrawString(companyTelNo, font_Seven, Brushes.Black, new PointF(60, 95));

            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));

            // Particulars memo
            int tableWidth = 750;
            int memoTableHeight = 65;

            string particularMemo = billData[0].Memo;

            Rectangle rectMemoHeader = new Rectangle(50, 230, tableWidth, 20);
            Rectangle rectMemoBody = new Rectangle(50, 230 + 20, tableWidth, 32);

            e.Graphics.DrawRectangle(Pens.Black, rectMemoHeader);
            e.Graphics.DrawRectangle(Pens.Black, rectMemoBody);

            e.Graphics.DrawString("PARTICULARS", font_Header, Brushes.Black, rectMemoHeader, sfAlignCenter);
            e.Graphics.DrawString(particularMemo, font_Data, Brushes.Black, rectMemoBody, sfAlignCenter);

            // 1st Table - Details
            int tableHeight = 40;
            int firstTableYPos = 180 + tableHeight + 7 + memoTableHeight;

            int payeeWidth = tableWidth - 475; // 450
            //e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 50, 150, tableHeight + 10); // CV Ref. No.
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 100, 150, tableHeight); // Print Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 140, payeeWidth, tableHeight); // Payee
            e.Graphics.DrawRectangle(Pens.Black, 50 + payeeWidth, 140, 150 + 25, tableHeight); // Bank
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 300, 140, 150, tableHeight); // Check Number
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 140, 150, tableHeight); // Check Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 180, tableWidth - 150, tableHeight); // Amount in Words
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 180, 150, tableHeight); // Amount

            // 1st Table Header
            //e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 50 + 2, 150, tableHeight + 10));
            e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + tableWidth - 190, 100 - 8 - 1, 150, tableHeight), sfAlignCenter);

            e.Graphics.DrawString("Payee", font_Header, Brushes.Black, new RectangleF(50 + 3, 140 + 2, payeeWidth, tableHeight));
            e.Graphics.DrawString("Bank", font_Header, Brushes.Black, new RectangleF(50 + 3 + payeeWidth, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Number", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 300, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Date", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 140 + 2, 150, tableHeight));

            e.Graphics.DrawString("Amount in Words", font_Header, Brushes.Black, new RectangleF(50 + 3, 180 + 2, tableWidth - 150, tableHeight));
            e.Graphics.DrawString("Amount", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 180 + 2, 150, tableHeight));

            // 1st Table Data
            string payee = billData[0].PayeeFullName;
            string bankAccount = billData[0].BankAccount;
            string checkNumber = billData[0].RefNumber;
            double amount = billData[0].Amount;
            string amountInWords = AmountToWordsConverter.Convert(amount);
            DateTime chequeDate = billData[0].DateCreated;
            DateTime dateTime = DateTime.Now; //PRINT DATE


            //e.Graphics.DrawString(seriesNumber, font_TenBold, Brushes.Black, new RectangleF(50 + tableWidth - 150, 50 + 6, 150, tableHeight + 10), sfAlignCenter); // CV Ref. No.
            e.Graphics.DrawString(seriesNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 + 8, 150, tableHeight), sfAlignCenter); // Print Date

            e.Graphics.DrawString(payee, font_Data, Brushes.Black, new RectangleF(50 + 15, 140 + 6, payeeWidth, tableHeight), sfAlignLeftCenter); // Payee
            e.Graphics.DrawString(bankAccount, font_Data, Brushes.Black, new RectangleF(50 + payeeWidth, 140 + 6, 150 + 25, tableHeight), sfAlignCenter); // Bank
            e.Graphics.DrawString(checkNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 300, 140 + 6, 150, tableHeight), sfAlignCenter); // Check Number
            e.Graphics.DrawString(chequeDate.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 140 + 8, 150, tableHeight), sfAlignCenter); // Check Date

            e.Graphics.DrawString(amountInWords, font_Data, Brushes.Black, new RectangleF(50 + 15, 180 + 6, tableWidth - 150, tableHeight), sfAlignLeftCenter); // Amount in words
            e.Graphics.DrawString("₱", font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 + 10, 180 + 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 - 10, 180 + 6, 150, tableHeight), sfAlignCenterRight); // Amount

            // 2nd Table - Particulars
            int secondTableHeight = 60; // 75
            int secondTableYPos = firstTableYPos + 40 + secondTableHeight; // 40

            double debitTotalAmount = billData[0].Amount;
            double creditTotalAmount = billData[0].Amount;

            // 2nd Table Headers
            e.Graphics.DrawRectangle(Pens.Black, 50, firstTableYPos, tableWidth - (300 + 100), 20); // Particular header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 300 + 50, firstTableYPos, 100, 20); // Class header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 450, firstTableYPos, 150, 20); // Debit header
            e.Graphics.DrawRectangle(Pens.Black, 50 + 600, firstTableYPos, 150, 20); // Credit header

            e.Graphics.DrawString("Particular", font_Header, Brushes.Black, new RectangleF(50, firstTableYPos, tableWidth - (300 + 100), 20), sfAlignCenter);
            e.Graphics.DrawString("Class", font_Header, Brushes.Black, new RectangleF(50 + 300 + 50, firstTableYPos, 100, 20), sfAlignCenter);
            e.Graphics.DrawString("Debit", font_Header, Brushes.Black, new RectangleF(50 + 450, firstTableYPos, 150, 20), sfAlignCenter);
            e.Graphics.DrawString("Credit", font_Header, Brushes.Black, new RectangleF(50 + 600, firstTableYPos, 150, 20), sfAlignCenter);

            e.Graphics.DrawLine(Pens.Black, 50 + 300 + 50, firstTableYPos + 20, 50 + 300 + 50, secondTableYPos); // Line ha class
            e.Graphics.DrawLine(Pens.Black, 50 + 450, firstTableYPos + 20, 50 + 450, secondTableYPos); // Line ha debit
            e.Graphics.DrawLine(Pens.Black, 50 + 600, firstTableYPos + 20, 50 + 600, secondTableYPos); // Line ha credit
            e.Graphics.DrawLine(Pens.Black, 50, secondTableYPos, tableWidth + 50, secondTableYPos); // Line ha ubos

            // 2nd Table Data
            int perItemHeight = 90;

            string particularAccount = string.Empty;

            /*if (apvFieldText == "")
            {
                //particularAccount = billData[0].AccountNumber + " - " + billData[0].AccountName + "(Bill#: " + billData[0].AppliedRefNumber + ")";
                particularAccount = "Accounts Payable" + "(Bill#: " + billData[0].AppliedRefNumber + ")";
            }
            else
            {
                //particularAccount = billData[0].AccountNumber + " - " + billData[0].AccountName + "(Bill#: " + billData[0].AppliedRefNumber + " / " + "APV#: " + apvFieldText + ")";
                particularAccount = "Accounts Payable" + "(Bill#: " + billData[0].AppliedRefNumber + " / " + "APV#: " + apvFieldText + ")";
            }*/
            string apvFieldText = string.Empty;
            if (apvFieldText == "")
            {
                particularAccount = billData[0].AccountNumber + " - " + billData[0].APAccountRefFullName + " (Bill#: " + billData[0].AppliedRefNumber + ")";
            }
            else
            {
                particularAccount = billData[0].AccountNumber + " - " + billData[0].AccountNumber + billData[0].APAccountRefFullName + " (Bill#: " + billData[0].AppliedRefNumber + " / " + "APV#: " + apvFieldText + ")";
            }


            string particularBank = billData[0].BankAccount;
            //string particularMemo = billData[0].Memo; // Remark or Memo

            //e.Graphics.DrawRectangle(Pens.Red, 50, firstTableYPos + 20, tableWidth - (300 + 100), perItemHeight); // Particular 
            //e.Graphics.DrawRectangle(Pens.Blue, 50 + 70, firstTableYPos + 50, tableWidth - (300 + 170), perItemHeight - 30); // Remark 
            //e.Graphics.DrawRectangle(Pens.Orange, 50 + 300 + 50, firstTableYPos + 20, 100, perItemHeight); // Class 
            //e.Graphics.DrawRectangle(Pens.Yellow, 50 + 450, firstTableYPos + 20, 150, perItemHeight); // Debit 
            //e.Graphics.DrawRectangle(Pens.Green, 50 + 600, firstTableYPos + 20, 150, perItemHeight); // Credit 

            e.Graphics.DrawString(particularAccount, font_Data, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4, tableWidth - (300 + 100), perItemHeight)); // Item1
            e.Graphics.DrawString(particularBank, font_Data, Brushes.Black, new RectangleF(50 + 15, firstTableYPos + 20 + 12 + 4, tableWidth - (300 + 110), perItemHeight)); // Item1 Bank
            //e.Graphics.DrawString("*Remarks: ", font_Data, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 20 + 30 + 4, tableWidth - (300 + 100), perItemHeight - 30)); // Item1 Remark / Memo
            //e.Graphics.DrawString(particularMemo, font_Seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 50 + 4, tableWidth - (300 + 170), perItemHeight)); // Item1 Remark / Memo

            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4, 150, perItemHeight), sfAlignRight); // Debit
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 20 + 12 + 4, 150, perItemHeight), sfAlignRight); // Credit - bank

            // Debit Credit Total
            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 450 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(debitTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 450 - 5, secondTableYPos + 7, 150, tableHeight), sfAlignRight);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 25, 50 + 450 - 5 + 150, secondTableYPos + 25);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 28, 50 + 450 - 5 + 150, secondTableYPos + 28);

            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 600 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(creditTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 600 - 5, secondTableYPos + 7, 150, tableHeight), sfAlignRight);
            e.Graphics.DrawLine(Pens.Black, 50 + 600 + 5, secondTableYPos + 25, 50 + 600 - 5 + 150, secondTableYPos + 25);
            e.Graphics.DrawLine(Pens.Black, 50 + 600 + 5, secondTableYPos + 28, 50 + 600 - 5 + 150, secondTableYPos + 28);

            // Others Header
            int signWidth = 187;
            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 2, secondTableYPos + 45, signWidth, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 3, secondTableYPos + 45, signWidth, tableHeight - 20);

            e.Graphics.DrawString("Prepared By:", font_Header, Brushes.Black, new RectangleF(50, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Checked By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Approved By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Received By:", font_Header, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45, signWidth, tableHeight - 20), sfAlignCenter);

            // Others Data
            int othersYPos = secondTableYPos + 45 + 20 + 75;
            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 2, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + signWidth * 3, secondTableYPos + 45 + 20, signWidth, tableHeight + 35);

            AccessToDatabase accessToDatabase = new AccessToDatabase();

            var (PreparedByName, PreparedByPosition,
                ReviewedByName, ReviewedByPosition,
                RecommendingApprovalName, RecommendingApprovalPosition,
                ApprovedByName, ApprovedByPosition,
                ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

            e.Graphics.DrawString(PreparedByName, font_SevenBold, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(PreparedByPosition, font_Seven, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            e.Graphics.DrawString(ReviewedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ReviewedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            //e.Graphics.DrawString(data.RecommendingApprovalName, font_seven_Bold, Brushes.Black, new RectangleF(50 + 300, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            //e.Graphics.DrawString(data.RecommendingApprovalPosition, font_seven, Brushes.Black, new RectangleF(50 + 300, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);

            e.Graphics.DrawString(ApprovedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ApprovedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth * 2, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            e.Graphics.DrawString(ReceivedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45 + 60, signWidth, 20), sfAlignCenter);
            e.Graphics.DrawString(ReceivedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + signWidth * 3, secondTableYPos + 45 + 75, signWidth, 20), sfAlignCenter);

            // Received the amount of
            e.Graphics.DrawString("Received the amount of Php " + amount.ToString("N2"), font_Data, Brushes.Black, new PointF(50 + 430, othersYPos + 20));
            e.Graphics.DrawLine(Pens.Black, 50 + 435, othersYPos + 82, 50 + 605, othersYPos + 82); // line kanan sign han name
            e.Graphics.DrawString("(Sign over printed name)", font_Data, Brushes.Black, new PointF(50 + 455, othersYPos + 85));

            e.Graphics.DrawString("Date:", font_Data, Brushes.Black, new PointF(50 + 620, othersYPos + 82 - 15));
            e.Graphics.DrawLine(Pens.Black, 50 + 650, othersYPos + 82, 50 + 750, othersYPos + 82); // line kanan date

        }

    }
}
