using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using static VoucherPro.DataClass;
using static VoucherPro.AccessToDatabase;

namespace VoucherPro
{
    public class Layouts
    {
        private AccessToDatabase accessToDatabase;

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

        public void PrintPage(object sender, PrintPageEventArgs e, int layoutIndex)
        {
            StringFormat sfAlignRight = new StringFormat { Alignment = StringAlignment.Far | StringAlignment.Far };
            StringFormat sfAlignCenterRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignLeftCenter = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

            accessToDatabase = new AccessToDatabase();

            switch (layoutIndex)
            {
                case 1: // Check
                    Layout_Check(e, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                case 2: // Check Voucher
                    Layout_CheckVoucher_Check(e, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    break;
                case 3: // Accounts Payable Voucher
                    Layout_APVoucher(e, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                case 4: // Item Receipt
                    Layout_ItemReceipt(e, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                default:
                    throw new ArgumentException("Invalid layout index");
            }
        }

        private void Layout_Check(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {

        }

        private void Layout_CheckVoucher_Check(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, StringFormat sfAlignRight)
        {
            Font font_Data = font_Eight;
            Font font_Header = font_EightBold;

            string companyName = "Sample Company Name";
            string companyAddress = "Sample Address";
            string companyTelNo = "Telephone no. or Email";
            string cvText = "CHECK VOUCHER";

            // SHORT LOGO
            e.Graphics.DrawString(companyName, font_NineBold, Brushes.Black, new PointF(150 + 35, 50));
            e.Graphics.DrawString(companyAddress, font_Eight, Brushes.Black, new PointF(150 + 35, 65));
            e.Graphics.DrawString(companyTelNo, font_Eight, Brushes.Black, new PointF(150 + 35, 80));
            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));

            // LONG LOGO
            /*e.Graphics.DrawString(companyName, font_NineBold, Brushes.Black, new PointF(150 + 70, 50));
            e.Graphics.DrawString(companyAddress, font_Eight, Brushes.Black, new PointF(150 + 70, 65));
            e.Graphics.DrawString(companyTelNo, font_Eight, Brushes.Black, new PointF(150 + 70, 90));
            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));*/

            // 1st Table - Details
            int tableWidth = 750;
            int tableHeight = 40;
            int firstTableYPos = 180 + tableHeight + 7;
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 50, 150, tableHeight + 10); // CV Ref. No.
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 100, 150, tableHeight); // Print Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 140, tableWidth - 450, tableHeight); // Payee
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 450, 140, 150, tableHeight); // Bank
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 300, 140, 150, tableHeight); // Check Number
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 140, 150, tableHeight); // Check Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 180, tableWidth - 150, tableHeight); // Amount in Words
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 180, 150, tableHeight); // Amount

            // 1st Table Header
            e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 50 + 2, 150, tableHeight + 10));
            e.Graphics.DrawString("Print Date", font_Header, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 - 10 - 1, 150, tableHeight), sfAlignCenter);

            e.Graphics.DrawString("Payee", font_Header, Brushes.Black, new RectangleF(50 + 3, 140 + 2, tableWidth - 450, tableHeight));
            e.Graphics.DrawString("Bank", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 450, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Number", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 300, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Check Date", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 140 + 2, 150, tableHeight));

            e.Graphics.DrawString("Amount in Words", font_Header, Brushes.Black, new RectangleF(50 + 3, 180 + 2, tableWidth - 150, tableHeight));
            e.Graphics.DrawString("Amount", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 180 + 2, 150, tableHeight));

            // 1st Table Data
            string payee = "Sample payee";
            string bankAccount = " CIB-AUB Peso";
            string checkNumber = "0001";
            string seriesNumber = "CV0001";
            double amount = 1000.00;
            string amountInWords = AmountToWordsConverter.Convert(amount);
            DateTime chequeDate = DateTime.Now;
            DateTime dateTime = DateTime.Now; //PRINT DATE


            e.Graphics.DrawString(seriesNumber, font_TenBold, Brushes.Black, new RectangleF(50 + tableWidth - 150, 50 + 6, 150, tableHeight + 10), sfAlignCenter); // CV Ref. No.
            e.Graphics.DrawString(dateTime.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 + 8, 150, tableHeight), sfAlignCenter); // Print Date

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
            //int rows = Math.Min(checkData.Count, itemsLimit);
            int rows = Math.Min(8, itemsLimit);

            for (int i = 0; i < rows; i++)
            {
                secondTableHeight += perItemHeight;
            }
            //secondTableHeight -= 40;

            int secondTableYPos = firstTableYPos + 40 + secondTableHeight;

            double debitTotalAmount = 1000;
            double creditTotalAmount = 1000;

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

            //string particularAccount = checkData[0].AccountNumber + " - " + checkData[0].AccountName;
            /*string particularBank = checkData[0].BankAccount;
            string particularMemo = checkData[0].Memo;*/ // Remark or Memo

            int itemAccountHeight = 0;

            /*for (int i = 0; i < rows; i++)
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

            // Total Amount
            /*for (int i = 0; i < checkData.Count; i++)
            {
                if (checkData[i].ItemAmount != 0)
                {
                    if (checkData[i].ItemAmount > 0) // Debit
                    {
                        debitTotalAmount += checkData[i].ItemAmount;
                    }
                    else // Credit
                    {
                        double absoluteAmount = Math.Abs(checkData[i].ItemAmount);
                        creditTotalAmount += absoluteAmount;
                    }
                }
                else if (checkData[i].ExpensesAmount != 0)
                {
                    if (checkData[i].ExpensesAmount > 0) // Debit
                    {
                        debitTotalAmount += checkData[i].ExpensesAmount;
                    }
                    else // Credit
                    {
                        double absoluteAmount = Math.Abs(checkData[i].ExpensesAmount);
                        creditTotalAmount += absoluteAmount;
                    }
                }
            }
            creditTotalAmount += amount;*/

            //e.Graphics.DrawRectangle(Pens.Blue, 50 + 70, firstTableYPos + 50 - 5 + itemAccountHeight - 40, tableWidth - (300 + 170), 50); // Remark 
            //e.Graphics.DrawRectangle(Pens.Orange, 50 + 300 + 50, firstTableYPos + 20, 100, perItemHeight); // Class 
            //e.Graphics.DrawRectangle(Pens.Yellow, 50 + 450, firstTableYPos + 20, 150, perItemHeight); // Debit 
            //e.Graphics.DrawRectangle(Pens.Green, 50 + 600, firstTableYPos + 20, 150, perItemHeight); // Credit 

            //e.Graphics.DrawString(particularAccount, font_eight, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4, tableWidth - (300 + 100), perItemHeight)); // Item1
            /*e.Graphics.DrawString(particularBank, font_eight, Brushes.Black, new RectangleF(50 + 10 - 2, firstTableYPos + 30 + itemAccountHeight, tableWidth - (300 + 110), perItemHeight)); // Item1 Bank
            e.Graphics.DrawString(amount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 30 + itemAccountHeight, 150, perItemHeight), sfAlignRight); // Credit - bank

            e.Graphics.DrawString("*Remarks: ", font_eight, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 50 - 5 + itemAccountHeight, tableWidth - (300 + 100), 60)); // Item1 Remark / Memo
            e.Graphics.DrawString(particularMemo, font_seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 50 - 3 + itemAccountHeight, tableWidth - (300 + 170), 60)); // Item1 Remark / Memo
*/
            //e.Graphics.DrawString(amount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4, 150, perItemHeight), sfAlignRight); // Debit

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
            /*e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45, 180, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 180, secondTableYPos + 45, 180, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 360, secondTableYPos + 45, 180, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 540, secondTableYPos + 45, 210, tableHeight - 20);*/

            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45, 150, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 150, secondTableYPos + 45, 150, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 300, secondTableYPos + 45, 150, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 450, secondTableYPos + 45, 150, tableHeight - 20);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 600, secondTableYPos + 45, 150, tableHeight - 20);
            e.Graphics.DrawString("Prepared By:", font_Header, Brushes.Black, new RectangleF(50, secondTableYPos + 45, 150, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Reviewed By:", font_Header, Brushes.Black, new RectangleF(50 + 150, secondTableYPos + 45, 150, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Recommending Approval:", font_Header, Brushes.Black, new RectangleF(50 + 300, secondTableYPos + 45, 150, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Approved By:", font_Header, Brushes.Black, new RectangleF(50 + 450, secondTableYPos + 45, 150, tableHeight - 20), sfAlignCenter);
            e.Graphics.DrawString("Received By:", font_Header, Brushes.Black, new RectangleF(50 + 600, secondTableYPos + 45, 150, tableHeight - 20), sfAlignCenter);

            // Others Data
            int othersYPos = secondTableYPos + 45 + 20 + 75;

            e.Graphics.DrawRectangle(Pens.Black, 50, secondTableYPos + 45 + 20, 150, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 150, secondTableYPos + 45 + 20, 150, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 300, secondTableYPos + 45 + 20, 150, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 450, secondTableYPos + 45 + 20, 150, tableHeight + 35);
            e.Graphics.DrawRectangle(Pens.Black, 50 + 600, secondTableYPos + 45 + 20, 150, tableHeight + 35);


            var data = accessToDatabase.RetrieveAllSignatoryData();

            e.Graphics.DrawString(data.PreparedByName, font_SevenBold, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            e.Graphics.DrawString(data.PreparedByPosition, font_Seven, Brushes.Black, new RectangleF(50, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);

            e.Graphics.DrawString(data.ReviewedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + 150, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            e.Graphics.DrawString(data.ReviewedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + 150, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);

            e.Graphics.DrawString(data.RecommendingApprovalName, font_SevenBold, Brushes.Black, new RectangleF(50 + 300, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            e.Graphics.DrawString(data.RecommendingApprovalPosition, font_Seven, Brushes.Black, new RectangleF(50 + 300, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);

            e.Graphics.DrawString(data.ApprovedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + 450, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            e.Graphics.DrawString(data.ApprovedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + 450, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);

            e.Graphics.DrawString(data.ReceivedByName, font_SevenBold, Brushes.Black, new RectangleF(50 + 600, secondTableYPos + 45 + 60, 150, 20), sfAlignCenter);
            e.Graphics.DrawString(data.ReceivedByPosition, font_Seven, Brushes.Black, new RectangleF(50 + 600, secondTableYPos + 45 + 75, 150, 20), sfAlignCenter);


            // Received the amount of
            //e.Graphics.DrawString("Received the amount of Php " + amount.ToString("N2"), font_eight, Brushes.Black, new PointF(50 + 430, othersYPos + 20));
            e.Graphics.DrawString("Received the amount of Php 1000.00", font_Data, Brushes.Black, new PointF(50 + 430, othersYPos + 20));
            e.Graphics.DrawLine(Pens.Black, 50 + 435, othersYPos + 82, 50 + 605, othersYPos + 82); // line kanan sign han name
            e.Graphics.DrawString("(Sign over printed name)", font_Data, Brushes.Black, new PointF(50 + 455, othersYPos + 85));

            e.Graphics.DrawString("Date:", font_Data, Brushes.Black, new PointF(50 + 620, othersYPos + 82 - 15));
            e.Graphics.DrawLine(Pens.Black, 50 + 650, othersYPos + 82, 50 + 750, othersYPos + 82); // line kanan date

        }

        private void Layout_CheckVoucher_Bill(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {
            Font font_Data = font_Eight;
            Font font_Header = font_EightBold;
        }

        private void Layout_APVoucher(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {

        }

        private void Layout_ItemReceipt(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {

        }
    }
}
