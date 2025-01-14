using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static VoucherPro.DataClass;
using static VoucherPro.AccessToDatabase;
using System.Drawing.Drawing2D;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//using static System.Net.Mime.MediaTypeNames;
namespace VoucherPro.Clients
{
    public class Layouts_LEADS
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

        public void PrintPage_LEADS(object sender, PrintPageEventArgs e, int layoutIndex, string seriesNumber, object data)
        {
            StringFormat sfAlignRight = new StringFormat { Alignment = StringAlignment.Far | StringAlignment.Far };
            StringFormat sfAlignCenterRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignLeftCenter = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

            switch (layoutIndex)
            {
                case 1:
                    Layout_Check(e, data as List<CheckTable>, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                case 2:
                    if (data is List<CheckTableExpensesAndItems>)
                    {
                        Layout_CheckVoucher_Check(e, data as List<CheckTableExpensesAndItems>, seriesNumber, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    }
                    else if (data is List<BillTable>)
                    {
                        Layout_CheckVoucher_Bill(e, data as List<BillTable>, seriesNumber, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    }
                    break;
                case 3:
                    Layout_APVoucher(e, data as List<BillTable>, seriesNumber, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, sfAlignRight);
                    break;
                case 4:
                    Layout_ItemReceipt(e, receiptData: (List<ItemReciept>)data, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                default:
                    throw new ArgumentException("Invalid layout index");
            }
        }

        private void Layout_Check(PrintPageEventArgs e, List<CheckTable> checkTableData, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {
            DateTime dateCreated = Convert.ToDateTime(checkTableData[0].DateCreated).Date;

            string month = dateCreated.ToString("MM");
            string day = dateCreated.ToString("dd");
            string year = dateCreated.ToString("yyyy");

            // Insert spaces between characters
            string formattedMonth = string.Join("   ", month.ToCharArray());
            string formattedDay = string.Join("   ", day.ToCharArray());
            string formattedYear = string.Join("   ", year.ToCharArray());

            // Combine the parts with additional spaces between sections
            string formattedDate = $"{formattedMonth}     {formattedDay}     {formattedYear}";

            /*string formattedDate = dateCreated.ToString("MM    dd     yyyy");
            formattedDate = string.Join(" ", formattedDate.Select(c => c.ToString()));*/

            string payee = checkTableData[0].PayeeFullName.ToString();
            double amount = checkTableData[0].Amount;
            string amountInWords = AmountToWordsConverter.Convert(amount);

            Font amountinWordsFont = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            Font dateFont = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
            Font payeeFont = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
            /*
                e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550 + 50, 60 + 10)); // amount x 600 y 93
                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(90 - 30 + 50, 60 + 10)); // payee x 110 y 90
                e.Graphics.DrawString(formattedDate, dateFont, Brushes.Black, new PointF(520 + 60, 30 + 10)); // date x 570 y 60
                e.Graphics.DrawString(amountInWords, dateFont, Brushes.Black, new PointF(55 - 30 + 50, 90 + 10)); // amountinwords x 75 y 120
            */
            // Rotate the content -90 degrees
            if (GlobalVariables.isPrinting)
            {
                e.Graphics.RotateTransform(-90);
                e.Graphics.TranslateTransform(-e.MarginBounds.Height + 180, 0 - 70);

                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(60, 400 + 10));
                e.Graphics.DrawString(formattedDate, dateFont, Brushes.Black, new PointF(520 + 10, 370 + 10));
                e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550, 38 + 345 + 30));
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(25, 430 + 15));
            }
            else
            {
                /*e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550 + 50, 38 + 70 - 15));
                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(90 - 30 + 50, 20 + 70));
                e.Graphics.DrawString(formattedDate, dateFont, Brushes.Black, new PointF(520 + 3 + 50, 20 + 70 - 30));
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(55 - 30 + 50, 70 + 50));*/

                /*e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550, 38 + 380 - 15));
                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(90 - 30, 20 + 380));
                e.Graphics.DrawString(formattedDate, dateFont, Brushes.Black, new PointF(520 + 3, 20 + 380 - 30));
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(55 - 30, 50 + 380));*/
                int minusX = 30;
                int minusY = 35 + 15;

                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(155 - minusX, 110 - minusY));
                e.Graphics.DrawString(formattedDate, payeeFont, Brushes.Black, new PointF(625 + 3 - minusX, 75 + 4 - minusY));
                e.Graphics.DrawString(amount.ToString("N2"), payeeFont, Brushes.Black, new PointF(520 + 115 - minusX, 110 + 4 - minusY));
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(125 - minusX, 145 - minusY));
            }
        }

        private void Layout_CheckVoucher_Check(PrintPageEventArgs e, List<CheckTableExpensesAndItems> checkData, string seriesNumber, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, StringFormat sfAlignRight)
        {
            Font font_Header = font_EightBold;
            Font font_Data = font_Eight;

            string companyName = "LEADS ENVIRONMENTAL HEALTH PRODUCTS CORP.";
            string companyTIN = "VAT Reg. TIN: 243-354-422-00000";
            string companyAddress = "LOT 14-A BLOCK 83 RODEO DRIVE, LAGUNA BEL AIR 2,\nBRGY. DON JOSE, 4026 CITY OF SANTA ROSA, LAGUNA, PHILIPPINES";
            string companyTelNo = "Tel. No.: (049) 501-8125";
            string cvText = "CHECK VOUCHER";

            Image image = Properties.Resources.leads_logo2;
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

            e.Graphics.DrawString(companyName, font_NineBold, Brushes.Black, new PointF(200, 50));
            e.Graphics.DrawString(companyTIN, font_Eight, Brushes.Black, new PointF(200, 65));
            e.Graphics.DrawString(companyAddress, font_Seven, Brushes.Black, new PointF(200, 80));
            e.Graphics.DrawString(companyTelNo, font_Seven, Brushes.Black, new PointF(200, 106));

            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));

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
            string payee = checkData[0].PayeeFullName;
            string bankAccount = checkData[0].BankAccount;
            string checkNumber = checkData[0].RefNumber;
            double amount = checkData[0].TotalAmount;
            string amountInWords = AmountToWordsConverter.Convert(amount);
            DateTime chequeDate = checkData[0].DateCreated;
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

            Dictionary<string, double> groupedItemData = new Dictionary<string, double>();
            //Dictionary<string, double> groupedExpenseData = new Dictionary<string, double>();

            try
            {
                for (int i = 0; i < checkData.Count; i++)
                {
                    string itemName = checkData[i].AccountName;
                    double itemAmount = checkData[i].ItemAmount;

                    if (itemName != "" && itemAmount != 0)
                    {
                        if (groupedItemData.ContainsKey(itemName))
                        {
                            groupedItemData[itemName] += itemAmount;
                        }
                        else
                        {
                            groupedItemData[itemName] = itemAmount;
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
                //Console.WriteLine($"Account Name: {entry.Key}, Total Amount: {entry.Value}");
                e.Graphics.DrawString($"{item.Key}", font_Nine, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 100), 25));
                e.Graphics.DrawString($"{item.Value:N2}", font_Nine, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + pos, 150, perItemHeight), sfAlignRight); // Credit

                if (item.Value > 0)
                {
                    debitTotalAmount += item.Value;
                }
                pos += 25;
            }

            amountPos += pos;

            foreach (var check in checkData)
            {
                e.Graphics.DrawString(check.Account, font_Data, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 100), perItemHeight)); // Item

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

            e.Graphics.DrawString("*Remarks: ", font_Data, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 50 - 5 + pos, tableWidth - (300 + 100), 60)); // Item1 Remark / Memo
            e.Graphics.DrawString(particularMemo, font_Seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 50 - 3 + pos, tableWidth - (300 + 170), 60)); // Item1 Remark / Memo

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

            string companyName = "LEADS ENVIRONMENTAL HEALTH PRODUCTS CORP.";
            string companyTIN = "VAT Reg. TIN: 243-354-422-00000";
            string companyAddress = "LOT 14-A BLOCK 83 RODEO DRIVE, LAGUNA BEL AIR 2,\nBRGY. DON JOSE, 4026 CITY OF SANTA ROSA, LAGUNA, PHILIPPINES";
            string companyTelNo = "Tel. No.: (049) 501-8125";
            string cvText = "CHECK VOUCHER";

            Image image = Properties.Resources.leads_logo2;
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

            e.Graphics.DrawString(companyName, font_NineBold, Brushes.Black, new PointF(200, 50));
            e.Graphics.DrawString(companyTIN, font_Eight, Brushes.Black, new PointF(200, 65));
            e.Graphics.DrawString(companyAddress, font_Seven, Brushes.Black, new PointF(200, 80));
            e.Graphics.DrawString(companyTelNo, font_Seven, Brushes.Black, new PointF(200, 106));

            e.Graphics.DrawString(cvText, font_TwelveBold, Brushes.Black, new PointF(500 - 15, 110 + 5));

            // 1st Table - Details
            int tableWidth = 750;
            int tableHeight = 40;
            int firstTableYPos = 180 + tableHeight + 7;

            int payeeWidth = tableWidth - 475; // 450
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 50, 150, tableHeight + 10); // CV Ref. No.
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 100, 150, tableHeight); // Print Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 140, payeeWidth, tableHeight); // Payee
            e.Graphics.DrawRectangle(Pens.Black, 50 + payeeWidth, 140, 150 + 25, tableHeight); // Bank
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 300, 140, 150, tableHeight); // Check Number
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 140, 150, tableHeight); // Check Date

            e.Graphics.DrawRectangle(Pens.Black, 50, 180, tableWidth - 150, tableHeight); // Amount in Words
            e.Graphics.DrawRectangle(Pens.Black, 50 + tableWidth - 150, 180, 150, tableHeight); // Amount

            // 1st Table Header
            e.Graphics.DrawString("CV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 50 + 2, 150, tableHeight + 10));
            e.Graphics.DrawString("Print Date", font_Header, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 - 10 - 1, 150, tableHeight), sfAlignCenter);

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


            e.Graphics.DrawString(seriesNumber, font_TenBold, Brushes.Black, new RectangleF(50 + tableWidth - 150, 50 + 6, 150, tableHeight + 10), sfAlignCenter); // CV Ref. No.
            e.Graphics.DrawString(dateTime.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 + 8, 150, tableHeight), sfAlignCenter); // Print Date

            e.Graphics.DrawString(payee, font_Data, Brushes.Black, new RectangleF(50 + 15, 140 + 6, payeeWidth, tableHeight), sfAlignLeftCenter); // Payee
            e.Graphics.DrawString(bankAccount, font_Data, Brushes.Black, new RectangleF(50 + payeeWidth, 140 + 6, 150 + 25, tableHeight), sfAlignCenter); // Bank
            e.Graphics.DrawString(checkNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 300, 140 + 6, 150, tableHeight), sfAlignCenter); // Check Number
            e.Graphics.DrawString(chequeDate.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 140 + 8, 150, tableHeight), sfAlignCenter); // Check Date

            e.Graphics.DrawString(amountInWords, font_Data, Brushes.Black, new RectangleF(50 + 15, 180 + 6, tableWidth - 150, tableHeight), sfAlignLeftCenter); // Amount in words
            e.Graphics.DrawString("₱", font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 + 10, 180 + 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 - 10, 180 + 6, 150, tableHeight), sfAlignCenterRight); // Amount

            // 2nd Table - Particulars
            int secondTableHeight = 60; // 75
            int secondTableYPos = firstTableYPos + 40 + secondTableHeight;

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
                particularAccount = billData[0].APAccountRefFullName + " (Bill#: " + billData[0].AppliedRefNumber + ")";
            }
            else
            {
                particularAccount = billData[0].APAccountRefFullName + " (Bill#: " + billData[0].AppliedRefNumber + " / " + "APV#: " + apvFieldText + ")";
            }


            string particularBank = billData[0].BankAccount;
            string particularMemo = billData[0].Memo; // Remark or Memo

            //e.Graphics.DrawRectangle(Pens.Red, 50, firstTableYPos + 20, tableWidth - (300 + 100), perItemHeight); // Particular 
            //e.Graphics.DrawRectangle(Pens.Blue, 50 + 70, firstTableYPos + 50, tableWidth - (300 + 170), perItemHeight - 30); // Remark 
            //e.Graphics.DrawRectangle(Pens.Orange, 50 + 300 + 50, firstTableYPos + 20, 100, perItemHeight); // Class 
            //e.Graphics.DrawRectangle(Pens.Yellow, 50 + 450, firstTableYPos + 20, 150, perItemHeight); // Debit 
            //e.Graphics.DrawRectangle(Pens.Green, 50 + 600, firstTableYPos + 20, 150, perItemHeight); // Credit 

            e.Graphics.DrawString(particularAccount, font_Data, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4, tableWidth - (300 + 100), perItemHeight)); // Item1
            e.Graphics.DrawString(particularBank, font_Data, Brushes.Black, new RectangleF(50 + 15, firstTableYPos + 20 + 12 + 4, tableWidth - (300 + 110), perItemHeight)); // Item1 Bank
            e.Graphics.DrawString("*Remarks: ", font_Data, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 20 + 30 + 4, tableWidth - (300 + 100), perItemHeight - 30)); // Item1 Remark / Memo
            e.Graphics.DrawString(particularMemo, font_Seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 50 + 4, tableWidth - (300 + 170), perItemHeight)); // Item1 Remark / Memo

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

        private void Layout_APVoucher(PrintPageEventArgs e, List<BillTable> billData, string seriesNumber, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, StringFormat sfAlignRight)
        {
            Font font_Header = font_EightBold;
            Font font_Data = font_Eight;

            string companyName = "LEADS ENVIRONMENTAL HEALTH PRODUCTS CORP.";
            string companyTIN = "VAT Reg. TIN: 243-354-422-00000";
            string companyAddress = "LOT 14-A BLOCK 83 RODEO DRIVE, LAGUNA BEL AIR 2,\nBRGY. DON JOSE, 4026 CITY OF SANTA ROSA, LAGUNA, PHILIPPINES";
            string companyTelNo = "Tel. No.: (049) 501-8125";
            string apvText = "ACCOUNTS PAYABLE VOUCHER";

            Image image = Properties.Resources.leads_logo2;
            Bitmap resizedBitmap = null;

            if (image != null)
            {
                int imageWidth = 140; // 90
                int imageHeight = (int)((double)image.Height / image.Width * imageWidth);
                resizedBitmap = new Bitmap(image, new Size(imageWidth, imageHeight));
            }

            if (resizedBitmap != null)
            {
                e.Graphics.DrawImage(resizedBitmap, new PointF(50, 40)); // an logo
            }

            e.Graphics.DrawString(companyName, font_NineBold, Brushes.Black, new PointF(200, 50));
            e.Graphics.DrawString(companyTIN, font_Eight, Brushes.Black, new PointF(200, 65));
            e.Graphics.DrawString(companyAddress, font_Seven, Brushes.Black, new PointF(200, 80));
            e.Graphics.DrawString(companyTelNo, font_Seven, Brushes.Black, new PointF(200, 106));
            e.Graphics.DrawString(apvText, font_TwelveBold, Brushes.Black, new PointF(370 - 15, 110 + 5));

            bool isPaid = billData[0].IsPaid;
            if (isPaid)
            {
                string text = "PAID";
                Font font = new Font("Stencil", 28, FontStyle.Bold);

                // Save the current state of the Graphics object
                GraphicsState state = e.Graphics.Save();

                // Apply a rotation transformation
                e.Graphics.TranslateTransform(540, 88); // Move the origin to the specified position
                e.Graphics.RotateTransform(-30); // Rotate the text by -45 degrees (adjust angle as needed)

                e.Graphics.DrawString(text, font, Brushes.Gray, 0, 0);

                // Restore the original state of the Graphics object
                e.Graphics.Restore(state);
            }

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
            e.Graphics.DrawString("APV Ref. No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 50 + 2, 150, tableHeight + 10));
            e.Graphics.DrawString("Bill Date", font_Header, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 - 10 - 1, 150, tableHeight), sfAlignCenter);

            e.Graphics.DrawString("Payee", font_Header, Brushes.Black, new RectangleF(50 + 3, 140 + 2, tableWidth - 450, tableHeight));
            e.Graphics.DrawString("Terms", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 450, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Bill No.", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 300, 140 + 2, 150, tableHeight));
            e.Graphics.DrawString("Due Date", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 140 + 2, 150, tableHeight));

            e.Graphics.DrawString("Amount in Words", font_Header, Brushes.Black, new RectangleF(50 + 3, 180 + 2, tableWidth - 150, tableHeight));
            e.Graphics.DrawString("Amount", font_Header, Brushes.Black, new RectangleF(50 + 3 + tableWidth - 150, 180 + 2, 150, tableHeight));

            // 1st Table Data
            string payee = billData[0].Vendor;
            string terms = billData[0].TermsRefFullName;
            string billNumber = billData[0].RefNumber;
            double amount = billData[0].AmountDue;
            string amountInWords = AmountToWordsConverter.Convert(amount);
            DateTime billDate = billData[0].DateCreated; //BILL DATE
            DateTime dueDate = billData[0].DueDate;


            e.Graphics.DrawString(seriesNumber, font_TenBold, Brushes.Black, new RectangleF(50 + tableWidth - 150, 50 + 6, 150, tableHeight + 10), sfAlignCenter); // APV Ref. No.
            e.Graphics.DrawString(billDate.ToString("dd-MMM-yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 100 + 8, 150, tableHeight), sfAlignCenter); // Bill Date

            e.Graphics.DrawString(payee, font_Data, Brushes.Black, new RectangleF(50 + 15, 140 + 6, tableWidth - 450, tableHeight), sfAlignLeftCenter); // Payee
            e.Graphics.DrawString(terms, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 450, 140 + 6, 150, tableHeight), sfAlignCenter); // Terms
            e.Graphics.DrawString(billNumber, font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 300, 140 + 6, 150, tableHeight), sfAlignCenter); // Bill No.
            e.Graphics.DrawString(dueDate.ToString("MM/dd/yyyy"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150, 140 + 8, 150, tableHeight), sfAlignCenter); // Due Date

            e.Graphics.DrawString(amountInWords, font_Data, Brushes.Black, new RectangleF(50 + 15, 180 + 6, tableWidth - 150, tableHeight), sfAlignLeftCenter); // Amount in words
            e.Graphics.DrawString("₱", font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 + 10, 180 + 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + tableWidth - 150 - 10, 180 + 6, 150, tableHeight), sfAlignCenterRight); // Amount

            // 2nd Table - Particulars
            int secondTableHeight = 75 - 14; // 75

            // Dictionary to store the grouped amounts by AccountNameParticulars
            /*Dictionary<string, double> groupedItemData = new Dictionary<string, double>();
            Dictionary<string, double> groupedExpenseData = new Dictionary<string, double>();

            int totalItemCount = 0;
            
            foreach (var bill in billData)
            {
                try
                {
                    for (int i = 0; i < bill.AccountNameParticularsList.Count; i++)
                    {
                        //string itemName = bill.AccountNameParticularsList[i];
                        string itemName = bill.ItemDetails[i].ItemLineItemRefFullName;
                        double itemAmount = bill.ItemDetails[i].ItemLineAmount;
                        string itemClass = bill.ItemDetails[i].ItemLineClassRefFullName;

                        if (groupedItemData.ContainsKey(itemName))
                        {
                            groupedItemData[itemName] += itemAmount;
                        }
                        else
                        {
                            groupedItemData[itemName] = itemAmount;
                        }
                        totalItemCount++;
                    }

                    foreach (var item in bill.ItemDetails)
                    {
                        if (!string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                        {
                            string expenseName = item.ExpenseLineItemRefFullName;
                            double expenseAmount = item.ExpenseLineAmount;

                            if (groupedExpenseData.ContainsKey(expenseName))
                            {
                                groupedExpenseData[expenseName] += expenseAmount;
                            }
                            else
                            {
                                groupedExpenseData[expenseName] = expenseAmount;
                            }
                            totalItemCount++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while grouping entries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                Console.WriteLine($"Total Count: {totalItemCount}");
            }*/

            Dictionary<string, List<APVData>> groupedItemData = new Dictionary<string, List<APVData>>();
            Dictionary<string, List<APVData>> groupedExpenseData = new Dictionary<string, List<APVData>>();

            int totalItemCount = 0;

            foreach (var bill in billData)
            {
                try
                {
                    for (int i = 0; i < bill.AccountNameParticularsList.Count; i++)
                    {
                        string itemName = bill.ItemDetails[i].ItemLineItemRefFullName;
                        double itemAmount = bill.ItemDetails[i].ItemLineAmount;
                        string itemClass = bill.ItemDetails[i].ItemLineClassRefFullName;

                        if (!groupedItemData.ContainsKey(itemName))
                        {
                            groupedItemData[itemName] = new List<APVData>();
                        }
                        groupedItemData[itemName].Add(new APVData { Amount = itemAmount, Class = itemClass });

                        totalItemCount++;
                    }

                    foreach (var item in bill.ItemDetails)
                    {
                        if (!string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                        {
                            string expenseName = item.ExpenseLineItemRefFullName;
                            double expenseAmount = item.ExpenseLineAmount;
                            string expenseClass = item.ExpenseLineClassRefFullName;

                            if (!groupedExpenseData.ContainsKey(expenseName))
                            {
                                groupedExpenseData[expenseName] = new List<APVData>();
                            }
                            groupedExpenseData[expenseName].Add(new APVData { Amount = expenseAmount, Class = expenseClass });

                            totalItemCount++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while grouping entries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                Console.WriteLine($"Total Count: {totalItemCount}");
            }

            /*foreach (var bill in billData)
            {
                foreach (var c in bill.ItemDetails)
                {
                    secondTableHeight += 25;
                }
            }*/
            try
            {
                /*foreach (var bill in billData)
                {
                    foreach (var entry in groupedItemData)
                    {
                        secondTableHeight += 25;
                    }

                    foreach (var entry in groupedExpenseData)
                    {
                        secondTableHeight += 25;
                    }
                }*/
                for (int i = 0; i < totalItemCount; i++)
                {
                    secondTableHeight += 25;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while adding height: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            int secondTableYPos = firstTableYPos + 40 + secondTableHeight;


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

            //string particularAccount = billData[0].AccountNumberParticulars + " - " + billData[0].AccountNameParticulars;
            string particularAPAccount = billData[0].AccountNumber + " - " + billData[0].AccountName;
            //string particularBank = checkData[0].BankAccount;
            string particularMemo = billData[0].Memo; // Remark or Memo

            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            int pos = 0;
            int amountPos = 0;

            // Assuming these variables are defined somewhere in your class
            /*int pageCounter = 0;
            int itemCounter = 0;
            int index = 0;
            int count = 0;*/

            /*int itemsPerPage = 10;
            int currentPage = 0;
            var paginatedData = billData.Skip(currentPage * itemsPerPage).Take(itemsPerPage).ToList();

            foreach (var bill in paginatedData)
            {
                foreach (var item in groupedItemData)
                {
                    if (itemCounter >= GlobalVariables.itemsPerPageAPV) break; // Stop if we reach the items per page limit
                    Console.WriteLine($"Grouped Items: Account Name: {item.Key}, Total Amount: {item.Value}, Item Counter: {itemCounter}");
                    itemCounter++;
                    count++;
                }

                // Display grouped expenses
                foreach (var item in groupedExpenseData)
                {
                    if (itemCounter >= GlobalVariables.itemsPerPageAPV) break; // Stop if we reach the items per page limit
                    Console.WriteLine($"Grouped Expense: Account Name: {item.Key}, Total Amount: {item.Value}, Item Counter: {itemCounter}");
                    itemCounter++;
                    count++;
                }
            }
            // Indicate if there are more pages
            if (currentPage < 2 - 1)
            {
                e.HasMorePages = true;
            }*/

            // Loop to implement pagination and break when it's time for the next page
            /*while (itemCounter < totalItemCount && count < GlobalVariables.itemsPerPageAPV)
            {
                // Reset the count for the current page
                count = 0;

                // Display grouped items
                foreach (var item in groupedItemData)
                {
                    if (itemCounter >= GlobalVariables.itemsPerPageAPV) break; // Stop if we reach the items per page limit
                    Console.WriteLine($"Grouped Items: Account Name: {item.Key}, Total Amount: {item.Value}, Item Counter: {itemCounter}");
                    itemCounter++;
                    count++;
                }

                // Display grouped expenses
                foreach (var item in groupedExpenseData)
                {
                    if (itemCounter >= GlobalVariables.itemsPerPageAPV) break; // Stop if we reach the items per page limit
                    Console.WriteLine($"Grouped Expense: Account Name: {item.Key}, Total Amount: {item.Value}, Item Counter: {itemCounter}");
                    itemCounter++;
                    count++;
                }

                // If we have printed all items for the current page
                if (itemCounter < totalItemCount)
                {
                    pageCounter++;
                    e.HasMorePages = true; // Indicates more pages to print
                    Console.WriteLine($"Page Counter add: {pageCounter}");
                }
                else
                {
                    e.HasMorePages = false; // No more pages to print
                    Console.WriteLine($"Page Counter end: {pageCounter}");
                }

                // Break out of the loop after the current page is processed
                //if (!e.HasMorePages) break;
            }*/

            foreach (var bill in billData)
            {
                try
                {
                    for (int i = 0; i < bill.AccountNameParticularsList.Count; i++)
                    {
                        //string itemName = bill.AccountNameParticularsList[i];
                        string itemName = bill.ItemDetails[i].ItemLineItemRefFullName;
                        string itemClass = bill.ItemDetails[i].ItemLineClassRefFullName;
                        double itemAmount = bill.ItemDetails[i].ItemLineAmount;

                        if (itemAmount < 0)
                        {
                            double absoluteAmount = Math.Abs(itemAmount);
                            creditTotalAmount += absoluteAmount;
                        }
                        else if (itemAmount > 0)
                        {
                            debitTotalAmount += itemAmount;
                        }

                        Console.WriteLine($"Items: Item Name: {itemName}, Item Price: {itemAmount}, Credit: {creditTotalAmount}, Debit: {debitTotalAmount}");
                    }

                    foreach (var item in bill.ItemDetails)
                    {
                        if (!string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                        {
                            string expenseName = item.ExpenseLineItemRefFullName;
                            string expenseClass = item.ExpenseLineClassRefFullName;
                            double expenseAmount = item.ExpenseLineAmount;

                            if (expenseAmount > 0)
                            {
                                debitTotalAmount += expenseAmount;
                            }

                            Console.WriteLine($"Expenses: Account Name: {expenseName}, Expense Amount: {expenseAmount}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while printing entries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            //e.HasMorePages = currentPrintIndex < itemCounter;

            // adi main
            /*foreach (var bill in billData)
            {
                foreach (var item in groupedItemData)
                {
                    //Console.WriteLine($"Grouped Items: Account Name: {item.Key}, Total Amount: {item.Value}");
                    e.Graphics.DrawString($"{item.Key}", font_Nine, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 100), 25));
                    e.Graphics.DrawString($"{item.Value:N2}", font_Nine, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + pos, 150, perItemHeight), sfAlignRight); // Credit

                    if (item.Value > 0)
                    {
                        debitTotalAmount += item.Value;
                    }
                    pos += 25;
                }

                amountPos += pos + 25;

                foreach (var items in groupedExpenseData)
                {
                    if (items.Key != "")
                    {
                        //Console.WriteLine($"Grouped Expense: Account Name: {items.Key}, Total Amount: {items.Value}");
                        e.Graphics.DrawString(items.Key, font_Nine, Brushes.Black, new RectangleF(50 + 4, firstTableYPos + amountPos, tableWidth - (300 + 100), perItemHeight));
                        if (items.Value < 0)
                        {
                            double absoluteAmount = Math.Abs(items.Value);
                            e.Graphics.DrawString(absoluteAmount.ToString("N2"), font_Nine, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + amountPos, 150, perItemHeight), sfAlignRight); // Credit
                            creditTotalAmount += absoluteAmount;
                        }
                        else if (items.Value > 0)
                        {
                            e.Graphics.DrawString(items.Value.ToString("N2"), font_Nine, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + amountPos, 150, perItemHeight), sfAlignRight); // Debit
                            debitTotalAmount += items.Value;
                        }
                    }
                    //pos += 25;
                    amountPos += 25;
                }
            }*/

            // latest 09-01-25
            /*foreach (var bill in billData)
            {
                foreach (var item in groupedItemData)
                {
                    //Console.WriteLine($"Grouped Items: Account Name: {item.Key}, Total Amount: {item.Value}");
                    foreach (var data in item.Value) // Iterate through each ItemData object for the current key
                    {
                        e.Graphics.DrawString($"{item.Key}", font_Nine, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 100), 25));
                        e.Graphics.DrawString($"{data.Class}", font_Nine, Brushes.Black, new RectangleF(50 + 300 + 50 + 3, firstTableYPos + 20 + 4 + pos, 100, 25));
                        e.Graphics.DrawString($"{data.Amount:N2}", font_Nine, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4 + pos, 150, perItemHeight), sfAlignRight);

                        if (data.Amount > 0)
                        {
                            debitTotalAmount += data.Amount;
                        }
                        pos += 25;
                    }
                }

                amountPos += pos + 25;

                foreach (var expense in groupedExpenseData)
                {
                    foreach (var data in expense.Value) // Iterate through each ItemData object for the current expense key
                    {
                        if (!string.IsNullOrEmpty(expense.Key))
                        {
                            e.Graphics.DrawString(expense.Key, font_Nine, Brushes.Black, new RectangleF(50 + 4, firstTableYPos + amountPos, tableWidth - (300 + 100), perItemHeight));
                            e.Graphics.DrawString($"{data.Class}", font_Nine, Brushes.Black, new RectangleF(50 + 300 + 50 + 3, firstTableYPos + amountPos, 100, perItemHeight));
                            
                            if (data.Amount < 0)
                            {
                                double absoluteAmount = Math.Abs(data.Amount);
                                e.Graphics.DrawString(absoluteAmount.ToString("N2"), font_Nine, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + amountPos, 150, perItemHeight), sfAlignRight); // Credit
                                creditTotalAmount += absoluteAmount;
                            }
                            else if (data.Amount > 0)
                            {
                                e.Graphics.DrawString(data.Amount.ToString("N2"), font_Nine, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + amountPos, 150, perItemHeight), sfAlignRight); // Debit
                                debitTotalAmount += data.Amount;
                            }

                            amountPos += 25;
                        }
                    }
                }
            }*/

            int totalPages = (int)Math.Ceiling((double)totalItemCount / 10);

            int pageCounter = 1;
            int itemCounter = 0;
            int totalItemCounter = 0;
            //int xkek = 0;
            Console.WriteLine($"Initial Page Count: {pageCounter}, Item Count: {itemCounter}, Total Item Count: {totalItemCounter}, Total Page Count (supposedly) {totalPages}");
            foreach (var bill in billData)
            {
                for (int i = 0; i < bill.ItemDetails.Count; i++)
                {
                    if (itemCounter < 10 && totalItemCounter < bill.ItemDetails.Count)
                    {
                        e.Graphics.DrawString($"{bill.ItemDetails[i].ItemLineItemRefFullName}", font_Nine, Brushes.Black, new RectangleF(50 + 5, firstTableYPos + 20 + 4 + pos, tableWidth - (300 + 100), 25));

                        pos += 25;
                        itemCounter++;
                        totalItemCounter++;

                        Console.WriteLine($"I = {i}, Page Count: {pageCounter}, Item Count: {itemCounter}, Total Item Count: {totalItemCounter}");
                        e.HasMorePages = false;
                    }
                    //pageCounter < totalPages && 
                    else if (itemCounter >= 10 && totalItemCounter < bill.ItemDetails.Count)
                    {
                        pos = 0;
                        itemCounter = 0;
                        i--;
                        pageCounter++;
                        //e.HasMorePages = pageCounter < totalPages;
                        e.HasMorePages = true;
                        if (e.HasMorePages)
                        {
                            return;
                        }
                        else
                        {
                            e.HasMorePages = false;
                            break;
                        }
                    }
                }
            }
            //e.HasMorePages = false;
            /*if (pageCounter < totalPages) // Check if there are more pages to print
            {
                e.Graphics.DrawString($"{pageCounter}", new Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 60, 10);
                pageCounter++; // Move to the next page
                e.HasMorePages = pageCounter < totalPages; // Set HasMorePages based on the current page index
            }
            else
            {
                e.HasMorePages = false; // No more pages to print
            }*/
            //e.HasMorePages = false;


            creditTotalAmount += amount;

            e.Graphics.DrawString(particularAPAccount, font_Data, Brushes.Black, new RectangleF(50 + 15, firstTableYPos + amountPos, tableWidth - (300 + 110), perItemHeight)); // Item1 AP Account
            e.Graphics.DrawString("*Remarks: ", font_Data, Brushes.Black, new RectangleF(50 + 10, firstTableYPos + 15 + amountPos, tableWidth - (300 + 100), perItemHeight - 30)); // Item1 Remark / Memo
            e.Graphics.DrawString(amount.ToString("N2"), font_Data, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + amountPos, 150, perItemHeight), sfAlignRight); // Credit - bank
            e.Graphics.DrawString(particularMemo, font_Seven, Brushes.Black, new RectangleF(50 + 75, firstTableYPos + 20 - 4 + amountPos, tableWidth - (300 + 170), 45)); // Item1 Remark / Memo

            //e.Graphics.DrawString(amount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 450 - 5, firstTableYPos + 20 + 4, 150, perItemHeight), sfAlignRight); // Debit
            //e.Graphics.DrawString(amount.ToString("N2"), font_eight, Brushes.Black, new RectangleF(50 + 600 - 5, firstTableYPos + 20 + 12 + 4, 150, perItemHeight), sfAlignRight); // Credit - bank

            // Debit Credit Total
            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 450 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(debitTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 450 - 5, secondTableYPos + 9, 150, tableHeight), sfAlignRight);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 25, 50 + 450 - 5 + 150, secondTableYPos + 25);
            e.Graphics.DrawLine(Pens.Black, 50 + 450 + 5, secondTableYPos + 28, 50 + 450 - 5 + 150, secondTableYPos + 28);

            e.Graphics.DrawString("₱", font_Header, Brushes.Black, new RectangleF(50 + 600 + 5, secondTableYPos - 6, 150, tableHeight), sfAlignLeftCenter);
            e.Graphics.DrawString(creditTotalAmount.ToString("N2"), font_Header, Brushes.Black, new RectangleF(50 + 600 - 5, secondTableYPos + 9, 150, tableHeight), sfAlignRight);
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

        }

        private void Layout_ItemReceipt(PrintPageEventArgs e, List<ItemReciept> receiptData, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {
            if (GlobalVariables.includeImage)
            {
                Image image = Properties.Resources.rr_sample;
                e.Graphics.DrawImage(image, e.PageBounds);
            }

            Font font_Details = font_Ten;

            Rectangle rectReceivingPoint = new Rectangle(165, 226 - 10, 359, 15);
            Rectangle rectReceivingAddress = new Rectangle(110, 257 - 10, 610, 15);
            Rectangle rectDate = new Rectangle(616 + 10, 226 - 10, 200, 15);

            //e.Graphics.DrawRectangle(Pens.Black, rectReceivingPoint);
            //e.Graphics.DrawRectangle(Pens.Black, rectReceivingAddress);
            //e.Graphics.DrawRectangle(Pens.Black, rectDate);

            string vendor = receiptData[0].PayeeFullName;
            string fullAddress = receiptData[0].Addr1.ToString() + receiptData[0].Addr2.ToString() + receiptData[0].Addr3.ToString() + receiptData[0].Addr4.ToString() + receiptData[0].AddrCity.ToString();

            //e.Graphics.DrawString(receiptData[0].Memo.ToString(), font_Details, Brushes.Black, rectReceivingPoint);
            //e.Graphics.DrawString(fullAddress, font_Details, Brushes.Black, rectReceivingAddress);
            e.Graphics.DrawString(receiptData[0].DateCreated.ToString("MM/dd/yyyy"), font_Details, Brushes.Black, rectDate);

            // TABLE
            Rectangle rectItemNo = new Rectangle(44, 329 - 12, 52, 24);
            Rectangle rectItemQuantity = new Rectangle(44 + 52, 329 - 12, 92, 24);
            Rectangle rectItemUnit = new Rectangle(44 + 52 + 92, 329 - 12, 140, 24);
            Rectangle rectItemDescription = new Rectangle(44 + 52 + 92 + 140 + 15, 329 - 12, 487, 24);

            //e.Graphics.DrawRectangle(Pens.Black, rectItemNo);
            //e.Graphics.DrawRectangle(Pens.Black, rectItemQuantity);
            //e.Graphics.DrawRectangle(Pens.Black, rectItemUnit);
            //e.Graphics.DrawRectangle(Pens.Black, rectItemDescription);

            int itemHeight = 0;
            int tabDataHeight = 23;
            int counter = 1;

            for (int i = 0; i < receiptData.Count; i++)
            {
                if (receiptData[i] == null)
                {
                    continue;
                }

                if (receiptData[i].ItemQuantity == 0.00)
                {
                    continue;
                }

                var itemUM = receiptData[i].ItemUM ?? string.Empty;
                var itemDescription = receiptData[i].ItemDescription ?? string.Empty;

                e.Graphics.DrawString(counter.ToString(), font_Details, Brushes.Black, new Rectangle(44, 320 + itemHeight, 52, 24 + tabDataHeight), sfAlignCenter);
                e.Graphics.DrawString(receiptData[i].ItemQuantity.ToString("N2"), font_Details, Brushes.Black, new Rectangle(44 + 52, 320 + itemHeight, 92, 24 + tabDataHeight), sfAlignCenter);
                e.Graphics.DrawString(itemUM, font_Details, Brushes.Black, new Rectangle(96 + 92, 320 + itemHeight, 140, 24 + tabDataHeight), sfAlignCenter);
                e.Graphics.DrawString(itemDescription, font_Details, Brushes.Black, new Rectangle(44 + 52 + 92 + 140, 320 + itemHeight, 487, 24 + tabDataHeight), sfAlignLeftCenter);

                itemHeight += tabDataHeight;
                counter++;
            }

            // SIGNATORY || LEFT
            Rectangle rectSupplier = new Rectangle(128 + 5, 534, 270, 18);
            Rectangle rectBroker = new Rectangle(128 + 5, 538 + 18, 270, 18);
            //Rectangle rectAddress = new Rectangle(128, 541 + 36, 270, 30);
            Rectangle rectAddress = new Rectangle(128 + 5, 541 + 40, 270, 30);

            Rectangle rectDeliveryReceiptNo = new Rectangle(154, 619, 244, 18);
            Rectangle rectInvoiceNo = new Rectangle(154, 621 + 18, 244, 18);
            Rectangle rectDateDelivered = new Rectangle(154, 642 + 18, 244, 18);

            //e.Graphics.DrawRectangle(Pens.Black, rectSupplier);
            //e.Graphics.DrawRectangle(Pens.Black, rectBroker);
            //e.Graphics.DrawRectangle(Pens.Black, rectAddress);

            e.Graphics.DrawString(vendor, font_Details, Brushes.Black, rectSupplier);
            e.Graphics.DrawString(fullAddress, font_Details, Brushes.Black, rectAddress);

            //e.Graphics.DrawRectangle(Pens.Black, rectDeliveryReceiptNo);
            //e.Graphics.DrawRectangle(Pens.Black, rectInvoiceNo);
            //e.Graphics.DrawRectangle(Pens.Black, rectDateDelivered);

            // RIGHT
            Rectangle rectReceivedBy = new Rectangle(540, 536, 274, 18);
            Rectangle rectCheckedBy = new Rectangle(540, 560 + 18, 274, 18);

            Rectangle rectDeliveredBy = new Rectangle(540, 640, 274, 18);
            Rectangle rectPlateNo = new Rectangle(540, 645 + 18, 274, 18);

            AccessToDatabase accessToDatabase = new AccessToDatabase();
            var text = accessToDatabase.RetrieveSignatoryRRData();

            e.Graphics.DrawString(text.ReceivedBy, font_Details, Brushes.Black, rectReceivedBy, sfAlignCenter);
            e.Graphics.DrawString(text.CheckedBy, font_Details, Brushes.Black, rectCheckedBy, sfAlignCenter);

            //e.Graphics.DrawRectangle(Pens.Black, rectReceivedBy);
            //e.Graphics.DrawRectangle(Pens.Black, rectCheckedBy);

            //e.Graphics.DrawRectangle(Pens.Black, rectDeliveredBy);
            //e.Graphics.DrawRectangle(Pens.Black, rectPlateNo);
        }
    }
}
