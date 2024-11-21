using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using static VoucherPro.DataClass;
namespace VoucherPro.Clients
{
    public class Layouts_LEADS
    {
        Font font_Six = new Font("Microsoft Sans Serif", 6, FontStyle.Regular);
        Font font_Seven = new Font("Microsoft Sans Serif", 7, FontStyle.Regular);
        Font font_Eight = new Font("Microsoft Sans Serif", 8, FontStyle.Regular);
        Font font_EightBold = new Font("Microsoft Sans Serif", 8, FontStyle.Bold);
        Font font_Nine = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
        Font font_NineBold = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
        Font font_Ten = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
        Font font_TenBold = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
        Font font_Eleven = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
        Font font_ElevenBold = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);

        public void PrintPage_LEADS(object sender, PrintPageEventArgs e, int layoutIndex)
        {
            StringFormat sfAlignRight = new StringFormat { Alignment = StringAlignment.Far | StringAlignment.Far };
            StringFormat sfAlignCenterRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignLeftCenter = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };


            switch (layoutIndex)
            {
                case 4:
                    Layout_ItemReceipt(e, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                default:
                    throw new ArgumentException("Invalid layout index");
            }
        }

        private void Layout_CheckVoucher(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {

        }

        private void Layout_ItemReceipt(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {
            if (GlobalVariables.includeImage)
            {
                Image image = Properties.Resources.rr_sample;
                e.Graphics.DrawImage(image, e.PageBounds);
            }

            Font font_Details = font_Eight;

            Rectangle rectReceivingPoint = new Rectangle(165, 226, 359, 15);
            Rectangle rectReceivingAddress = new Rectangle(110, 257, 410, 15);
            Rectangle rectDate = new Rectangle(616, 226, 200, 15);

            e.Graphics.DrawRectangle(Pens.Black, rectReceivingPoint);
            e.Graphics.DrawRectangle(Pens.Black, rectReceivingAddress);
            e.Graphics.DrawRectangle(Pens.Black, rectDate);

            // TABLE
            Rectangle rectItemNo = new Rectangle(44, 329, 52, 24);
            Rectangle rectItemQuantity = new Rectangle(44 + 52, 329, 92, 24);
            Rectangle rectItemUnit = new Rectangle(44 + 52 + 92, 329, 140, 24);
            Rectangle rectItemDescription = new Rectangle(44 + 52 + 92 + 140, 329, 487, 24);

            e.Graphics.DrawRectangle(Pens.Black, rectItemNo);
            e.Graphics.DrawRectangle(Pens.Black, rectItemQuantity);
            e.Graphics.DrawRectangle(Pens.Black, rectItemUnit);
            e.Graphics.DrawRectangle(Pens.Black, rectItemDescription);

            // SIGNATORY || LEFT
            Rectangle rectSupplier = new Rectangle(128, 534, 270, 18);
            Rectangle rectBroker = new Rectangle(128, 538 + 18, 270, 18);
            Rectangle rectAddress = new Rectangle(128, 541 + 36, 270, 18);

            Rectangle rectDeliveryReceiptNo = new Rectangle(154, 619, 244, 18);
            Rectangle rectInvoiceNo = new Rectangle(154, 621 + 18, 244, 18);
            Rectangle rectDateDelivered = new Rectangle(154, 642 + 18, 244, 18);

            e.Graphics.DrawRectangle(Pens.Black, rectSupplier);
            e.Graphics.DrawRectangle(Pens.Black, rectBroker);
            e.Graphics.DrawRectangle(Pens.Black, rectAddress);

            e.Graphics.DrawRectangle(Pens.Black, rectDeliveryReceiptNo);
            e.Graphics.DrawRectangle(Pens.Black, rectInvoiceNo);
            e.Graphics.DrawRectangle(Pens.Black, rectDateDelivered);

            // RIGHT
            Rectangle rectReceivedBy = new Rectangle(540, 536, 274, 18);
            Rectangle rectCheckedBy = new Rectangle(540, 560 + 18, 274, 18);

            Rectangle rectDeliveredBy = new Rectangle(540, 640, 274, 18);
            Rectangle rectPlateNo = new Rectangle(540, 645 + 18, 274, 18);

            e.Graphics.DrawRectangle(Pens.Black, rectReceivedBy);
            e.Graphics.DrawRectangle(Pens.Black, rectCheckedBy);

            e.Graphics.DrawRectangle(Pens.Black, rectDeliveredBy);
            e.Graphics.DrawRectangle(Pens.Black, rectPlateNo);
        }
    }
}
