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
                case 1:
                    //Layout_SalesInvoice_ACOM(e, invoiceData, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter);
                    break;
                default:
                    throw new ArgumentException("Invalid layout index");
            }
        }

        private void Layout_CheckVoucher(PrintPageEventArgs e, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter)
        {

        }
    }
}
