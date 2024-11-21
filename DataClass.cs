using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VoucherPro
{
    public class DataClass
    {
        public class ItemReciept
        {
            public string PayeeFullName { get; set; }
            public string Addr1 { get; set; }
            public string Addr2 { get; set; }
            public string Addr3 { get; set; }
            public string Addr4 { get; set; }
            public string AddrCity { get; set; }
            public string RefNumber { get; set; }
            public DateTime DateCreated { get; set; }
            public string Memo { get; set; }
            public double TotalAmount { get; set; }
            public string BankAccount { get; set; }
            public string Account { get; set; }
            public string ExpensesMemo { get; set; }
            public string ItemDescription { get; set; }
            public string ItemClass { get; set; }
            public string ItemCustomerJob { get; set; }
            public string Item { get; set; }
            public string ItemUM { get; set; }
            public double ItemQuantity { get; set; }
            public double ItemCost { get; set; }
            public double ExpensesAmount { get; set; }
            public double ItemAmount { get; set; }
            public ReceiptItemType ReceiptItemType { get; set; }
        }
        public enum ReceiptItemType
        {
            ReceiptItem,
            RecieptExpense
        }
    }
}
