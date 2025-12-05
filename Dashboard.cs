using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows.Forms;

using VoucherPro.Clients;
using static VoucherPro.DataClass;

namespace VoucherPro
{
    public class GlobalVariables
    {
        public static string client = "IVP";
        public static bool includeImage = true;
        public static bool includeItemReceipt = true;
        public static bool testWithoutData = false;
        public static bool isPrinting = false;
        public static bool useCrystalReports_LEADS = true;
        public static int itemsPerPageAPV = 10;
    }
    public partial class Dashboard : Form
    {
        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase accessToDatabase;

        ComboBox comboBox_Forms;

        Label label_SeriesNumberText;
        Label label_SignatoryRRStatus;

        TextBox textBox_SeriesNumber;
        TextBox textBox_ReceivedByRR;
        TextBox textBox_CheckedByRR;

        Panel panel_Main;
        Panel panel_Main_CR;

        FlowLayoutPanel panel_Printing;
        FlowLayoutPanel panel_SeriesNumber;
        FlowLayoutPanel panel_Signatory;
        FlowLayoutPanel panel_RRSignatory;
        FlowLayoutPanel panel_RefNumber;
        FlowLayoutPanel panel_RefNumberCrystalReport;

        List<CheckTable> cheque = new List<CheckTable>();
        List<BillTable> bills = new List<BillTable>();
        List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();
        List<ItemReciept> receipts = new List<ItemReciept>();
        List<BillTable> apvData = new List<BillTable>();
        List<CheckTableExpensesAndItems> cvData = new List<CheckTableExpensesAndItems>();

        static int sideBarWidth = 250;
        int seriesNumber = 1;

        //private const int itemsPerPage = 16;
        private int itemCounter;
        private int pageCounter;

        Font font_Label = new Font("Microsoft Sans Serif", 9);
        public Dashboard()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase();

            this.WindowState = FormWindowState.Maximized;
            this.Text = "VoucherPro";

            Bitmap bitmapIcon = Properties.Resources.logo1;
            this.Icon = Icon.FromHandle(bitmapIcon.GetHicon());

            Panel panel_Container = ContainerPanel();
            this.Controls.Add(panel_Container);
        }

        private Panel ContainerPanel()
        {
            Panel panel_Container = new Panel
            {
                Dock = DockStyle.Fill,
            };

            Panel panel_Title = TitlePanel();
            panel_Main = MainPanel();
            panel_Main_CR = MainPanel_CR();
            Panel panel_SideBar = SideBarPanel();

            panel_SideBar.Parent = panel_Container;
            panel_Title.Parent = panel_Container;
            panel_Main.Parent = panel_Container;
            panel_Main_CR.Parent = panel_Container;

            return panel_Container;
        }

        private Panel TitlePanel()
        {
            Panel panel_Title = new Panel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(5),
                Height = 50,
                BackColor = Color.FromArgb(51, 183, 240),
            };

            Label labelTop = new Label
            {
                Parent = panel_Title,
                Font = new Font("Microsoft Sans Serif", 12, FontStyle.Regular),
                Dock = DockStyle.Fill,
                //Text = "QUICKBOOKS SALES INVOICE",
                Text = "V o u c h e r P r o",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
            };

            return panel_Title;
        }

        private Panel MainPanel()
        {
            Panel panel_Main = new Panel
            {
                BackColor = Color.LightGray,
                Dock = DockStyle.Fill,
                Padding = new Padding(sideBarWidth, 50, 0, 0),
                //Height = 300,
            };

            printPreviewControl = new PrintPreviewControl
            {
                Parent = panel_Main,
                Dock = DockStyle.Fill,
                Zoom = 1,
                Visible = false,
            };

            return panel_Main;
        }

        private Panel MainPanel_CR()
        {
            Panel panel_Main_CR = new Panel
            {
                BackColor = Color.LightGray,
                Dock = DockStyle.Fill,
                Padding = new Padding(sideBarWidth, 50, 0, 0),
                //Height = 300,
            };

            reportViewer = new CrystalReportViewer
            {
                Parent = panel_Main_CR,
                Dock = DockStyle.Fill,
                //ReportSource = report2,
                ShowCopyButton = false,
                //ShowPrintButton = false,
                ShowExportButton = false,
                ShowRefreshButton = false,
                ShowGroupTreeButton = false,
                ShowTextSearchButton = false,
                ShowParameterPanelButton = false,
                ToolPanelView = ToolPanelViewType.None
            };

            return panel_Main_CR;
        }

        private Panel SideBarPanel()
        {
            FlowLayoutPanel panel_SideBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = sideBarWidth,
                Padding = new Padding(2),
                //BackColor = Color.Green,
                BackColor = Color.FromArgb(9, 102, 176)
            };

            // - FORMS --------------------------------------------------
            FlowLayoutPanel panels_Forms = Panel_SBForms();
            panels_Forms.Parent = panel_SideBar;

            // - SERIES NUMBER ------------------------------------------
            panel_SeriesNumber = Panel_SBSeriesNumber();
            panel_SeriesNumber.Parent = panel_SideBar;
            panel_SeriesNumber.Visible = false;

            // - REF NUMBER ---------------------------------------------
            panel_RefNumber = Panel_SBRefNumber();
            panel_RefNumberCrystalReport = Panel_SBRefNumber_CR();
            panel_RefNumber.Parent = panel_SideBar;
            panel_RefNumberCrystalReport.Parent = panel_SideBar;
            panel_RefNumber.Visible = false;
            panel_RefNumberCrystalReport.Visible = false;

            // - SIGNATORY ----------------------------------------------
            panel_Signatory = Panel_SBSignatory();
            panel_Signatory.Parent = panel_SideBar;
            panel_Signatory.Visible = false;

            // - RR SIGNATORY -------------------------------------------
            if (GlobalVariables.client == "LEADS")
            {
                panel_RRSignatory = Panel_SBRRSignatory();
                panel_RRSignatory.Parent = panel_SideBar;
                panel_RRSignatory.Visible = false;
            }

            // - PRINTING -----------------------------------------------
            FlowLayoutPanel panel_Printing = Panel_SBPrinting();
            panel_Printing.Parent = panel_SideBar;

            // ----------------------------------------------------------

            return panel_SideBar;
        }

        private FlowLayoutPanel Panel_SBForms()
        {
            FlowLayoutPanel panel_Forms = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 61,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
            };

            Label label_FormText = new Label
            {
                Parent = panel_Forms,
                Width = sideBarWidth - 10,
                Text = "SELECT FORM:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            comboBox_Forms = new ComboBox
            {
                Parent = panel_Forms,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };
            if (GlobalVariables.client == "LEADS")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check",
                "Check Voucher",
                "Accounts Payable Voucher",
                "Item Receipt / Receiving Report",
            });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check Voucher",
            });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }
            else if (GlobalVariables.client == "CPI")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check Voucher",
                "Check",
            });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }
            else if (GlobalVariables.client == "IVP")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check Voucher",
                "Journal Voucher",
            });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }



            return panel_Forms;
        }

        private FlowLayoutPanel Panel_SBSeriesNumber()
        {
            FlowLayoutPanel panel_SeriesNumber = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 62,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false,
            };

            label_SeriesNumberText = new Label
            {
                Parent = panel_SeriesNumber,
                Width = sideBarWidth - 30,
                Text = "Current Series Number:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            textBox_SeriesNumber = new TextBox
            {
                Parent = panel_SeriesNumber,
                Width = 156,
                Font = new Font("Microsoft Sans Serif", 10),
            };
            textBox_SeriesNumber.TextChanged += TextBox_SeriesNumber_TextChanged;
            textBox_SeriesNumber.Leave += TextBox_SeriesNumber_Leave;

            Button button_Decrement = new Button
            {
                Parent = panel_SeriesNumber,
                Height = 28,
                Width = 28,
                Text = "-",
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 1, 0, 0),
                BackColor = Color.Transparent,
            };
            button_Decrement.Click += (sender, e) =>
            {
                if (GlobalVariables.client == "LEADS")
                {
                    if (seriesNumber != 0)
                    {
                        seriesNumber--;
                        UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
                    }
                }
                else if (GlobalVariables.client == "KAYAK")
                {
                    if (seriesNumber != 0)
                    {
                        seriesNumber--;
                        UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                    }
                }
                else if (GlobalVariables.client == "CPI")
                {
                    if (seriesNumber != 0)
                    {
                        seriesNumber--;
                        UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                    }
                }
                else if (GlobalVariables.client == "IVP")
                {
                    if (seriesNumber != 0)
                    {
                        seriesNumber--;
                        // Update this logic to check for index 2
                        string prefix = comboBox_Forms.SelectedIndex == 2 ? "JV" : "CV";
                        UpdateSeriesNumber(prefix);
                    }
                }




            };

            Button button_Increment = new Button
            {
                Parent = panel_SeriesNumber,
                Height = 28,
                Width = 28,
                Text = "+",
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(3, 1, 3, 0),
                BackColor = Color.Transparent,
            };
            button_Increment.Click += (sender, e) =>
            {
                if (GlobalVariables.client == "LEADS")
                {
                    seriesNumber++;
                    UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
                }
                else if (GlobalVariables.client == "KAYAK")
                {
                    seriesNumber++;
                    UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                }
                else if (GlobalVariables.client == "CPI")
                {
                    seriesNumber++;
                    UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                }
                else if (GlobalVariables.client == "IVP")
                {
                    seriesNumber++;
                    // Update this logic to check for index 2
                    string prefix = comboBox_Forms.SelectedIndex == 2 ? "JV" : "CV";
                    UpdateSeriesNumber(prefix);
                }




            };

            return panel_SeriesNumber;
        }

        private FlowLayoutPanel Panel_SBRefNumber_CR()
        {
            FlowLayoutPanel panel_RefNumber_CR = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label label_RefNumberText = new Label
            {
                Parent = panel_RefNumber_CR,
                Width = sideBarWidth - 30,
                Text = "ENTER REFERENCE NUMBER: CR",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            TextBox textBox_ReferenceNumber_CR = new TextBox
            {
                Parent = panel_RefNumber_CR,
                Width = sideBarWidth - 30, // 190
                Font = font_Label,
            };

            Button button_SearchRefNum_CR = new Button
            {
                Parent = panel_RefNumber_CR,
                Height = 26,
                Width = sideBarWidth - 30,
                Text = "SEARCH",
                BackColor = Color.Transparent,
            };
            button_SearchRefNum_CR.Click += (sender, e) =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form.", "Notice", MessageBoxButtons.OK);
                }
                else if (comboBox_Forms.SelectedIndex != 0 && textBox_ReferenceNumber_CR.Text != "")
                {
                    if(GlobalVariables.client == "LEADS")
                    {
                        try
                        {
                            CRAPV_LEADS cRAPV_LEADS = new CRAPV_LEADS();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRAPV_LEADS, databasePath);

                            AccessQueries accessQueries = new AccessQueries();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            apvData = new List<BillTable>();
                            apvData = accessQueries.GetAccountsPayableData_LEADS(refNumberCR);



                            if (apvData.Count > 0)
                            {
                                TextObject textObject_RefNumber = cRAPV_LEADS.ReportDefinition.ReportObjects["TextRefNo"] as TextObject;
                                TextObject textObject_Paid = cRAPV_LEADS.ReportDefinition.ReportObjects["TextPaid"] as TextObject;
                                TextObject textObject_Payee = cRAPV_LEADS.ReportDefinition.ReportObjects["TextPayee"] as TextObject;
                                TextObject textObject_APVSeries = cRAPV_LEADS.ReportDefinition.ReportObjects["TextSeriesNumber"] as TextObject;
                                TextObject textObject_BillDate = cRAPV_LEADS.ReportDefinition.ReportObjects["TextBillDate"] as TextObject;
                                TextObject textObject_DueDate = cRAPV_LEADS.ReportDefinition.ReportObjects["TextDueDate"] as TextObject;
                                TextObject textObject_Terms = cRAPV_LEADS.ReportDefinition.ReportObjects["TextTerms"] as TextObject;
                                TextObject textObject_Amount = cRAPV_LEADS.ReportDefinition.ReportObjects["TextAmount"] as TextObject;
                                TextObject textObject_AmountInWords = cRAPV_LEADS.ReportDefinition.ReportObjects["TextAmountInWords"] as TextObject;
                                TextObject textObject_TotalDebit = cRAPV_LEADS.ReportDefinition.ReportObjects["TextTotalDebit"] as TextObject;
                                TextObject textObject_TotalCredit = cRAPV_LEADS.ReportDefinition.ReportObjects["TextTotalCredit"] as TextObject;

                                TextObject textObject_PreparedBy = cRAPV_LEADS.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_PreparedByPos = cRAPV_LEADS.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRAPV_LEADS.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_CheckedByPos = cRAPV_LEADS.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRAPV_LEADS.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                TextObject textObject_ApprovedByPos = cRAPV_LEADS.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                TextObject textObject_ReceivedBy = cRAPV_LEADS.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                TextObject textObject_ReceivedByPos = cRAPV_LEADS.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                AccessToDatabase accessToDatabase = new AccessToDatabase();

                                var (PreparedByName, PreparedByPosition,
                                    ReviewedByName, ReviewedByPosition,
                                    RecommendingApprovalName, RecommendingApprovalPosition,
                                    ApprovedByName, ApprovedByPosition,
                                    ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                                string refNumber = textBox_ReferenceNumber_CR.Text;
                                double amount = apvData[0].AmountDue;
                                string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                                if (apvData[0].IsPaid)
                                {
                                    textObject_Paid.Text = "PAID";
                                }
                                else
                                {
                                    textObject_Paid.Text = "";
                                }

                                textObject_RefNumber.Text = refNumber;
                                textObject_Payee.Text = apvData[0].Vendor.ToString();
                                textObject_APVSeries.Text = textBox_SeriesNumber.Text;
                                textObject_BillDate.Text = apvData[0].DateCreated.ToString("dd-MMM-yyyy");
                                textObject_DueDate.Text = apvData[0].DueDate.ToString("MM/dd/yyyy");
                                textObject_Terms.Text = apvData[0].TermsRefFullName;
                                textObject_Amount.Text = amount.ToString("N2");
                                textObject_AmountInWords.Text = amountInWords;
                                textObject_PreparedBy.Text = PreparedByName;
                                textObject_PreparedByPos.Text = PreparedByPosition;
                                textObject_CheckedBy.Text = ReviewedByName;
                                textObject_CheckedByPos.Text = ReviewedByPosition;
                                textObject_ApprovedBy.Text = ApprovedByName;
                                textObject_ApprovedByPos.Text = ApprovedByPosition;
                                textObject_ReceivedBy.Text = ReceivedByName;
                                textObject_ReceivedByPos.Text = ReceivedByPosition;

                                double debitTotalAmount = 0;
                                double creditTotalAmount = 0;

                                foreach (var bill in apvData)
                                {
                                    try
                                    {
                                        for (int i = 0; i < bill.AccountNameParticularsList.Count; i++)
                                        {
                                            double itemAmount = bill.ItemDetails[i].ItemLineAmount;

                                            if (itemAmount > 0)
                                            {
                                                debitTotalAmount += itemAmount;
                                            }
                                            else if (itemAmount < 0)
                                            {
                                                creditTotalAmount += Math.Abs(itemAmount);
                                            }
                                        }

                                        foreach (var item in bill.ItemDetails)
                                        {
                                            if (!string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                                            {
                                                double expenseAmount = item.ExpenseLineAmount;

                                                if (expenseAmount > 0)
                                                {
                                                    debitTotalAmount += expenseAmount;
                                                }
                                                else if (expenseAmount < 0)
                                                {
                                                    creditTotalAmount += Math.Abs(expenseAmount);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"An error occurred while computing for total debit and credit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }

                                textObject_TotalDebit.Text = debitTotalAmount.ToString("N2");
                                textObject_TotalCredit.Text = debitTotalAmount.ToString("N2");

                                // Locate the subreport object in the main report
                                SubreportObject subreportObject = cRAPV_LEADS.ReportDefinition.ReportObjects["Subreport2"] as SubreportObject;

                                if (subreportObject != null)
                                {
                                    // Get the ReportDocument of the subreport
                                    ReportDocument subReportDocument = cRAPV_LEADS.OpenSubreport(subreportObject.SubreportName);

                                    // Access the desired TextObject in the subreport

                                    TextObject textObject_Payable = subReportDocument.ReportDefinition.ReportObjects["TextPayable"] as TextObject;
                                    TextObject textObject_PayableAmount = subReportDocument.ReportDefinition.ReportObjects["TextPayableAmount"] as TextObject;
                                    TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;
                                    debitTotalAmount -= creditTotalAmount;

                                    textObject_PayableAmount.Text = debitTotalAmount.ToString("N2");
                                    textObject_Payable.Text = apvData[0].APAccountRefFullName.ToString();
                                    textObject_Remarks.Text = "Remarks: " + apvData[0].Memo.ToString();

                                    DataTable dataTable = new DataTable();
                                    dataTable.Columns.Add("Particulars", typeof(string)); // First column
                                    dataTable.Columns.Add("Memo", typeof(string)); // First column
                                    dataTable.Columns.Add("Class", typeof(string)); // Second column
                                    dataTable.Columns.Add("Debit", typeof(string)); // Third column
                                    dataTable.Columns.Add("Credit", typeof(string)); // Fourth column

                                    InsertDataToCVCompiled(refNumber, apvData);
                                }


                                cRAPV_LEADS.SetParameterValue("ReferenceNumber", refNumber);

                                reportViewer.ReportSource = cRAPV_LEADS;
                                reportViewer.RefreshReport();
                            }
                            else
                            {
                                MessageBox.Show("No data found for the provided reference number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"An error occurred while loading the report:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else if (GlobalVariables.client == "CPI")
                    {
                        try
                        {
                            CRCV_KAYAK cRCV_Kayak = new CRCV_KAYAK();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRCV_Kayak, databasePath);

                            AccessQueries accessQueries = new AccessQueries();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            cvData = new List<CheckTableExpensesAndItems>();
                            cvData = accessQueries.GetCheckExpensesAndItemsData_KAYAK(refNumberCR);


                            if (cvData.Count > 0)
                            {
                                TextObject textObject_CVCheckNumber = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVCheckNumber"] as TextObject;
                                TextObject textObject_CVRefNumber = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                TextObject textObject_CVAmountInWords = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject;
                                TextObject textObject_CVBank = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVBank"] as TextObject;
                                TextObject textObject_CVCheckDate = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                TextObject textObject_CVPayee = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                TextObject textObject_CVTotalAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalAmount"] as TextObject;
                                TextObject textObject_CVTotalDebitAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalDebitAmount"] as TextObject;
                                TextObject textObject_CVTotalCreditAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalCreditAmount"] as TextObject;
                                TextObject textObject_Paid = cRCV_Kayak.ReportDefinition.ReportObjects["TextPaid"] as TextObject;


                                TextObject textObject_PreparedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_PreparedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_CheckedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                TextObject textObject_ApprovedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                TextObject textObject_ReceivedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                TextObject textObject_ReceivedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                AccessToDatabase accessToDatabase = new AccessToDatabase();

                                var (PreparedByName, PreparedByPosition,
                                   ReviewedByName, ReviewedByPosition,
                                   RecommendingApprovalName, RecommendingApprovalPosition,
                                   ApprovedByName, ApprovedByPosition,
                                   ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                                double amount = cvData[0].TotalAmount;
                                string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                                textObject_Paid.Text = "";

                                textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                textObject_CVAmountInWords.Text = amountInWords;
                                textObject_CVBank.Text = cvData[0].BankAccount;
                                textObject_CVCheckDate.Text = cvData[0].DateCreated.ToString("dd-MMM-yyyy");
                                textObject_CVPayee.Text = cvData[0].PayeeFullName.ToString();
                                textObject_CVTotalAmount.Text = cvData[0].TotalAmount.ToString("N2");


                                string refNumber = textBox_ReferenceNumber_CR.Text;
                                textObject_CVCheckNumber.Text = refNumber;

                                textObject_PreparedBy.Text = PreparedByName;
                                textObject_PreparedByPos.Text = PreparedByPosition;
                                textObject_CheckedBy.Text = ReviewedByName;
                                textObject_CheckedByPos.Text = ReviewedByPosition;
                                textObject_ApprovedBy.Text = ApprovedByName;
                                textObject_ApprovedByPos.Text = ApprovedByPosition;
                                textObject_ReceivedBy.Text = ReceivedByName;
                                textObject_ReceivedByPos.Text = ReceivedByPosition;

                                double debitTotalAmount = 0;
                                double creditTotalAmount = 0;

                                foreach (var data in cvData)
                                {
                                    try
                                    {
                                        // Handling Item Amounts
                                        double itemAmount = data.ItemAmount;
                                        if (itemAmount > 0)
                                        {
                                            debitTotalAmount += itemAmount;
                                        }
                                        else if (itemAmount < 0)
                                        {
                                            creditTotalAmount += Math.Abs(itemAmount);
                                        }

                                        // Handling Expenses Amounts
                                        if (!string.IsNullOrEmpty(data.AccountNameCheck))
                                        {
                                            double expenseAmount = data.ExpensesAmount;
                                            if (expenseAmount > 0)
                                            {
                                                debitTotalAmount += expenseAmount;
                                            }
                                            else if (expenseAmount < 0)
                                            {
                                                creditTotalAmount += Math.Abs(expenseAmount);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"An error occurred while computing for total debit and credit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }


                                textObject_CVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                textObject_CVTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                                // Locate the subreport object in the main report
                                SubreportObject subreportObject = cRCV_Kayak.ReportDefinition.ReportObjects["SubreportCVDetails"] as SubreportObject;

                                if (subreportObject != null)
                                {
                                    // Get the ReportDocument of the subreport
                                    ReportDocument subReportDocument = cRCV_Kayak.OpenSubreport(subreportObject.SubreportName);

                                    // Access the desired TextObject in the subreport

                                    TextObject textObject_AccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextAccountPayable"] as TextObject;
                                    TextObject textObject_TextAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextAmountPayable"] as TextObject;
                                    TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;

                                    textObject_AccountPayable.Text = cvData[0].BankAccountNumber + " - " + cvData[0].BankAccount.ToString();
                                    textObject_TextAmountPayable.Text = debitTotalAmount.ToString("N2");
                                    textObject_Remarks.Text = cvData[0].Memo.ToString();
                                    // Create a DataTable with 4 columns

                                    DataTable dataTable = new DataTable();
                                    dataTable.Columns.Add("Particulars", typeof(string)); // First column
                                    dataTable.Columns.Add("Class", typeof(string)); // Second column
                                    dataTable.Columns.Add("Debit", typeof(string)); // Third column
                                    dataTable.Columns.Add("Credit", typeof(string)); // Fourth column

                                    InsertDataToCheckVoucherCompiled(refNumber, cvData);
                                }

                                cRCV_Kayak.SetParameterValue("ReferenceNumber", refNumber);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRCV_Kayak;
                                reportViewer.RefreshReport();

                            }
                            else
                            {
                                //SearchBillsByReference(refNumberCR);
                                try
                                {
                                    CRCV_CPIBILL cRCVCPIBILL = new CRCV_CPIBILL();
                                    string databasePath2 = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                    SetDatabaseLocation(cRCVCPIBILL, databasePath2);

                                    AccessQueries accessQueries2 = new AccessQueries();
                                    string refNumberCR2 = textBox_ReferenceNumber_CR.Text;

                                    bills = new List<BillTable>();
                                    bills = accessQueries.GetBillData_CPI(refNumberCR);

                                    if (bills.Count > 0)
                                    {
                                        TextObject textObject_CVCheckNumber = cRCVCPIBILL.ReportDefinition.ReportObjects["TextCVSeriesnumber"] as TextObject;
                                        TextObject textObject_CVRefNumber = cRCVCPIBILL.ReportDefinition.ReportObjects["TextRefNumber"] as TextObject;
                                        TextObject textObject_CVAmountInWords = cRCVCPIBILL.ReportDefinition.ReportObjects["TextAmountInWords"] as TextObject;
                                        TextObject textObject_CVBank = cRCVCPIBILL.ReportDefinition.ReportObjects["TextBankAccount"] as TextObject;
                                        TextObject textObject_CVCheckDate = cRCVCPIBILL.ReportDefinition.ReportObjects["TextCheckDate"] as TextObject;
                                        TextObject textObject_CVPayee = cRCVCPIBILL.ReportDefinition.ReportObjects["TextPayeeAccount"] as TextObject;
                                        TextObject textObject_CVTotalAmount = cRCVCPIBILL.ReportDefinition.ReportObjects["TextTotalAmount"] as TextObject;
                                        TextObject textObject_CVTotalDebitAmount = cRCVCPIBILL.ReportDefinition.ReportObjects["TextDebitTotalAmount"] as TextObject;
                                        TextObject textObject_CVTotalCreditAmount = cRCVCPIBILL.ReportDefinition.ReportObjects["TextCreditTotalAmount"] as TextObject;
                                        //TextObject textObject_Paid = cRCVCPIBILL.ReportDefinition.ReportObjects["TextPaid"] as TextObject;


                                        TextObject textObject_PreparedBy = cRCVCPIBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                        TextObject textObject_PreparedByPos = cRCVCPIBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                        TextObject textObject_CheckedBy = cRCVCPIBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                        TextObject textObject_CheckedByPos = cRCVCPIBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                        TextObject textObject_ApprovedBy = cRCVCPIBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                        TextObject textObject_ApprovedByPos = cRCVCPIBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                        TextObject textObject_ReceivedBy = cRCVCPIBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                        TextObject textObject_ReceivedByPos = cRCVCPIBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                        AccessToDatabase accessToDatabase = new AccessToDatabase();

                                        var (PreparedByName, PreparedByPosition,
                                           ReviewedByName, ReviewedByPosition,
                                           RecommendingApprovalName, RecommendingApprovalPosition,
                                           ApprovedByName, ApprovedByPosition,
                                           ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                                        double amount = bills[0].AmountDue;
                                        string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                                        //textObject_Paid.Text = "";

                                        textObject_CVCheckNumber.Text = textBox_SeriesNumber.Text;
                                        textObject_CVAmountInWords.Text = amountInWords;
                                        textObject_CVBank.Text = bills[0].BankAccount;
                                        textObject_CVCheckDate.Text = bills[0].DateCreated.ToString("dd-MMM-yyyy");
                                        textObject_CVPayee.Text = bills[0].PayeeFullName.ToString();
                                        textObject_CVTotalAmount.Text = bills[0].AmountDue.ToString("N2");


                                        string refNumber = textBox_ReferenceNumber_CR.Text;
                                        textObject_CVRefNumber.Text = refNumber;

                                        textObject_PreparedBy.Text = PreparedByName;
                                        textObject_PreparedByPos.Text = PreparedByPosition;
                                        textObject_CheckedBy.Text = ReviewedByName;
                                        textObject_CheckedByPos.Text = ReviewedByPosition;
                                        textObject_ApprovedBy.Text = ApprovedByName;
                                        textObject_ApprovedByPos.Text = ApprovedByPosition;
                                        textObject_ReceivedBy.Text = ReceivedByName;
                                        textObject_ReceivedByPos.Text = ReceivedByPosition;

                                        double debitTotalAmount = 0;
                                        double creditTotalAmount = 0;

                                        foreach (var bill in bills) // 'bills' is List<BillTable>
                                        {
                                            foreach (var item in bill.ItemDetails)
                                            {
                                                try
                                                {
                                                    // Handle ItemLineAmount
                                                    if (item.ItemLineAmount != 0)
                                                    {
                                                        if (item.ItemLineAmount > 0)
                                                            debitTotalAmount += item.ItemLineAmount;
                                                        else
                                                            creditTotalAmount += Math.Abs(item.ItemLineAmount);
                                                    }

                                                    // Handle ExpenseLineAmount
                                                    if (item.ExpenseLineAmount != 0)
                                                    {
                                                        if (item.ExpenseLineAmount > 0)
                                                            debitTotalAmount += item.ExpenseLineAmount;
                                                        else
                                                            creditTotalAmount += Math.Abs(item.ExpenseLineAmount);
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show($"Error processing item detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                }
                                            }
                                        }

                                        textObject_CVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                        textObject_CVTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                                        // Locate the subreport object in the main report
                                        SubreportObject subreportObject = cRCVCPIBILL.ReportDefinition.ReportObjects["SubreportBill1"] as SubreportObject;

                                        if (subreportObject != null)
                                        {
                                            // Get the ReportDocument of the subreport
                                            ReportDocument subReportDocument = cRCVCPIBILL.OpenSubreport(subreportObject.SubreportName);

                                            // Access the desired TextObject in the subreport

                                            TextObject textObject_AccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextPayable"] as TextObject;
                                            TextObject textObject_TextAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextPayableAmount"] as TextObject;
                                            TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;

                                            textObject_AccountPayable.Text = bills[0].BankAccount.ToString();
                                            textObject_TextAmountPayable.Text = (debitTotalAmount - creditTotalAmount).ToString("N2");
                                            textObject_Remarks.Text = bills[0].Memo.ToString();
                                            // Create a DataTable with 4 columns

                                            DataTable dataTable = new DataTable();
                                            dataTable.Columns.Add("Particulars", typeof(string)); // First column
                                            dataTable.Columns.Add("Class", typeof(string)); // Second column
                                            dataTable.Columns.Add("Debit", typeof(string)); // Third column
                                            dataTable.Columns.Add("Credit", typeof(string)); // Fourth column

                                            InsertDataToBillCompiled(refNumber, bills);
                                        }

                                        cRCV_Kayak.SetParameterValue("ReferenceNumber", refNumber);

                                        panel_Printing.Visible = false;
                                        panel_Signatory.Visible = true;
                                        panel_Main.Visible = false;
                                        panel_Main_CR.Visible = true;

                                        reportViewer.ReportSource = cRCVCPIBILL;
                                        reportViewer.RefreshReport();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    else if (GlobalVariables.client == "IVP")
                    {
                        // -------------------------------------------------------------
                        // OPTION 1: CHECK VOUCHER
                        // -------------------------------------------------------------
                        if (comboBox_Forms.SelectedIndex == 1)
                        {
                            bool cvDataExists = false;
                            try
                            {
                                CRCV_IVP cRCV_IVP = new CRCV_IVP();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRCV_IVP, databasePath);

                                AccessQueries accessQueries = new AccessQueries();
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;

                                cvData = accessQueries.GetCheckExpensesAndItemsData_IVP(refNumberCR);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    TextObject textObject_CVRefNumber = cRCV_IVP.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                    TextObject textObject_CVAmountInWords = cRCV_IVP.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject;
                                    TextObject textObject_CVCheckDate = cRCV_IVP.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                    TextObject textObject_CVPayee = cRCV_IVP.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                    TextObject textObject_CVTotalAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalAmount"] as TextObject;
                                    TextObject textObject_CVTotalDebitAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalDebitAmount"] as TextObject;
                                    TextObject textObject_CVTotalCreditAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalCreditAmount"] as TextObject;

                                    TextObject textObject_PreparedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                    TextObject textObject_PreparedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                    TextObject textObject_CheckedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                    TextObject textObject_CheckedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                    TextObject textObject_ApprovedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                    TextObject textObject_ApprovedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                    TextObject textObject_ReceivedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                    TextObject textObject_ReceivedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                    AccessToDatabase accessToDatabase = new AccessToDatabase();
                                    var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                                    textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                    textObject_CVAmountInWords.Text = amountInWords;
                                    textObject_CVCheckDate.Text = cvData[0].DateCreated.ToString("MMMM dd, yyyy");
                                    textObject_CVPayee.Text = cvData[0].PayeeFullName;
                                    textObject_CVTotalAmount.Text = cvData[0].TotalAmount.ToString("N2");

                                    textObject_PreparedBy.Text = signatories.PreparedByName;
                                    textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                                    textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                    textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                    textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;
                                    textObject_ReceivedBy.Text = signatories.ReceivedByName;
                                    textObject_ReceivedByPos.Text = signatories.ReceivedByPosition;

                                    double debitTotalAmount = 0;
                                    double creditTotalAmount = 0;

                                    foreach (var data in cvData)
                                    {
                                        try
                                        {
                                            double itemAmount = data.ItemAmount;
                                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                                            if (!string.IsNullOrEmpty(data.Account))
                                            {
                                                double expenseAmount = data.ExpensesAmount;
                                                if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                                                else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);
                                            }
                                        }
                                        catch (Exception ex) { MessageBox.Show($"Error computing totals: {ex.Message}"); }
                                    }

                                    textObject_CVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                    textObject_CVTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                                    SubreportObject subreportObject = cRCV_IVP.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                                    if (subreportObject != null)
                                    {
                                        ReportDocument subReportDocument = cRCV_IVP.OpenSubreport(subreportObject.SubreportName);
                                        TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;
                                        TextObject textObject_CVSubTotal = subReportDocument.ReportDefinition.ReportObjects["TextCVSubTotalAmount"] as TextObject;
                                        TextObject textObject_CVSubCheckNumber = subReportDocument.ReportDefinition.ReportObjects["TextCVSubCheckNumber"] as TextObject;
                                        TextObject textObject_CVSubCheckDate = subReportDocument.ReportDefinition.ReportObjects["TextCVSubCheckDate"] as TextObject;
                                        TextObject textObject_SubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAccountPayable"] as TextObject;
                                        TextObject textObject_SubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAmountPayable"] as TextObject;

                                        textObject_Remarks.Text = cvData[0].Memo;
                                        textObject_CVSubTotal.Text = cvData[0].TotalAmount.ToString("N2");
                                        textObject_CVSubCheckNumber.Text = cvData[0].RefNumber;
                                        textObject_CVSubCheckDate.Text = cvData[0].DateCreated.ToString("MMMM dd, yyyy");
                                        textObject_SubAccountPayable.Text = cvData[0].BankAccount;
                                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");

                                        InsertDataToCheckVoucherCompiledIVP(refNumberCR, cvData);
                                    }

                                    cRCV_IVP.SetParameterValue("ReferenceNumber", refNumberCR);

                                    panel_Printing.Visible = false;
                                    panel_Signatory.Visible = true;
                                    panel_Main.Visible = false;
                                    panel_Main_CR.Visible = true;

                                    reportViewer.ReportSource = cRCV_IVP;
                                    reportViewer.RefreshReport();
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"IVP CV ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            if (!cvDataExists)
                            {
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;
                                GenerateBillPaymentReport_IVP(refNumberCR);
                            }
                        }

                        // -------------------------------------------------------------
                        // OPTION 2: JOURNAL VOUCHER
                        // -------------------------------------------------------------
                        else if (comboBox_Forms.SelectedIndex == 2)
                        {

                        }
                    }





                    else if (GlobalVariables.client == "KAYAK")
                    {
                        try
                        {
                            CRCV_KAYAK cRCV_Kayak = new CRCV_KAYAK();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRCV_Kayak, databasePath);

                            AccessQueries accessQueries = new AccessQueries();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            cvData = new List<CheckTableExpensesAndItems>();
                            cvData = accessQueries.GetCheckExpensesAndItemsData_KAYAK(refNumberCR);


                            if (cvData.Count > 0)
                            {
                                TextObject textObject_CVCheckNumber = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVCheckNumber"] as TextObject;
                                TextObject textObject_CVRefNumber = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                TextObject textObject_CVAmountInWords = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject;
                                TextObject textObject_CVBank = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVBank"] as TextObject;
                                TextObject textObject_CVCheckDate = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                TextObject textObject_CVPayee = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                TextObject textObject_CVTotalAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalAmount"] as TextObject;
                                TextObject textObject_CVTotalDebitAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalDebitAmount"] as TextObject;
                                TextObject textObject_CVTotalCreditAmount = cRCV_Kayak.ReportDefinition.ReportObjects["TextCVTotalCreditAmount"] as TextObject;
                                TextObject textObject_Paid = cRCV_Kayak.ReportDefinition.ReportObjects["TextPaid"] as TextObject;


                                TextObject textObject_PreparedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_PreparedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_CheckedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                TextObject textObject_ApprovedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                TextObject textObject_ReceivedBy = cRCV_Kayak.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                TextObject textObject_ReceivedByPos = cRCV_Kayak.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                AccessToDatabase accessToDatabase = new AccessToDatabase();

                                var (PreparedByName, PreparedByPosition,
                                   ReviewedByName, ReviewedByPosition,
                                   RecommendingApprovalName, RecommendingApprovalPosition,
                                   ApprovedByName, ApprovedByPosition,
                                   ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                                double amount = cvData[0].TotalAmount;
                                string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                                textObject_Paid.Text = "";

                                textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                textObject_CVAmountInWords.Text = amountInWords;
                                textObject_CVBank.Text = cvData[0].BankAccount;
                                textObject_CVCheckDate.Text = cvData[0].DateCreated.ToString("dd-MMM-yyyy");
                                textObject_CVPayee.Text = cvData[0].PayeeFullName.ToString();
                                textObject_CVTotalAmount.Text = cvData[0].TotalAmount.ToString("N2");


                                string refNumber = textBox_ReferenceNumber_CR.Text;
                                textObject_CVCheckNumber.Text = refNumber;

                                textObject_PreparedBy.Text = PreparedByName;
                                textObject_PreparedByPos.Text = PreparedByPosition;
                                textObject_CheckedBy.Text = ReviewedByName;
                                textObject_CheckedByPos.Text = ReviewedByPosition;
                                textObject_ApprovedBy.Text = ApprovedByName;
                                textObject_ApprovedByPos.Text = ApprovedByPosition;
                                textObject_ReceivedBy.Text = ReceivedByName;
                                textObject_ReceivedByPos.Text = ReceivedByPosition;

                                double debitTotalAmount = 0;
                                double creditTotalAmount = 0;

                                foreach (var data in cvData)
                                {
                                    try
                                    {
                                        // Handling Item Amounts
                                        double itemAmount = data.ItemAmount;
                                        if (itemAmount > 0)
                                        {
                                            debitTotalAmount += itemAmount;
                                        }
                                        else if (itemAmount < 0)
                                        {
                                            creditTotalAmount += Math.Abs(itemAmount);
                                        }

                                        // Handling Expenses Amounts
                                        if (!string.IsNullOrEmpty(data.AccountNameCheck))
                                        {
                                            double expenseAmount = data.ExpensesAmount;
                                            if (expenseAmount > 0)
                                            {
                                                debitTotalAmount += expenseAmount;
                                            }
                                            else if (expenseAmount < 0)
                                            {
                                                creditTotalAmount += Math.Abs(expenseAmount);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"An error occurred while computing for total debit and credit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }


                                textObject_CVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                textObject_CVTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                                // Locate the subreport object in the main report
                                SubreportObject subreportObject = cRCV_Kayak.ReportDefinition.ReportObjects["SubreportCVDetails"] as SubreportObject;

                                if (subreportObject != null)
                                {
                                    // Get the ReportDocument of the subreport
                                    ReportDocument subReportDocument = cRCV_Kayak.OpenSubreport(subreportObject.SubreportName);

                                    // Access the desired TextObject in the subreport

                                    TextObject textObject_AccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextAccountPayable"] as TextObject;
                                    TextObject textObject_TextAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextAmountPayable"] as TextObject;
                                    TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;

                                    textObject_AccountPayable.Text = cvData[0].BankAccountNumber + " - " + cvData[0].BankAccount.ToString();
                                    textObject_TextAmountPayable.Text = debitTotalAmount.ToString("N2");
                                    textObject_Remarks.Text = cvData[0].Memo.ToString();
                                    // Create a DataTable with 4 columns

                                    DataTable dataTable = new DataTable();
                                    dataTable.Columns.Add("Particulars", typeof(string)); // First column
                                    dataTable.Columns.Add("Class", typeof(string)); // Second column
                                    dataTable.Columns.Add("Debit", typeof(string)); // Third column
                                    dataTable.Columns.Add("Credit", typeof(string)); // Fourth column

                                    InsertDataToCheckVoucherCompiled(refNumber, cvData);
                                }

                                cRCV_Kayak.SetParameterValue("ReferenceNumber", refNumber);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRCV_Kayak;
                                reportViewer.RefreshReport();

                            }
                            else
                            {
                                SearchBillsByReference(refNumberCR);
                            }
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    
                }
                else
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK);
                }
            };

            return panel_RefNumber_CR;
        }


        private bool GenerateBillPaymentReport_IVP(string refNumberCR)
        {
            try
            {
                CRCV_IVPBILL cRCV_IVPBILL = new CRCV_IVPBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_IVPBILL, databasePathBILL);

                AccessQueries accessQueries = new AccessQueries();
                List<BillTable> bills = accessQueries.GetBillData_IVP(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLCheckNumber = null;
                TextObject textObject_CVBILLAmountInWords = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLTotalAmount = null;
                TextObject textObject_CVBILLTotalDebitAmount = null;
                TextObject textObject_CVBILLTotalCreditAmount = null;
                TextObject textObject_PreparedBy = null;
                TextObject textObject_PreparedByPos = null;
                TextObject textObject_CheckedBy = null;
                TextObject textObject_CheckedByPos = null;
                TextObject textObject_ApprovedBy = null;
                TextObject textObject_ApprovedByPos = null;
                TextObject textObject_ReceivedBy = null;
                TextObject textObject_ReceivedByPos = null;

                try
                {
                    textObject_CVBILLCheckNumber = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLSeriesnumber"] as TextObject;
                    textObject_CVBILLAmountInWords = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLAmountInWords"] as TextObject;
                    textObject_CVBILLCheckDate = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLPayee = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    textObject_CVBILLTotalAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalAmount"] as TextObject;
                    textObject_CVBILLTotalDebitAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalDebitAmount"] as TextObject;
                    textObject_CVBILLTotalCreditAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalCreditAmount"] as TextObject;


                    textObject_PreparedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase accessToDatabase = new AccessToDatabase();

                    var (PreparedByName, PreparedByPosition,
                       ReviewedByName, ReviewedByPosition,
                       RecommendingApprovalName, RecommendingApprovalPosition,
                       ApprovedByName, ApprovedByPosition,
                       ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();


                    double debitTotalAmount = 0;
                    double creditTotalAmount = 0;

                    textObject_PreparedBy.Text = PreparedByName;
                    textObject_PreparedByPos.Text = PreparedByPosition;
                    textObject_CheckedBy.Text = ReviewedByName;
                    textObject_CheckedByPos.Text = ReviewedByPosition;
                    textObject_ApprovedBy.Text = ApprovedByName;
                    textObject_ApprovedByPos.Text = ApprovedByPosition;
                    textObject_ReceivedBy.Text = ReceivedByName;
                    textObject_ReceivedByPos.Text = ReceivedByPosition;

                    foreach (var bill in bills) // 'bills' is List<BillTable>
                    {
                        foreach (var item in bill.ItemDetails)
                        {
                            try
                            {
                                // Handle ItemLineAmount
                                if (item.ItemLineAmount != 0)
                                {
                                    if (item.ItemLineAmount > 0)
                                        debitTotalAmount += item.ItemLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ItemLineAmount);
                                }

                                // Handle ExpenseLineAmount
                                if (item.ExpenseLineAmount != 0)
                                {
                                    if (item.ExpenseLineAmount > 0)
                                        debitTotalAmount += item.ExpenseLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ExpenseLineAmount);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error processing item detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }

                    textObject_CVBILLTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                    textObject_CVBILLTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                }
                catch
                {
                    throw;
                }


                double amount = bills[0].AmountDue;
                string amountInWords = AccessToDatabase.AmountToWordsConverter.Convert(amount);

                if (textObject_CVBILLCheckNumber != null) textObject_CVBILLCheckNumber.Text = textBox_SeriesNumber.Text;
                if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = bills[0].DateCreated.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = bills[0].AmountDue.ToString("N2");

                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRCV_IVPBILL.ReportDefinition.ReportObjects["SubreportCVBILLDetailsIVP"] as SubreportObject;
                }
                catch
                {
                    throw;
                }

                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = null;
                    try
                    {
                        subReportDocument = cRCV_IVPBILL.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                        TextObject textObject_BILLCVSubCheckDate = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckDate"] as TextObject;
                        TextObject textObject_BILLCVSubTotal = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSUBTotalAmount"] as TextObject;
                        TextObject textObject_BILLCVSubCheckNumber = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckNumber"] as TextObject;
                        TextObject textObject_BILLSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountPayable"] as TextObject;
                        TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;


                        if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].BillMemo ?? "";
                        if (textObject_BILLCVSubCheckDate != null) textObject_BILLCVSubCheckDate.Text = bills[0].DateCreated.ToString("MMMM dd, yyyy");
                        if (textObject_BILLCVSubTotal != null) textObject_BILLCVSubTotal.Text = bills[0].AmountDue.ToString("N2");
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLCVSubCheckNumber.Text = bills[0].RefNumber ?? "";
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLSubAccountPayable.Text = bills[0].BankAccount ?? "";
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLSubAmountPayable.Text = bills[0].AmountDue.ToString("N2") ?? "";

                        InsertDataToBillCompiled(refNumberCR, bills);
                    }
                    catch
                    {
                        throw;
                    }
                }

                cRCV_IVPBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_IVPBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }




        public static void InsertDataToCheckVoucherCompiledIVP(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // 1. Clear old data
                string deleteQuery = "DELETE FROM CheckVoucherCompiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from CheckVoucherCompiled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 2. Prepare Insert Query
                string insertQuery = @"
                        INSERT INTO CheckVoucherCompiled 
                        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                foreach (var check in checkData)
                {
                    try
                    {
                        // COMMON FIELDS
                        string memoValue = string.IsNullOrEmpty(check.ExpensesMemo) ? "" : check.ExpensesMemo;
                        string customerJob = string.IsNullOrEmpty(check.ExpensesCustomerJob) ? "" : check.ExpensesCustomerJob;

                        // ---------------------------------------------------------
                        // FIXED SECTION: INSERT ITEM ENTRY
                        // Changed 'check.ItemName' to 'check.Item'
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Item))
                        {
                            string itemName = check.Item; // FIXED: use check.Item
                            string itemClass = check.ItemClass;
                            double itemAmount = check.ItemAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", itemName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }

                        // ---------------------------------------------------------
                        // FIXED SECTION: INSERT EXPENSE ENTRY
                        // Changed 'check.AccountNameCheck' to 'check.Account'
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Account))
                        {
                            // Note: check.AccountNumber might also be empty if you didn't assign it in the retrieval function.
                            // If check.AccountNumber is null, this line might look like " - Utilities". 
                            // You might want to just use check.Account if you don't have numbers.
                            string expenseName = check.Account;

                            // If you have account numbers populated, use this format instead:
                            // string expenseName = check.AccountNumber + " - " + check.Account;

                            string expenseClass = check.ExpenseClass; // Ensure this property name matches too (ExpenseClass vs AccountClassCheck)
                            double expenseAmount = check.ExpensesAmount;

                            string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                            string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                            if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                            else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", expenseName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing check data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }
        public static void InsertDataToCVCompiled(string refNumber, List<BillTable> billData)
        {
            string connectionString = AccessToDatabase.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Delete all existing data in CV_compiled table
                string deleteQuery = "DELETE FROM CV_compiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from CV_compiled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred while deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit if delete fails
                    }
                }

                string insertQuery = "INSERT INTO CV_compiled (RefNumber, Particulars, [Memo], [Class], Debit, Credit) VALUES (@RefNumber, @Particulars, @Memo, @Class, @Debit, @Credit)";

                // Process bills and insert data directly
                foreach (var bill in billData)
                {
                    try
                    {
                        // Process item details
                        for (int i = 0; i < bill.AccountNameParticularsList.Count; i++)
                        {
                            string itemName = bill.ItemDetails[i].ItemLineItemRefFullName;
                            string itemClass = bill.ItemDetails[i].ItemLineClassRefFullName;
                            string itemMemo = bill.ItemDetails[i].ItemLineMemo;
                            double itemAmount = bill.ItemDetails[i].ItemLineAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0)
                            {
                                debitTotalAmount += itemAmount;
                            }
                            else if (itemAmount < 0)
                            {
                                creditTotalAmount += Math.Abs(itemAmount);
                            }

                            //string insertQuery = "INSERT INTO CV_compiled (Particulars, Class, Debit, Credit) VALUES (@Particulars, @Class, @Debit, @Credit)";
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", itemName);
                                command.Parameters.AddWithValue("@Memo", string.IsNullOrEmpty(itemMemo) ? (object)DBNull.Value : itemMemo);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);

                                command.ExecuteNonQuery();
                            }

                            Console.WriteLine($"Inserted Item: {itemName}, Memo: {itemMemo}, Debit: {debit}, Credit: {credit}");
                        }

                        // Process expense details
                        foreach (var item in bill.ItemDetails)
                        {
                            if (!string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                            {
                                string expenseName = item.ExpenseLineItemRefFullName;
                                string expenseClass = item.ExpenseLineClassRefFullName;
                                string expenseMemo = item.ExpenseLineMemo;
                                double expenseAmount = item.ExpenseLineAmount;

                                string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                                string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                                if (expenseAmount > 0)
                                {
                                    debitTotalAmount += expenseAmount;
                                }
                                else if (expenseAmount < 0)
                                {
                                    creditTotalAmount += Math.Abs(expenseAmount);
                                }

                                //string insertQuery = "INSERT INTO CV_compiled (Particulars, Class, Debit, Credit) VALUES (@Particulars, @Class, @Debit, @Credit)";
                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber);
                                    command.Parameters.AddWithValue("@Particulars", expenseName);
                                    command.Parameters.AddWithValue("@Memo", string.IsNullOrEmpty(expenseMemo) ? (object)DBNull.Value : expenseMemo);
                                    command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                    command.Parameters.AddWithValue("@Debit", debit);
                                    command.Parameters.AddWithValue("@Credit", credit);

                                    command.ExecuteNonQuery();
                                }

                                Console.WriteLine($"Inserted Expense: {expenseName},  Memo: {expenseMemo}, Debit: {debit}, Credit: {credit}");
                            }
                        }
                        debitTotalAmount -= creditTotalAmount;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred while processing bill data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Close the connection
                connection.Close();
            }
            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
            //MessageBox.Show("Data has been inserted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /*public static void InsertDataToCheckVoucherCompiled(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Clear old data
                string deleteQuery = "DELETE FROM CheckVoucherCompiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from CheckVoucherCompiled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                string insertQuery = "INSERT INTO CheckVoucherCompiled (RefNumber, Particulars, [Class], Debit, Credit) VALUES (@RefNumber, @Particulars, @Class, @Debit, @Credit)";

                foreach (var check in checkData)
                {
                    try
                    {
                       
                        // Insert expense details

                        // KULANG HIN ACCOUNT NUMBER
                        if (!string.IsNullOrEmpty(check.ItemName))
                        {
                            string itemName = check.ItemName;
                            string itemClass = check.ItemClass;
                            double itemAmount = check.ItemAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0)
                            {
                                debitTotalAmount += itemAmount;
                            }
                            else if (itemAmount < 0)
                            {
                                creditTotalAmount += Math.Abs(itemAmount);
                            }

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", itemName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.ExecuteNonQuery();
                            }

                            Console.WriteLine($"Inserted Expense: {itemName}, Debit: {debit}, Credit: {credit}");
                        }

                        if (!string.IsNullOrEmpty(check.AccountNameCheck))
                        {
                            string expenseName = check.AccountNumber +" - "+ check.AccountNameCheck;
                            string expenseClass = check.AccountClassCheck;
                            double expenseAmount = check.ExpensesAmount;

                            string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                            string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                            if (expenseAmount > 0)
                            {
                                debitTotalAmount += expenseAmount;
                            }
                            else if (expenseAmount < 0)
                            {
                                creditTotalAmount += Math.Abs(expenseAmount);
                            }

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", expenseName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.ExecuteNonQuery();
                            }

                            Console.WriteLine($"Inserted Expense: {expenseName}, Debit: {debit}, Credit: {credit}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing check data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }*/


        public static void InsertDataToCheckVoucherCompiled(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Clear old data
                string deleteQuery = "DELETE FROM CheckVoucherCompiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from CheckVoucherCompiled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // INSERT QUERY NOW HAS Memo + CustomerJob
                string insertQuery = @"
                                        INSERT INTO CheckVoucherCompiled 
                                        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                                        VALUES 
                                        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                foreach (var check in checkData)
                {
                    try
                    {
                        // COMMON FIELDS
                        string memoValue = string.IsNullOrEmpty(check.ExpensesMemo) ? "" : check.ExpensesMemo;
                        string customerJob = string.IsNullOrEmpty(check.ExpensesCustomerJob) ? "" : check.ExpensesCustomerJob;

                        //
                        // INSERT ITEM ENTRY
                        //
                        if (!string.IsNullOrEmpty(check.ItemName))
                        {
                            string itemName = check.ItemName;
                            string itemClass = check.ItemClass;
                            double itemAmount = check.ItemAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", itemName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }

                        //
                        // INSERT EXPENSE ENTRY
                        //
                        if (!string.IsNullOrEmpty(check.AccountNameCheck))
                        {
                            string expenseName = check.AccountNumber + " - " + check.AccountNameCheck;
                            string expenseClass = check.AccountClassCheck;
                            double expenseAmount = check.ExpensesAmount;

                            string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                            string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                            if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                            else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", expenseName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing check data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }


        public static void InsertDataToBillCompiled(string refNumber, List<BillTable> bills)
        {
            string connectionString = AccessToDatabase.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Clear old data
                string deleteQuery = "DELETE FROM Bill_Compiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from Bill_Compiled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Insert query
                string insertQuery = @"INSERT INTO Bill_Compiled 
                       (RefNumber, Particulars, [Class], [Memo], [CustomerJob], Debit, Credit) 
                       VALUES (@RefNumber, @Particulars, @Class, @Memo, @CustomerJob, @Debit, @Credit)";

                foreach (var bill in bills)
                {
                    foreach (var detail in bill.ItemDetails)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(detail.ItemLineItemRefFullName))
                            {
                                string particulars = detail.ItemLineItemRefFullName ?? "";
                                string itemClass = detail.ItemLineClassRefFullName ?? "";
                                string memo = detail.ItemLineMemo ?? "";
                                string customerJob = detail.ItemLineCustomerJob ?? "";

                                double amount = detail.ItemLineAmount;

                                string debit = amount > 0 ? amount.ToString("N2") : "";
                                string credit = amount < 0 ? Math.Abs(amount).ToString("N2") : "";

                                if (amount > 0) debitTotalAmount += amount;
                                else if (amount < 0) creditTotalAmount += Math.Abs(amount);

                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber);
                                    command.Parameters.AddWithValue("@Particulars", particulars);
                                    command.Parameters.AddWithValue("@Class", itemClass);
                                    command.Parameters.AddWithValue("@Memo", memo);
                                    command.Parameters.AddWithValue("@CustomerJob", customerJob);
                                    command.Parameters.AddWithValue("@Debit", debit);
                                    command.Parameters.AddWithValue("@Credit", credit);
                                    command.ExecuteNonQuery();
                                }

                                Console.WriteLine($"Inserted Item: {particulars}, Debit: {debit}, Credit: {credit}");
                            }

                            if (!string.IsNullOrEmpty(detail.ExpenseLineItemRefFullName))
                            {
                                string particulars =
                                    (bill.AccountNumber != null ? bill.AccountNumber + " - " : "") +
                                    detail.ExpenseLineItemRefFullName;

                                string expClass = detail.ExpenseLineClassRefFullName ?? "";
                                string memo = detail.ExpenseLineMemo ?? "";
                                string customerJob = detail.ExpenseLineCustomerJob ?? "";

                                double amount = detail.ExpenseLineAmount;

                                string debit = amount > 0 ? amount.ToString("N2") : "";
                                string credit = amount < 0 ? Math.Abs(amount).ToString("N2") : "";

                                if (amount > 0) debitTotalAmount += amount;
                                else if (amount < 0) creditTotalAmount += Math.Abs(amount);

                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber);
                                    command.Parameters.AddWithValue("@Particulars", particulars);
                                    command.Parameters.AddWithValue("@Class", expClass);
                                    command.Parameters.AddWithValue("@Memo", memo);
                                    command.Parameters.AddWithValue("@CustomerJob", customerJob);
                                    command.Parameters.AddWithValue("@Debit", debit);
                                    command.Parameters.AddWithValue("@Credit", credit);
                                    command.ExecuteNonQuery();
                                }

                                Console.WriteLine($"Inserted Expense: {particulars}, Debit: {debit}, Credit: {credit}");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error processing bill data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }





        private void SearchBillsByReference(string refNumber)
        {
            if (GlobalVariables.client == "KAYAK")
            {
                AccessQueries queries = new AccessQueries();
                bills = queries.GetBillData_KAYAK(refNumber);
                object data = bills;

                if (bills.Count > 0)
                {
                    Layouts_KAYAK layouts_KAYAK = new Layouts_KAYAK();

                    printDocument = new PrintDocument();
                    printDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                    printPreviewControl.StartPage = 0;

                    printDocument.PrintPage += (s, ev) =>
                    {
                        layouts_KAYAK.PrintPage_KAYAK(s, ev, 1, textBox_SeriesNumber.Text, data);
                    };

                    // 👇 Update panel visibility here
                    panel_Main.Visible = true;
                    panel_Signatory.Visible = true;
                    panel_Main_CR.Visible = false;

                    printPreviewControl.Document = printDocument;
                    printPreviewControl.Visible = true;
                    panel_Printing.Visible = true;
                }
            }
            else if (GlobalVariables.client == "CPI")
            {
                AccessQueries queries = new AccessQueries();
                bills = queries.GetBillData_CPI(refNumber);
                object data = bills;

                if (bills.Count > 0)
                {
                    Layouts_CPI layouts_CPI = new Layouts_CPI();

                    printDocument = new PrintDocument();
                    printDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                    printPreviewControl.StartPage = 0;

                    printDocument.PrintPage += (s, ev) =>
                    {
                        layouts_CPI.PrintPage_CPI(s, ev, 1, textBox_SeriesNumber.Text, data);
                    };

                    // 👇 Update panel visibility here
                    panel_Main.Visible = true;
                    panel_Signatory.Visible = true;
                    panel_Main_CR.Visible = false;

                    printPreviewControl.Document = printDocument;
                    printPreviewControl.Visible = true;
                    panel_Printing.Visible = true;
                }
            }



            else if (GlobalVariables.client == "IVP")
            {
                AccessQueries queries = new AccessQueries();
                bills = queries.GetBillData_CPI(refNumber);
                object data = bills;

                if (bills.Count > 0)
                {
                    Layouts_CPI layouts_CPI = new Layouts_CPI();

                    printDocument = new PrintDocument();
                    printDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                    printPreviewControl.StartPage = 0;

                    printDocument.PrintPage += (s, ev) =>
                    {
                        layouts_CPI.PrintPage_CPI(s, ev, 1, textBox_SeriesNumber.Text, data);
                    };

                    // 👇 Update panel visibility here
                    panel_Main.Visible = true;
                    panel_Signatory.Visible = true;
                    panel_Main_CR.Visible = false;

                    printPreviewControl.Document = printDocument;
                    printPreviewControl.Visible = true;
                    panel_Printing.Visible = true;
                }
            }


            else
            {
                MessageBox.Show("No data found in bills either.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        private FlowLayoutPanel Panel_SBRefNumber()
        {
            FlowLayoutPanel panel_RefNumber = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label label_RefNumberText = new Label
            {
                Parent = panel_RefNumber,
                Width = sideBarWidth - 30,
                Text = "ENTER REFERENCE NUMBER:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            TextBox textBox_ReferenceNumber = new TextBox
            {
                Parent = panel_RefNumber,
                Width = sideBarWidth - 30, // 190
                Font = font_Label,
            };

            Button button_SearchRefNum = new Button
            {
                Parent = panel_RefNumber,
                Height = 26,
                Width = sideBarWidth - 30,
                Text = "SEARCH",
                BackColor = Color.Transparent,
            };

            button_SearchRefNum.Click += (sender, e) =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form.", "Notice", MessageBoxButtons.OK);
                }
                else if (comboBox_Forms.SelectedIndex != 0 && textBox_ReferenceNumber.Text != "")
                {
                    string refNumber = textBox_ReferenceNumber.Text;
                    AccessQueries queries = new AccessQueries();

                    cheque = new List<CheckTable>();
                    bills = new List<BillTable>();
                    checks = new List<CheckTableExpensesAndItems>();
                    receipts = new List<ItemReciept>();
                    apvData = new List<BillTable>();

                    object data = null;
                    if (GlobalVariables.client == "LEADS")
                    {
                        if (comboBox_Forms.SelectedIndex == 1) // Check
                        {
                            cheque = queries.GetCheckData(refNumber);
                            data = cheque;
                        }
                        else if (comboBox_Forms.SelectedIndex == 2) // CV
                        {
                            checks = queries.GetCheckExpensesAndItemsData_LEADS(refNumber);
                            if (checks.Count == 0)
                            {
                                bills = queries.GetBillData_LEADS(refNumber);
                                data = bills;
                            }
                            else
                            {
                                data = checks;
                            }
                        }
                        else if (comboBox_Forms.SelectedIndex == 3) // APV
                        {
                            apvData = queries.GetAccountsPayableData_LEADS(refNumber);
                            data = apvData;
                        }
                        else if (comboBox_Forms.SelectedIndex == 4)
                        {
                            receipts = queries.GetItemRecieptData_LEADS(refNumber);
                            data = receipts;
                        }
                    }
                    else if (GlobalVariables.client == "KAYAK")
                    {
                        if (comboBox_Forms.SelectedIndex == 1) // CV
                        {
                            checks = queries.GetCheckExpensesAndItemsData_KAYAK(refNumber);
                            if (checks.Count == 0)
                            {
                                bills = queries.GetBillData_KAYAK(refNumber);
                                data = bills;
                            }
                            else
                            {
                                data = checks;
                            }
                        }
                    }
                    else if (GlobalVariables.client == "CPI")
                    {
                        if (comboBox_Forms.SelectedIndex == 1) // CV
                        {
                            checks = queries.GetCheckExpensesAndItemsData_CPI(refNumber);
                            if (checks.Count == 0)
                            {
                                bills = queries.GetBillData_CPI(refNumber);
                                data = bills;
                            }
                            else
                            {
                                data = checks;
                            }
                        }
                        if (comboBox_Forms.SelectedIndex == 2) // Check
                        {
                            cheque = queries.GetCheckData(refNumber);
                            data = cheque;
                        }
                    }


                    else if (GlobalVariables.client == "IVP")
                    {
                        if (comboBox_Forms.SelectedIndex == 1) // CV
                        {
                            checks = queries.GetCheckExpensesAndItemsData_CPI(refNumber);
                            if (checks.Count == 0)
                            {
                                bills = queries.GetBillData_CPI(refNumber);
                                data = bills;
                            }
                            else
                            {
                                data = checks;
                            }
                        }
                        if (comboBox_Forms.SelectedIndex == 2) // Check
                        {
                            cheque = queries.GetCheckData(refNumber);
                            data = cheque;
                        }
                    }

                    //if (checks.Count > 0 || bills.Count > 0 || receipts.Count > 0)
                    if (data is System.Collections.ICollection colletion && colletion.Count > 0)
                    {
                        if (GlobalVariables.client == "LEADS")
                        {
                            Layouts_LEADS layouts_LEADS = new Layouts_LEADS();
                            //Layouts layouts = new Layouts();

                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Reset counters for new print job
                            itemCounter = 0;
                            pageCounter = 1;

                            int totalItemDetails = 0;
                            if (comboBox_Forms.SelectedIndex == 3) // APV
                            {
                                // Calculate the total number of pages
                                totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                                int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                                Console.WriteLine($"Generate: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                                printDocument.PrinterSettings.MaximumPage = totalPages;
                            }
                            
                            // Update preview control to start at the first page
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_LEADS.PrintPage_LEADS(s, ev, selectedIndex, seriesNumber, data);
                                /*if (pageCounter < totalItemDetails)
                                {
                                    pageCounter++;
                                    ev.HasMorePages = pageCounter != totalItemDetails;
                                }*/
                                //layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }
                        else if (GlobalVariables.client == "KAYAK")
                        {
                            Layouts_KAYAK layouts_KAYAK = new Layouts_KAYAK();
                            //Layouts layouts = new Layouts();

                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Reset counters for new print job
                            itemCounter = 0;
                            pageCounter = 1;

                            int totalItemDetails = 0;
                            if (comboBox_Forms.SelectedIndex == 1) // APV
                            {
                                // Calculate the total number of pages
                                totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                                int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                                Console.WriteLine($"Generate: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                                printDocument.PrinterSettings.MaximumPage = totalPages;
                            }

                            // Update preview control to start at the first page
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_KAYAK.PrintPage_KAYAK(s, ev, selectedIndex, seriesNumber, data);
                                /*if (pageCounter < totalItemDetails)
                                {
                                    pageCounter++;
                                    ev.HasMorePages = pageCounter != totalItemDetails;
                                }*/
                                //layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }
                        else if (GlobalVariables.client == "CPI")
                        {
                            Layouts_CPI layouts_KAYAK = new Layouts_CPI();
                            //Layouts layouts = new Layouts();

                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Reset counters for new print job
                            itemCounter = 0;
                            pageCounter = 1;

                            int totalItemDetails = 0;
                            if (comboBox_Forms.SelectedIndex == 1) // APV
                            {
                                // Calculate the total number of pages
                                totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                                int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                                Console.WriteLine($"Generate: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                                printDocument.PrinterSettings.MaximumPage = totalPages;
                            }

                            // Update preview control to start at the first page
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_KAYAK.PrintPage_CPI(s, ev, selectedIndex, seriesNumber, data);
                                /*if (pageCounter < totalItemDetails)
                                {
                                    pageCounter++;
                                    ev.HasMorePages = pageCounter != totalItemDetails;
                                }*/
                                //layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }

                        else if (GlobalVariables.client == "IVP")
                        {
                            Layouts_CPI layouts_KAYAK = new Layouts_CPI();
                            //Layouts layouts = new Layouts();

                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Reset counters for new print job
                            itemCounter = 0;
                            pageCounter = 1;

                            int totalItemDetails = 0;
                            if (comboBox_Forms.SelectedIndex == 1) // APV
                            {
                                // Calculate the total number of pages
                                totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                                int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                                Console.WriteLine($"Generate: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                                printDocument.PrinterSettings.MaximumPage = totalPages;
                            }

                            // Update preview control to start at the first page
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_KAYAK.PrintPage_CPI(s, ev, selectedIndex, seriesNumber, data);
                                /*if (pageCounter < totalItemDetails)
                                {
                                    pageCounter++;
                                    ev.HasMorePages = pageCounter != totalItemDetails;
                                }*/
                                //layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }


                        else
                        {
                            Layouts layouts = new Layouts();

                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }

                        printPreviewControl.Document = printDocument;
                        printPreviewControl.Visible = true;
                        panel_Printing.Visible = true;
                    }
                    else
                    {
                        MessageBox.Show("No data found for the provided reference number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    
                }
                else
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK);
                }
            };
            return panel_RefNumber;
        }

        private FlowLayoutPanel Panel_SBSignatory()
        {
            FlowLayoutPanel panel_Signatory = new FlowLayoutPanel
            {
                //Parent = groupBox_Signatory,
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 141,
                Width = sideBarWidth - 10,
                //BackColor = Color.Transparent,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 0),
                BorderStyle = BorderStyle.FixedSingle,
            };

            Label label_SignatoryText = new Label
            {
                Parent = panel_Signatory,
                Width = sideBarWidth - 30,
                Text = "SIGNATORY",
                TextAlign = ContentAlignment.MiddleCenter,
                //Font = new Font("Microsoft Sans Serif", 8, FontStyle.Bold),
                Font = font_Label,
            };

            ComboBox comboBox_Signatory = new ComboBox
            {
                Parent = panel_Signatory,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };

            if (GlobalVariables.client == "LEADS")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Received By:",
                });
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Received By:",
                });
            }
            else if (GlobalVariables.client == "CPI")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Received By:",
                });
            }

            else if (GlobalVariables.client == "IVP")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Released By:",
                });
            }


            else
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Noted By:",
                });
            }
            
            comboBox_Signatory.SelectedIndex = 0;

            Label label_SignatoryName = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Name:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            TextBox textBox_SignatoryName = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165, // 250
                Font = new Font("Microsoft Sans Serif", 8),
            };

            Label label_SignatoryPosition = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Position:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            TextBox textBox_SignatoryPosition = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165, // 250
                Font = new Font("Microsoft Sans Serif", 8),
            };

            Button button_SaveSignatory = new Button
            {
                Parent = panel_Signatory,
                Height = 25,
                Width = 100,
                Text = "SAVE",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                BackColor = Color.Transparent,
            };
            
            Label label_SignatoryStatus = new Label
            {
                Parent = panel_Signatory,
                Height = 22,
                Width = 110,
                //Text = "Saved!",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveSignatory.Click += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    MessageBox.Show("Please selecet an option");
                }
                else
                {
                    string signatoryName = textBox_SignatoryName.Text;
                    string signatoryPosition = textBox_SignatoryPosition.Text;

                    int choice = comboBox_Signatory.SelectedIndex;

                    accessToDatabase.SaveSignatoryData(choice, signatoryName, signatoryPosition);
                    label_SignatoryStatus.Text = "Saved";
                }
            };

            comboBox_Signatory.SelectedIndexChanged += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    textBox_SignatoryName.Text = "";
                    textBox_SignatoryPosition.Text = "";
                }
                else
                {
                    label_SignatoryStatus.Text = "";
                    int choice = comboBox_Signatory.SelectedIndex;
                    var signatoryData = accessToDatabase.RetrieveSignatoryData(choice);

                    textBox_SignatoryName.Text = signatoryData.Name;
                    textBox_SignatoryPosition.Text = signatoryData.Position;
                }
            };

            return panel_Signatory;
        }

        private FlowLayoutPanel Panel_SBRRSignatory()
        {
            FlowLayoutPanel panel_RRSignatory = new FlowLayoutPanel
            {
                //Parent = groupBox_Signatory,
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 106,
                Width = sideBarWidth - 10,
                //BackColor = Color.Transparent,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 0),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label panel_Title = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "SIGNATORY (RR)",
                Width = sideBarWidth - 30,
                //BackColor = Color.SandyBrown,
                TextAlign = ContentAlignment.MiddleCenter,
            };

            Label label_ReceivedBy = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "Received By:",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 71,
                //BackColor = Color.ForestGreen,
            };

            textBox_ReceivedByRR = new TextBox
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Width = 145,
                Margin = new Padding(0, 2, 0, 0),
            };

            Label label_CheckedBy = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "Checked By:",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 71,
                //BackColor = Color.ForestGreen,
            };

            textBox_CheckedByRR = new TextBox
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Width = 145,
                Margin = new Padding(0, 2, 0, 0),
            };

            Button button_SaveRRSignatory = new Button
            {
                Parent = panel_RRSignatory,
                Height = 25,
                Width = 100,
                Text = "SAVE",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                BackColor = Color.Transparent,
            };

            label_SignatoryRRStatus = new Label
            {
                Parent = panel_RRSignatory,
                Height = 22,
                Width = 110,
                //Text = "Saved!",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveRRSignatory.Click += (sender, e) =>
            {
                string signatoryName = textBox_ReceivedByRR.Text;
                string signatoryPosition = textBox_CheckedByRR.Text;

                //int choice = comboBox_Signatory.SelectedIndex;

                accessToDatabase.SaveSignatoryRRData(signatoryName, signatoryPosition);
                label_SignatoryRRStatus.Text = "Saved";
            };

            return panel_RRSignatory;
        }

        private FlowLayoutPanel Panel_SBPrinting()
        {
            panel_Printing = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 110,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false,
            };

            Button button_ZoomOut = new Button
            {
                Parent = panel_Printing,
                Text = "Zoom Out",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_ZoomOut.Click += (sender, e) =>
            {
                if (printPreviewControl.Zoom >= 0.1)
                {
                    printPreviewControl.Zoom -= 0.1;
                }
            };

            Button button_ZoomIn = new Button
            {
                Parent = panel_Printing,
                Text = "Zoom In",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_ZoomIn.Click += (sender, e) =>
            {
                printPreviewControl.Zoom += 0.1;
            };

            Button button_PreviousPage = new Button
            {
                Parent = panel_Printing,
                Text = "Previous Page",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_PreviousPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage > 0)
                {
                    printPreviewControl.StartPage--;
                }
            };

            Button button_NextPage = new Button
            {
                Parent = panel_Printing,
                Text = "Next Page",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_NextPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage < pageCounter - 1)
                {
                    printPreviewControl.StartPage++;
                }
            };

            Button button_Print = new Button
            {
                Parent = panel_Printing,
                Text = "Print",
                Height = 28,
                Width = 222,
                BackColor = Color.Transparent,
            };
            button_Print.Click += (sender, e) =>
            {
                try
                {
                    // Reset counters for new print job
                    itemCounter = 0;
                    pageCounter = 1;

                    if (comboBox_Forms.SelectedIndex == 3) // APV
                    {
                        /*// Calculate the total number of pages
                        int totalPages = (int)Math.Ceiling((double)apvData.Count / itemsPerPage);
                        printDocument.PrinterSettings.MaximumPage = totalPages;*/
                        // Calculate the total number of pages
                        int totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                        int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                        Console.WriteLine($"Print: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                        printDocument.PrinterSettings.MaximumPage = totalPages;
                    }
                    
                    // Update preview control to start at the first page
                    printPreviewControl.StartPage = 0;

                    PrintDialog printDialog = new PrintDialog
                    {
                        Document = printDocument,
                    };

                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        GlobalVariables.includeImage = false;

                        /*if (comboBox_Forms.SelectedIndex == 1)
                        {
                            GlobalVariables.isPrinting = true;
                        }*/
                        printDialog.Document.Print();
                        printPreviewControl.Visible = false;
                        printPreviewControl.Zoom = 1;
                        panel_Printing.Visible = false;
                        

                        if (GlobalVariables.client == "LEADS")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
                        }
                        else if (GlobalVariables.client == "KAYAK")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                        }
                        else if (GlobalVariables.client == "CPI")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                        }

                        else if (GlobalVariables.client == "IVP")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while printing: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                GlobalVariables.includeImage = true;
            };

            return panel_Printing;
        }

        private void ComboBox_Forms_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GlobalVariables.client == "LEADS")
            {
                string prefix = "";
                //panel_SeriesNumber.Visible = false;

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // Check
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = true;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_RRSignatory.Visible = false;

                        panel_Main.Visible = true;
                        panel_Main_CR.Visible = false;
                        break;
                    case 2: // CV
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = true;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = true;
                        panel_RRSignatory.Visible = false;
                        label_SeriesNumberText.Text = "Current Series Number: CV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("CVSeries");

                        panel_Main.Visible = true;
                        panel_Main_CR.Visible = false;
                        break;

                    case 3: // APV
                        prefix = "APV";
                        panel_SeriesNumber.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: APV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("APVSeries");

                        if (GlobalVariables.useCrystalReports_LEADS)
                        {
                            panel_RefNumber.Visible = false;
                            panel_RefNumberCrystalReport.Visible = true;
                            panel_Main.Visible = false;
                            panel_Main_CR.Visible = true;
                        }
                        else
                        {
                            panel_RefNumber.Visible = true;
                            panel_RefNumberCrystalReport.Visible = false;
                            panel_Main.Visible = true;
                            panel_Main_CR.Visible = false;
                        }

                        panel_Signatory.Visible = true;
                        panel_RRSignatory.Visible = false;
                        break;

                    case 4: // RR
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = true;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_RRSignatory.Visible = true;

                        var text = accessToDatabase.RetrieveSignatoryRRData();
                        label_SignatoryRRStatus.Text = "";
                        textBox_ReceivedByRR.Text = text.ReceivedBy;
                        textBox_CheckedByRR.Text = text.CheckedBy;

                        panel_Main.Visible = true;
                        panel_Main_CR.Visible = false;
                        break;

                    default:
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_RRSignatory.Visible = false;
                        panel_SeriesNumber.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = false;
                        return;
                }

                UpdateSeriesNumber(prefix);
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                string prefix = "";
                //panel_SeriesNumber.Visible = false;

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // CV
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("CVSeries");

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    default:
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_SeriesNumber.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = false;
                        return;
                }

                UpdateSeriesNumber(prefix);
            }
            else if (GlobalVariables.client == "CPI")
            {
                string prefix = "";
                //panel_SeriesNumber.Visible = false;

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // CV
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("CVSeries");

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;
                    case 2: // Check
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = true;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;

                        panel_Main.Visible = true;
                        panel_Main_CR.Visible = false;
                        break;

                    default:
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_SeriesNumber.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = false;
                        return;
                }

                UpdateSeriesNumber(prefix);
            }


            else if (GlobalVariables.client == "IVP")
            {
                string prefix = "";

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // CV
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("CVSeries");

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 2: // Journal Voucher (New Addition)
                        prefix = "JV";

                        // PANEL SETTINGS AS REQUESTED
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        panel_SeriesNumber.Visible = true;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;

                        // Series Number Logic
                        label_SeriesNumberText.Text = "Current Series Number: JV";
                        // Ensure you have "JVSeries" in your database, or this returns 0/default
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("JVSeries");
                        break;

                    default:
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_SeriesNumber.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = false;
                        return;
                }

                UpdateSeriesNumber(prefix);
            }



            else
            {
                if (comboBox_Forms.SelectedIndex == 1) // CV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: CV";
                    textBox_SeriesNumber.Text = "CV" + seriesNumber;
                }
            }
        }

        private void SetDatabaseLocation(ReportDocument reportDocument, string databasePath)
        {
            // Iterate through each table in the report
            foreach (Table table in reportDocument.Database.Tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;

                // Update the connection information
                tableLogOnInfo.ConnectionInfo.ServerName = databasePath;
                tableLogOnInfo.ConnectionInfo.DatabaseName = ""; //or databasePath
                tableLogOnInfo.ConnectionInfo.UserID = ""; // Leave blank for Access
                tableLogOnInfo.ConnectionInfo.Password = ""; // Leave blank for Access

                // Apply the updated information to the table
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

            // Update subreports if any
            foreach (Section section in reportDocument.ReportDefinition.Sections)
            {
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        SubreportObject subreportObject = (SubreportObject)reportObject;
                        ReportDocument subreportDocument = subreportObject.OpenSubreport(subreportObject.SubreportName);
                        SetDatabaseLocation(subreportDocument, databasePath);
                    }
                }
            }
        }

        private void SetDatabaseLocationIVP(ReportDocument reportDocument, string databasePath)
        {

            ConnectionInfo connectionInfo = new ConnectionInfo();
            connectionInfo.ServerName = databasePath;
            connectionInfo.DatabaseName = "";
            connectionInfo.UserID = "";
            connectionInfo.Password = "";
            connectionInfo.Type = ConnectionInfoType.CRQE;
            connectionInfo.IntegratedSecurity = false;

            foreach (Table table in reportDocument.Database.Tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                tableLogOnInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

            foreach (ReportDocument subreport in reportDocument.Subreports)
            {
                foreach (Table table in subreport.Database.Tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                    tableLogOnInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                }
            }

            reportDocument.VerifyDatabase();
        }


        private void TextBox_SeriesNumber_TextChanged(object sender, EventArgs e)
        {
            if (GlobalVariables.client == "LEADS")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    string prefix = comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV";
                    string input = textBox_SeriesNumber.Text.Replace(prefix, "").Trim();

                    if (int.TryParse(input, out int adjustedSeries))
                    {
                        seriesNumber = adjustedSeries;
                    }
                    else
                    {
                        MessageBox.Show("Invalid series number format. Please enter a numeric value.");
                        textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Revert to the current value
                    }
                }
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    string prefix = comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV";
                    string input = textBox_SeriesNumber.Text.Replace(prefix, "").Trim();

                    if (int.TryParse(input, out int adjustedSeries))
                    {
                        seriesNumber = adjustedSeries;
                    }
                    else
                    {
                        MessageBox.Show("Invalid series number format. Please enter a numeric value.");
                        textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Revert to the current value
                    }
                }
            }
            else if (GlobalVariables.client == "CPI")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    string prefix = comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV";
                    string input = textBox_SeriesNumber.Text.Replace(prefix, "").Trim();

                    if (int.TryParse(input, out int adjustedSeries))
                    {
                        seriesNumber = adjustedSeries;
                    }
                    else
                    {
                        MessageBox.Show("Invalid series number format. Please enter a numeric value.");
                        textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Revert to the current value
                    }
                }
            }
            else if (GlobalVariables.client == "IVP")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    // Determine prefix
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 2) prefix = "JV";

                    string input = textBox_SeriesNumber.Text.Replace(prefix, "").Trim();

                    if (int.TryParse(input, out int adjustedSeries))
                    {
                        seriesNumber = adjustedSeries;
                    }
                    else
                    {
                        MessageBox.Show("Invalid series number format. Please enter a numeric value.");
                        textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
                    }
                }
            }

        }
        private void TextBox_SeriesNumber_Leave(object sender, EventArgs e)
        {
            if (GlobalVariables.client == "LEADS")
            {
                string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
                accessToDatabase.UpdateManualSeriesNumber(columnName, seriesNumber); // Save manual adjustment
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                accessToDatabase.UpdateManualSeriesNumber(columnName, seriesNumber); // Save manual adjustment
            }
            else if (GlobalVariables.client == "CPI")
            {
                string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                accessToDatabase.UpdateManualSeriesNumber(columnName, seriesNumber); // Save manual adjustment
            }
            else if (GlobalVariables.client == "IVP")
            {
                string columnName = "";

                if (comboBox_Forms.SelectedIndex == 1) columnName = "CVSeries";
                else if (comboBox_Forms.SelectedIndex == 2) columnName = "JVSeries"; // <--- Added this

                if (!string.IsNullOrEmpty(columnName))
                {
                    accessToDatabase.UpdateManualSeriesNumber(columnName, seriesNumber);
                }
            }
        }

        private void UpdateSeriesNumber(string prefix)
        {
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Formats seriesNumber as a 3-digit number
        }
        private void RefreshSeriesNumber(string columnName)
        {
            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
            string prefix = comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV";
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
        }
    }
}
