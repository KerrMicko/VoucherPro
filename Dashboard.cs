using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using VoucherPro.Clients;
using static VoucherPro.DataClass;

namespace VoucherPro
{
    public class GlobalVariables
    {
        public static string client = "LEADS";
        public static bool includeImage = true;
        public static bool includeItemReceipt = true;
        public static bool testWithoutData = false;
        public static bool isPrinting = false;
        public static int itemsPerPageAPV = 10;
    }
    public partial class Dashboard : Form
    {
        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private AccessToDatabase accessToDatabase;

        ComboBox comboBox_Forms;

        Label label_SeriesNumberText;

        TextBox textBox_SeriesNumber;

        FlowLayoutPanel panel_Printing;
        FlowLayoutPanel panel_SeriesNumber;

        List<CheckTable> cheque = new List<CheckTable>();
        List<BillTable> bills = new List<BillTable>();
        List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();
        List<ItemReciept> receipts = new List<ItemReciept>();
        List<BillTable> apvData = new List<BillTable>();

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
            Panel panel_Main = MainPanel();
            Panel panel_SideBar = SideBarPanel();

            panel_SideBar.Parent = panel_Container;
            panel_Title.Parent = panel_Container;
            panel_Main.Parent = panel_Container;

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

            // - REF NUMBER ---------------------------------------------
            FlowLayoutPanel panel_RefNumber = Panel_SBRefNumber();
            panel_RefNumber.Parent = panel_SideBar;

            // - SIGNATORY ----------------------------------------------
            FlowLayoutPanel panel_Signatory = Panel_SBSignatory();
            panel_Signatory.Parent = panel_SideBar;
            
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

            /*if (GlobalVariables.client == "LEADS")
            {
                comboBox_Forms.Items.AddRange(new string[]
                {
                    "",
                    "Collection Receipt", // smol
                    "Delivery Receipt",
                    "Provisional Receipt", // smol
                    "Sales Invoice",
                    "Service Invoice",
                });
                comboBox_Forms.SelectedIndex = 0;
            }
            else
            {
                comboBox_Forms.Items.AddRange(new string[]
                {
                    "",
                    "Credit Note - BIR",
                    "Credit Note - Non BIR",
                    "Debit Note - BIR",
                    "Debit Note - Non BIR",
                    "Delivery Receipt",
                    "Office Receipt",
                    "Office Receipt - Non BIR",
                    "Sales Invoice",
                    "Sales Invoice Summary",
                });
                comboBox_Forms.SelectedIndex = 0;
            }*/

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
                if (seriesNumber != 0)
                {
                    seriesNumber--;
                    UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
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
                seriesNumber++;
                UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
            };

            return panel_SeriesNumber;
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

                    //if (checks.Count > 0 || bills.Count > 0 || receipts.Count > 0)
                    if (data is System.Collections.ICollection colletion && colletion.Count > 0)
                    {
                        if (GlobalVariables.client == "LEADS")
                        {
                            Layouts_LEADS layouts_LEADS = new Layouts_LEADS();
                            //Layouts layouts = new Layouts();

                            PaperSize paperSize = new PaperSize("Custom", 850, 1100);

                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Reset counters for new print job
                            itemCounter = 0;
                            pageCounter = 1;

                            if (comboBox_Forms.SelectedIndex == 3) // APV
                            {
                                // Calculate the total number of pages
                                int totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                                int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                                Console.WriteLine($"Generate: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                                printDocument.PrinterSettings.MaximumPage = totalPages;
                            }
                            
                            // Update preview control to start at the first page
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_LEADS.PrintPage_LEADS(s, ev, selectedIndex, seriesNumber, data);
                                //layouts.PrintPage(s, ev, comboBox_Forms.SelectedIndex);
                            };
                        }
                        else
                        {
                            Layouts layouts = new Layouts();

                            PaperSize paperSize = new PaperSize("Custom", 850, 1100);

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
            comboBox_Signatory.Items.AddRange(new string[]
            {
                "Select Signatory Option",
                "Prepared By:",
                "Checked By:",
                "Approved By:",
                "Noted By:",
            });
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
                        printDialog.Document.Print();
                        printPreviewControl.Visible = false;
                        printPreviewControl.Zoom = 1;
                        panel_Printing.Visible = false;
                        
                        string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
                        accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print
                      
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                        UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
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
                /*if (comboBox_Forms.SelectedIndex == 0)
                {
                    panel_SeriesNumber.Visible = false;
                }
                else if (comboBox_Forms.SelectedIndex == 1) // Check
                {
                    panel_SeriesNumber.Visible = false;
                }
                else if (comboBox_Forms.SelectedIndex == 2) // CV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: CV";
                    textBox_SeriesNumber.Text = "CV" + seriesNumber.ToString("D3");
                }
                else if (comboBox_Forms.SelectedIndex == 3) // APV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: APV";
                    textBox_SeriesNumber.Text = "APV" + seriesNumber.ToString("D3");
                }
                else if (comboBox_Forms.SelectedIndex == 4)
                {
                    panel_SeriesNumber.Visible = false;
                }*/

                string prefix = "";
                panel_SeriesNumber.Visible = false;

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 2: // CV
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("CVSeries");
                        break;

                    case 3: // APV
                        prefix = "APV";
                        panel_SeriesNumber.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: APV";
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase("APVSeries");
                        break;

                    default:
                        panel_SeriesNumber.Visible = false;
                        return;
                }

                UpdateSeriesNumber(prefix);
            }
            else
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    panel_SeriesNumber.Visible = false;
                }
                else if (comboBox_Forms.SelectedIndex == 1) // Check
                {
                    panel_SeriesNumber.Visible = false;
                }
                else if (comboBox_Forms.SelectedIndex == 2) // CV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: CV";
                    textBox_SeriesNumber.Text = "CV" + seriesNumber;
                }
                else if (comboBox_Forms.SelectedIndex == 3) // APV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: APV";
                    textBox_SeriesNumber.Text = "test2" + seriesNumber;
                }
                else if (comboBox_Forms.SelectedIndex == 4)
                {
                    panel_SeriesNumber.Visible = false;
                }
            }
        }

        private void TextBox_SeriesNumber_TextChanged(object sender, EventArgs e)
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
        private void TextBox_SeriesNumber_Leave(object sender, EventArgs e)
        {
            string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
            accessToDatabase.UpdateManualSeriesNumber(columnName, seriesNumber); // Save manual adjustment
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
