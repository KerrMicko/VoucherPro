using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VoucherPro.Properties;

namespace VoucherPro
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Dashboard());

            bool testingWithoutData = GlobalVariables.testWithoutData;

            if (!testingWithoutData)
            {
                Settings.Default.IsFirstRun = true;
                Settings.Default.Save();

                Dashboard dashboard = new Dashboard();
                dashboard.Load += async (sender, e) =>
                {
                    // Check if it's the first run
                    if (Settings.Default.IsFirstRun)
                    {
                        // Disable the form
                        dashboard.Enabled = false;

                        // Call the function that should only run on the first run
                        await FirstRunFunction();

                        // Enable the form after the function completes
                        dashboard.Enabled = true;

                        // Update the setting to indicate that the program has been run
                        Settings.Default.IsFirstRun = false;
                        Settings.Default.Save();
                    }
                };
                Application.Run(dashboard);
            }
            else
            {
                Application.Run(new Dashboard());
            }
        }

        private static async Task FirstRunFunction()
        {
            if (GlobalVariables.client == "IVP")
            {
                return;
            }
            else
            {
                try
                {
                    //MessageBox.Show("Welcome to the application!");
                    DialogResult result = MessageBox.Show("Welcome! Do you want to sync data from QuickBooks?",
                                         "Sync Confirmation",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        AccessToDatabase accessToDatabase = new AccessToDatabase();
                        accessToDatabase.DeleteSpecifiedTablesData();

                        using (var progressForm = new Form())
                        {
                            progressForm.StartPosition = FormStartPosition.CenterScreen;
                            progressForm.Size = new System.Drawing.Size(300, 100);
                            progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                            progressForm.MaximizeBox = false;
                            progressForm.MinimizeBox = false;
                            progressForm.ControlBox = false;
                            progressForm.Text = "Syncing";

                            var label = new Label
                            {
                                Text = "Syncing data from QuickBooks. Please wait...",
                                Dock = DockStyle.Fill,
                                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                            };

                            progressForm.Controls.Add(label);
                            progressForm.Show();
                            progressForm.BringToFront();

                            await accessToDatabase.FetchAndSaveData();

                            progressForm.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while syncing data: " + ex.Message);
                }
            }
        }
    }
}
