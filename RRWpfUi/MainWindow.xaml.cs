using System.Windows;
using System.Collections.Generic;
using ReconRunner.Controller;
using System.Windows.Forms;
using System;
using ReconRunner.Model;
using System.Deployment;
using System.Threading.Tasks;

namespace RRWpfUi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string cr = "\r\n";
        private string noReconText = "No recon reports loaded";
        private string chooseReconText = "Click to choose a report to edit";
        private RRController rrController = RRController.Instance;
        public ActionStatusEventHandler ActionStatusEvent;

        public MainWindow()
        {
            InitializeComponent();
            lblVersion.Content = getVersionText();
            rrController.ActionStatusEvent += new ActionStatusEventHandler(handleActionStatus);
            enableEditingControls(false);
        }

        #region Button Click Methods
        private void btnLoadSources_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Sources XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                rrController.ReadSourcesFromXMLFile(xmlFileOpenDialog.FileName);
                updateStatus(string.Format("Read sources from file {0}.", xmlFileOpenDialog.FileName));
            }
            else
                updateStatus("Creation of Sources collection cancelled.");

            checkReadyToRun();
        }

        private void btnLoadRecons_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Recons XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                rrController.ReadReconsFromXMLFile(xmlFileOpenDialog.FileName);
                loadCmbChooseRecon();
                updateStatus(string.Format("Loaded recons from file {0}", xmlFileOpenDialog.FileName));
            }
            else
                updateStatus("Creation of Recon collection cancelled.");

            checkReadyToRun();
        }

        private void loadCmbChooseRecon()
        {

            if (rrController.Recons.ReconReports.Count == 0)
            {
                cmbChooseItem.Items.Clear();
                cmbChooseItem.Items.Add(noReconText);
                cmbChooseItem.SelectedItem = cmbChooseItem.Items[0];
                cmbChooseItem.IsEnabled = false;
            }
            else
            {
                cmbChooseItem.Items.Clear();
                cmbChooseItem.Items.Add(chooseReconText);
                rrController.Recons.ReconReports.ForEach(recon => cmbChooseItem.Items.Add(recon.Name));
                cmbChooseItem.SelectedItem = cmbChooseItem.Items[0];
                cmbChooseItem.IsEnabled = true;
            }
        }

        private void checkReadyToRun()
        {
            btnRunRecons.IsEnabled = rrController.ReadyToRun();
            btnValidate.IsEnabled = !btnRunRecons.IsEnabled;
            if (btnRunRecons.IsEnabled)
                updateStatus("No validation errors. Ready to run recons.");
        }

        private void btnRunRecons_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var fileSaveDialog = getFileSaveDialog("Save recon results to spreadsheet");
                fileSaveDialog.DefaultExt = "xls";
                fileSaveDialog.Filter = "Excel Files|*.xlsx|All files|*.*";
                DialogResult excelFile = fileSaveDialog.ShowDialog();

                if (excelFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    updateStatus(string.Format("Processing recon reports and creating {0}.", fileSaveDialog.FileName));
                    Task.Factory.StartNew(() =>
                    {
                        rrController.RunRecons(fileSaveDialog.FileName);
                    });
                    updateStatus("Done.");
                }
                else
                {
                    updateStatus("Creation of Recon report Excel file cancelled.");
                }
            }
            catch (Exception ex)
            {
                var fullErrorMessage = rrController.GetFullErrorMessage(ex);
                updateStatus(string.Format("Error while running recons: {0}", fullErrorMessage));
            }
        }


        private void btnCreateSamples_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Save sources
                var fileSaveDialog = getFileSaveDialog("Save sources to file");
                fileSaveDialog.Title = "Save Sources";
                DialogResult xmlFile = fileSaveDialog.ShowDialog();

                if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    rrController.CreateSampleXMLSourcesFile(fileSaveDialog.FileName);
                    updateStatus(string.Format("Wrote sample sources to file {0}", fileSaveDialog.FileName));
                }
                else
                    updateStatus("Creation of Sources XML file cancelled.");

                // Save recons
                fileSaveDialog = getFileSaveDialog("Save recons to file");
                fileSaveDialog.Title = "Save Recons";
                xmlFile = fileSaveDialog.ShowDialog();

                if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    rrController.CreateSampleXMLReconReportFile(fileSaveDialog.FileName);
                    updateStatus(string.Format("Wrote sample recons to file {0}", fileSaveDialog.FileName));
                }
                else
                    updateStatus("Creation of Recons XML file cancelled.");
            }
            catch (Exception ex)
            {
                var fullErrorMessage = rrController.GetFullErrorMessage(ex);
                updateStatus(string.Format("Error while creating samples: {0}", fullErrorMessage));
            }
        }


        private void btnValidate_Click(object sender, RoutedEventArgs e)
        {
            reportValidationIssues();
        }


        /// <summary>
        /// Write out any validation errors. If no errors are found then also check
        /// for warnings and report any of those.
        /// </summary>
        private void reportValidationIssues()
        {
            updateStatus("Starting validation...");
            var validationMessages = new List<string>();
            validationMessages.AddRange(rrController.GetValidationErrors());
            //validationMessages.AddRange(rrController.GetValidationWarnings());
            if (validationMessages.Count == 0)
                updateStatus("No validation errors or warnings found.");
            else
            {
                validationMessages.ForEach(message => txtStatus.Text += message + cr);
            }
        }
        #endregion Button Click Methods

        #region Dialogs
        /// <summary>
        /// Create save file dialog for XML file.  Optionally pass window title to use
        /// </summary>
        /// <param name="dialogTitle">string.empty for none, otherwise title to show in dialog window</param>
        /// <returns></returns>
        private SaveFileDialog getFileSaveDialog(string dialogTitle)
        {
            var fileSaveDialog = new SaveFileDialog();
            fileSaveDialog.DefaultExt = "xml";
            fileSaveDialog.Filter = "XML Files|*.xml|All files|*.*";
            if (dialogTitle != string.Empty)
                fileSaveDialog.Title = dialogTitle;
            return fileSaveDialog;
        }


        /// <summary>
        /// Create save file dialog for XML file.  Optionally pass window title to use
        /// </summary>
        /// <param name="dialogTitle">string.empty for none, otherwise title to show in dialog window</param>
        /// <returns></returns>
        private OpenFileDialog getFileOpenDialog(string dialogTitle)
        {
            var fileOpenDialog = new OpenFileDialog();
            fileOpenDialog.DefaultExt = "xml";
            fileOpenDialog.Filter = "XML Files|*.xml|All files|*.*";
            if (dialogTitle != string.Empty)
                fileOpenDialog.Title = dialogTitle;
            return fileOpenDialog;
        }

        private string getVersionText()
        {
            string versionText = "Unknown";
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                versionText = "v" + string.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
            }
            else
                versionText = "version n/a until deployed";
            return versionText;
        }
        #endregion Dialogs  

        private void handleActionStatus(object sender, ActionStatusEventArgs e)
        {
            // Use Dispatcher.Invoke so UI responds and shows message immediately
            Dispatcher.Invoke(new Action(() =>
            {
                if (e.State == RequestState.Succeeded || e.State == RequestState.Information)
                    updateStatus(e.Message);
                else
                    // Something other than success status, so show the status as well as the message
                    updateStatus(string.Format("{0} Status: {1}", e.State.ToString(), e.Message));
            }));
        }

        /// <summary>
        /// Currently just appends text to a new line of txtStatus
        /// </summary>
        private void updateStatus(string statusText)
        {
            txtStatus.Text += statusText + cr;
        }

        private void btnClearStatus_Click(object sender, RoutedEventArgs e)
        {
            txtStatus.Text = string.Empty;
        }

        private void btnSaveEdits_Click(object sender, RoutedEventArgs e)
        {
            var fileSaveDialog = getFileSaveDialog("Save recons to file");
            fileSaveDialog.Title = "Save Recons";
            var xmlFile = fileSaveDialog.ShowDialog();

            try
            {
                if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    rrController.WriteReconsToXmlFile(fileSaveDialog.FileName);
                    updateStatus(string.Format("Wrote recons to file {0}", fileSaveDialog.FileName));
                }
                else
                    updateStatus("Creation of Recons XML file cancelled.");
                checkReadyToRun();
            }
            catch (Exception ex)
            {
                updateStatus(string.Format("Unable to save changes to file {0}: {1}", xmlFile, rrController.GetFullErrorMessage(ex)));
            }
        }

        private void cmbChooseRecon_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (rrPropertyGrid != null)
            {
                if (cmbChooseItem.Items.Count == 0 || cmbChooseItem.SelectedIndex == 0)
                {
                    rrPropertyGrid.Visibility = Visibility.Hidden;
                    enableEditingControls(false);
                    btnSaveEdits.IsEnabled = false;
                }
                else
                {
                    rrPropertyGrid.Visibility = Visibility.Visible;
                    rrPropertyGrid.SelectedObject = rrController.Recons.ReconReports.Find(rr => rr.Name == cmbChooseItem.SelectedItem.ToString());
                    enableEditingControls(true);
                    btnSaveEdits.IsEnabled = true;
                }
            }
        }

        private void enableEditingControls(bool isEnabled)
        {
            btnSaveEdits.IsEnabled = isEnabled;
            btnDeleteItem.IsEnabled = isEnabled;
            btnCopyItem.IsEnabled = isEnabled;
        }

        private void btnDeleteItem_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnCopyItem_Click(object sender, RoutedEventArgs e)
        {

        }

        private void saveImage_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            updateStatus("Ping...");
        }
    }
}
