using System.Windows;
using System.Collections.Generic;
using ReconRunner.Controller;
using System.Windows.Forms;
using System;

namespace RRWpfUi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string cr = "\r\n";
        private RRController rrController = RRController.Instance;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Button Click Methods
        private void btnLoadSources_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Sources XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += string.Format("Reading sources from file {0}.{1}", xmlFileOpenDialog.FileName, cr);
                txtStatus.Text += rrController.ReadSourcesFromXMLFile(xmlFileOpenDialog.FileName) + cr;

            }
            else
            {
                txtStatus.Text += "Creation of Sources collection cancelled." + cr;
            }
            checkReadyToRun();
        }

        private void btnLoadRecons_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Recons XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += string.Format("Reading recons from file {0}.{1}", xmlFileOpenDialog.FileName, cr);
                txtStatus.Text += rrController.ReadReconsFromXMLFile(xmlFileOpenDialog.FileName) + cr;
            }
            else
            {
                txtStatus.Text += "Creation of Recon collection cancelled." + cr;
            }
            checkReadyToRun();
        }

        private void checkReadyToRun()
        {
            btnRunRecons.IsEnabled = rrController.ReadyToRun();
            btnValidate.IsEnabled = !btnRunRecons.IsEnabled;
            if (btnRunRecons.IsEnabled)
                txtStatus.Text += "No validation errors. Ready to run recons." + cr;
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
                    txtStatus.Text += string.Format("Processing recon reports and creating {0}.{1}", fileSaveDialog.FileName, cr);
                    string results = rrController.RunRecons(fileSaveDialog.FileName);
                    txtStatus.Text += results;
                }
                else
                {
                    txtStatus.Text += "Creation of Recon report Excel file cancelled." + cr;
                }
            }
            catch (Exception ex)
            {
                var fullErrorMessage = rrController.GetFullErrorMessage(ex);
               txtStatus.Text += string.Format("Error while creating samples: {0}{1}", cr, fullErrorMessage);
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
                    txtStatus.Text += string.Format("Writing sources to file {0}.{1}",fileSaveDialog.FileName, cr);
                    rrController.CreateSampleXMLSourcesFile(fileSaveDialog.FileName);
                    txtStatus.Text += "Sources XML file created." + cr;
                }
                else
                {
                    txtStatus.Text += "Creation of Sources XML file cancelled." + cr;
                }

                // Save recons
                fileSaveDialog = getFileSaveDialog("Save recons to file");
                fileSaveDialog.Title = "Save Recons";
                xmlFile = fileSaveDialog.ShowDialog();

                if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    txtStatus.Text += string.Format("Writing recons to file {0}.{1}", fileSaveDialog.FileName, cr);
                    rrController.CreateSampleXMLReconReportFile(fileSaveDialog.FileName);
                    txtStatus.Text += "Recons XML file created." + cr;
                }
                else
                {
                    txtStatus.Text += "Creation of Recons XML file cancelled." + cr;
                }
            }
            catch (Exception ex)
            {
                var fullErrorMessage = rrController.GetFullErrorMessage(ex);
                txtStatus.Text += string.Format("Error while creating samples: {0}{1}", cr,fullErrorMessage);
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
            txtStatus.Text += "Starting validation..." + cr;
            var validationMessages = new List<string>();
            validationMessages.AddRange(rrController.GetValidationErrors());
            //validationMessages.AddRange(rrController.GetValidationWarnings());
            if (validationMessages.Count == 0)
                txtStatus.Text += "No validation errors or warnings found." + cr;
            else
            {
                validationMessages.ForEach(message => txtStatus.Text += message + cr);
                txtStatus.Text += "Done." + cr;
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
        #endregion Dialogs  
    }
}
