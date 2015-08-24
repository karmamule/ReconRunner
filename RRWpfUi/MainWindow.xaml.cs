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
        enum Entity
        {
            Sources, Recons, None
        }
        private string cr = "\r\n";
        private string noReconText = "No recon reports loaded";
        private string chooseReconText = "Click to choose a report to edit";
        private string noSourcesText = "Sources not loaded yet";
        private string chooseSourcesText = "Click if you want to edit sources";
        private RRController rrController = RRController.Instance;
        public ActionStatusEventHandler ActionStatusEvent;

        public MainWindow()
        {
            InitializeComponent();
            lblVersion.Content = getVersionText();
            rrController.ActionStatusEvent += new ActionStatusEventHandler(handleActionStatus);
            enableEditingControls(Entity.None);
            allowReports(false);
        }

        #region Button Click Methods
        private void btnLoadSources_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Sources XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                rrController.ReadSourcesFromXMLFile(xmlFileOpenDialog.FileName);
                if (radSources.IsChecked ?? false)
                    loadComboBox(Entity.Sources);
                updateStatus(string.Format("Read sources from file {0}.", xmlFileOpenDialog.FileName));
                //if (radSources.IsChecked ?? false)
            }
            else
                updateStatus("Creation of Sources collection cancelled.");

            //checkReadyToRun();
        }

        private void btnLoadRecons_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Recons XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                rrController.ReadReconsFromXMLFile(xmlFileOpenDialog.FileName);
                if (radRecons.IsChecked ?? false)
                    loadComboBox(Entity.Recons);
                updateStatus(string.Format("Loaded recons from file {0}", xmlFileOpenDialog.FileName));
            }
            else
                updateStatus("Creation of Recon collection cancelled.");

            //checkReadyToRun();
        }

        /// <summary>
        /// Load the combo box with either a list of recon reports available for edting
        /// or add entries for allowing to edit Sources info or not
        /// </summary>
        /// <param name="entity"></param>
        private void loadComboBox(Entity entity)
        {
            if (cmbChooseItem != null)
            {
                var noEntityText = entity == Entity.Recons ? noReconText : noSourcesText;
                var chooseEntityText = entity == Entity.Recons ? chooseReconText : chooseSourcesText;
                int entityCount = 0;
                if (entity == Entity.Recons && rrController.Recons != null)
                    entityCount = rrController.Recons.ReconReports.Count;
                else
                    if (entity == Entity.Sources && rrController.Sources != null)
                        entityCount = rrController.Sources.ConnectionStrings.Count;

                cmbChooseItem.Items.Clear();
                if (entityCount == 0)
                {
                    cmbChooseItem.Items.Add(noEntityText);
                    cmbChooseItem.SelectedItem = cmbChooseItem.Items[0];
                    cmbChooseItem.IsEnabled = false;
                }
                else
                {
                    cmbChooseItem.Items.Add(chooseEntityText);
                    cmbChooseItem.SelectedItem = cmbChooseItem.Items[0];
                    if (entity == Entity.Recons)
                        rrController.Recons.ReconReports.ForEach(recon => cmbChooseItem.Items.Add(recon.Name));
                    else
                        cmbChooseItem.Items.Add("Edit sources");
                    cmbChooseItem.IsEnabled = true;
                }
            }
        }

        /*
        private void checkReadyToRun()
        {
            btnRunRecons.IsEnabled = rrController.ReadyToRun();
            btnValidate.IsEnabled = !btnRunRecons.IsEnabled;
            if (btnRunRecons.IsEnabled)
                updateStatus("No validation errors. Ready to run recons.");
        }
        */

        private void allowReports(bool allowReports)
        {
            btnRunRecons.IsEnabled = allowReports;
            btnValidate.IsEnabled = !allowReports;
        }

        private void btnRunRecons_Click(object sender, RoutedEventArgs e)
        {
            if (!rrController.ReadyToRun())
            {
                allowReports(false);
                updateStatus("Data is no longer valid. Did you just change something?");
                reportValidationIssues();
            }
            else
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
                            try
                            {
                                rrController.RunRecons(fileSaveDialog.FileName);
                            }
                            catch (Exception ex)
                            {
                                string.Format("Problem while running recons: {0}", rrController.GetFullErrorMessage(ex));
                            }
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
                // If a message regarding data validity, act accordingly. Otherwise just show it to user
                switch (e.State)
                {
                    case RequestState.DataInvalid:
                        allowReports(false);
                        break;
                    case RequestState.DataValid:
                        allowReports(true);
                        break;
                    case RequestState.Succeeded:
                    case RequestState.Information:
                        updateStatus(e.Message);
                        break;
                    default:
                        updateStatus(string.Format("{0} Status: {1}", e.State.ToString(), e.Message));
                        break;
                }
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

        private void btnSaveToFile_Click(object sender, RoutedEventArgs e)
        {
            Entity entityBeingSaved;
            if (radRecons.IsChecked ?? false)
                entityBeingSaved = Entity.Recons;
            else
                entityBeingSaved = Entity.Sources;
            var entityText = entityBeingSaved.ToString();

            var fileSaveDialog = getFileSaveDialog(string.Format("Save {0} to file", entityText));
            fileSaveDialog.Title = entityText;
            var xmlFile = fileSaveDialog.ShowDialog();

            try
            {
                if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
                {
                    if (entityBeingSaved == Entity.Recons)
                    {
                        rrController.WriteReconsToXmlFile(fileSaveDialog.FileName);
                    }
                    else
                    {
                        rrController.WriteSourcesToXmlFile(fileSaveDialog.FileName);
                    }
                    updateStatus(string.Format("Wrote {0} to file {1}", entityText, fileSaveDialog.FileName));
                }
                else
                    updateStatus(string.Format("Creation of {0} XML file cancelled.", entityBeingSaved));
                //checkReadyToRun();
            }
            catch (Exception ex)
            {
                updateStatus(string.Format("Unable to save {0} changes to file {1}: {2}", entityText, xmlFile, rrController.GetFullErrorMessage(ex)));
            }
        }


        private void cmbChooseRecon_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (rrPropertyGrid != null)
            {
                if (cmbChooseItem.Items.Count == 0 || cmbChooseItem.SelectedIndex == 0)
                {
                    rrPropertyGrid.Visibility = Visibility.Hidden;
                    enableEditingControls(Entity.None);
                }
                else
                {
                    rrPropertyGrid.Visibility = Visibility.Visible;
                    if (radRecons.IsChecked ?? false)
                    {
                        rrPropertyGrid.SelectedObject = rrController.Recons.ReconReports.Find(rr => rr.Name == cmbChooseItem.SelectedItem.ToString());
                        enableEditingControls(Entity.Recons);
                    }
                    else
                    {
                        rrPropertyGrid.SelectedObject = rrController.Sources;
                        enableEditingControls(Entity.Sources);
                    }
                }
            }
        }

        private void enableEditingControls(Entity entitySelected)
        {
            btnSaveToFile.IsEnabled = false;
            btnDeleteItem.IsEnabled = false;
            btnCopyItem.IsEnabled = false;
            switch (entitySelected)
            {
                case Entity.Recons:
                    btnSaveToFile.IsEnabled = true;
                    btnDeleteItem.IsEnabled = true;
                    btnCopyItem.IsEnabled = true;
                    break;
                case Entity.Sources:
                    btnSaveToFile.IsEnabled = true;
                    break;
            }
        }

        private void btnDeleteItem_Click(object sender, RoutedEventArgs e)
        {
        }

        private void btnCopyItem_Click(object sender, RoutedEventArgs e)
        {

        }

        private void radEntity_Checked(object sender, RoutedEventArgs e)
        {
            Entity entityToLoad;
            if (radRecons.IsChecked ?? false)
                entityToLoad = Entity.Recons;
            else
                entityToLoad = Entity.Sources;
            loadComboBox(entityToLoad);
        }

        private void rrPropertyGrid_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            rrController.Sources = rrController.Sources;
        }
    }
}
