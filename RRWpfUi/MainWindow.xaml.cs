using System.Windows;
using ReconRunner.Controller;
using System.Windows.Forms;


namespace RRWpfUi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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
                txtStatus.Text += "Reading sources from file " + xmlFileOpenDialog.FileName + ".\r\n\r\n";
                txtStatus.Text += rrController.ReadSourcesFromXMLFile(xmlFileOpenDialog.FileName);
                txtStatus.Text += "\r\n\r\n";
                btnRunRecons.IsEnabled = rrController.ReadyToRun();
            }
            else
            {
                txtStatus.Text += "Creation of Sources collection cancelled.\r\n\r\n";
                btnRunRecons.IsEnabled = false;
            }
        }

        private void btnLoadRecons_Click(object sender, RoutedEventArgs e)
        {
            var xmlFileOpenDialog = getFileOpenDialog("Choose Recons XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += "Reading recons from file " + xmlFileOpenDialog.FileName + ".\r\n\r\n";
                txtStatus.Text += rrController.ReadReconsFromXMLFile(xmlFileOpenDialog.FileName);
                txtStatus.Text += "\r\n\r\n";
                btnRunRecons.IsEnabled = rrController.ReadyToRun();
            }
            else
            {
                txtStatus.Text += "Creation of Recon collection cancelled.\r\n\r\n";
                btnRunRecons.IsEnabled = false;
            }

        }

        private void btnRunRecons_Click(object sender, RoutedEventArgs e)
        {
            var fileSaveDialog = getFileSaveDialog("Save recon results to spreadsheet");
            fileSaveDialog.DefaultExt = "xls";
            fileSaveDialog.Filter = "XLS Files|*.xls|All files|*.*";
            DialogResult excelFile = fileSaveDialog.ShowDialog();

            if (excelFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += "Processing recon reports and creating " + fileSaveDialog.FileName + ".\r\n\r\n";
                string results = rrController.RunRecons(fileSaveDialog.FileName);
                txtStatus.Text += results + "\r\n\r\n";
            }
            else
            {
                txtStatus.Text += "Creation of Recon report Excel file cancelled.\r\n\r\n";
            }
        }

        private void btnCreateSamples_Click(object sender, RoutedEventArgs e)
        {
            // Save sources
            var fileSaveDialog = getFileSaveDialog("Save sources to file");
            fileSaveDialog.Title = "Save Sources";
            DialogResult xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += "Writing sources to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.CreateSampleXMLSourcesFile(fileSaveDialog.FileName);
                txtStatus.Text += "Sources XML file created.\r\n\r\n";
            }
            else
            {
                txtStatus.Text += "Creation of Sources XML file cancelled.\r\n\r\n";
            }

            // Save recons
            fileSaveDialog = getFileSaveDialog("Save recons to file");
            fileSaveDialog.Title = "Save Recons";
            xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != System.Windows.Forms.DialogResult.Cancel)
            {
                txtStatus.Text += "Writing recons to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.CreateSampleXMLReconReportFile(fileSaveDialog.FileName);
                txtStatus.Text += "Recons XML file created.\r\n\r\n";
            }
            else
            {
                txtStatus.Text += "Creation of Recons XML file cancelled.\r\n\r\n";
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
