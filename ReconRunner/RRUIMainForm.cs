using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ReconRunner.Controller;

namespace ReconRunner
{
    public partial class RRUIMainForm : Form
    {
        private RRController rrController = RRController.Instance;
        
        public RRUIMainForm()
        {
            InitializeComponent();
        }

        private void createReconXMLFile_Click(object sender, EventArgs e)
        {
            // Save sources
            var fileSaveDialog = getFileSaveDialog("Save sources to file");
            fileSaveDialog.Title = "Save Sources";
            DialogResult xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Writing sources to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.WriteSourcesToXMLFile(fileSaveDialog.FileName);
                statusText.Text += "Sources XML file created.\r\n\r\n";
            }
            else
            {
                statusText.Text += "Creation of Sources XML file cancelled.\r\n\r\n";
            }

            // Save recons
            fileSaveDialog = getFileSaveDialog("Save recons to file");
            fileSaveDialog.Title = "Save Recons";
            xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Writing recons to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.WriteReconsToXMLFile(fileSaveDialog.FileName);
                statusText.Text += "Recons XML file created.\r\n\r\n";
            }
            else
            {
                statusText.Text += "Creation of Recons XML file cancelled.\r\n\r\n";
            }
        }

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

        private void readFromSourcesXMLFile_Click(object sender, EventArgs e)
        {
            xmlFileOpenDialog = getFileOpenDialog("Choose Sources XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Reading sources from file " + xmlFileOpenDialog.FileName + ".\r\n\r\n";
                statusText.Text += rrController.ReadSourcesFromXMLFile(xmlFileOpenDialog.FileName);
                statusText.Text += "\r\n\r\n";
                runRecons.Enabled = rrController.ReadyToRun();
            }
            else
            {
                statusText.Text += "Creation of Sources collection cancelled.\r\n\r\n";
                runRecons.Enabled = false;
            }
        }

        private void readFromReconsXMLFile_Click(object sender, EventArgs e)
        {
            xmlFileOpenDialog = getFileOpenDialog("Choose Recons XML file");
            DialogResult xmlFile = xmlFileOpenDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Reading recons from file " + xmlFileOpenDialog.FileName + ".\r\n\r\n";
                statusText.Text += rrController.ReadReconsFromXMLFile(xmlFileOpenDialog.FileName);
                statusText.Text += "\r\n\r\n";
                runRecons.Enabled = rrController.ReadyToRun();
            }
            else
            {
                statusText.Text += "Creation of Recon collection cancelled.\r\n\r\n";
                runRecons.Enabled = false;
            }

        }

        private void createSampleReconXMLFile_Click(object sender, EventArgs e)
        {
            var fileSaveDialog = getFileSaveDialog("Sample Sources XML");
            DialogResult xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Writing sample sources XML to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.CreateSampleXMLSourcesFile(fileSaveDialog.FileName);
                statusText.Text += "Sources XML file created.\r\n\r\n";
            }
            else
            {
                statusText.Text += "Creation of Sources XML file cancelled.\r\n\r\n";
            }

            fileSaveDialog = getFileSaveDialog("Sample Recons XML");
            xmlFile = fileSaveDialog.ShowDialog();

            if (xmlFile != DialogResult.Cancel)
            {
                statusText.Text += "Writing sample recons XML to file " + fileSaveDialog.FileName + ".\r\n\r\n";
                rrController.CreateSampleXMLReconReportFile(fileSaveDialog.FileName);
                statusText.Text += "Recons XML file created.\r\n\r\n";
            }
            else
            {
                statusText.Text += "Creation of Recon XML file cancelled.\r\n\r\n";
            }
        }

        private void runRecons_Click(object sender, EventArgs e)
        {
            fileSaveDialog.DefaultExt = "xls";
            fileSaveDialog.Filter = "XLS Files|*.xls|All files|*.*";
            DialogResult excelFile = fileSaveDialog.ShowDialog();

            if (excelFile != DialogResult.Cancel)
            {
                statusText.Text += "Processing recon reports and creating " + fileSaveDialog.FileName + ".\r\n\r\n";
                string results = rrController.RunRecons(fileSaveDialog.FileName);
                statusText.Text += results + "\r\n\r\n";
            }
            else
            {
                statusText.Text += "Creation of Recon report Excel file cancelled.\r\n\r\n";
            }
        }

        private void RRUIMainForm_Load(object sender, EventArgs e)
        {

        }

     }
}